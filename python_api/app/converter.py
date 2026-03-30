from __future__ import annotations

import io
import re
import shutil
import socket
import subprocess
import zipfile
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Iterable
from urllib.parse import urlparse

import httpx

from .config import settings

SUPPORTED_EXTENSIONS = {".ppt", ".pptx"}
UNO_EXPORT_SCRIPT = Path(__file__).with_name("libreoffice_pdf_export.py")


class ConversionError(RuntimeError):
    pass


def get_soffice_binary() -> Path:
    candidate = settings.libreoffice_program_dir / "soffice"
    if candidate.exists():
        return candidate

    fallback = shutil.which("soffice")
    if fallback:
        return Path(fallback)

    raise ConversionError("Missing LibreOffice soffice executable.")


def get_uno_python_binary() -> Path:
    bundled_python = settings.libreoffice_program_dir / "python"
    if bundled_python.exists():
        return bundled_python

    system_python = Path("/usr/bin/python3")
    if system_python.exists():
        return system_python

    raise ConversionError("Missing LibreOffice UNO python runtime.")


def ensure_dependencies() -> None:
    missing = [command for command in ("pdfinfo", "pdftocairo") if shutil.which(command) is None]
    if missing:
        raise ConversionError(f"Missing system dependencies: {', '.join(missing)}")
    get_soffice_binary()
    get_uno_python_binary()
    if not UNO_EXPORT_SCRIPT.exists():
        raise ConversionError("Missing LibreOffice export helper script.")


def convert_ppt_url_to_svg_zip(source_url: str) -> tuple[str, bytes]:
    ensure_dependencies()
    settings.work_root.mkdir(parents=True, exist_ok=True)

    with TemporaryDirectory(dir=settings.work_root) as temp_dir:
        temp_path = Path(temp_dir)
        input_path = download_presentation(source_url, temp_path)
        pdf_dir = temp_path / "pdf"
        svg_dir = temp_path / "svg"
        office_profile_dir = temp_path / "lo-profile"
        pdf_dir.mkdir(parents=True, exist_ok=True)
        svg_dir.mkdir(parents=True, exist_ok=True)
        office_profile_dir.mkdir(parents=True, exist_ok=True)

        pdf_path = convert_to_pdf(input_path, pdf_dir, office_profile_dir)
        svg_files = convert_pdf_to_svgs(pdf_path, svg_dir)
        archive_name = f"{input_path.stem}.zip"
        archive_bytes = build_zip_bytes(svg_files)
        return archive_name, archive_bytes


def download_presentation(source_url: str, temp_dir: Path) -> Path:
    parsed = urlparse(source_url)
    if parsed.scheme not in {"http", "https"}:
        raise ConversionError("Only http and https URLs are supported.")

    guessed_name = Path(parsed.path).name or "source.pptx"
    extension = Path(guessed_name).suffix.lower()
    if extension not in SUPPORTED_EXTENSIONS:
        extension = ".pptx"
    safe_name = sanitize_filename(Path(guessed_name).stem) + extension
    target_path = temp_dir / safe_name

    downloaded_bytes = 0
    timeout = httpx.Timeout(settings.download_timeout_seconds)
    headers = {"User-Agent": "ppt-to-svg-api/1.0"}
    try:
        with httpx.stream(
            "GET",
            source_url,
            follow_redirects=True,
            timeout=timeout,
            headers=headers,
        ) as response:
            response.raise_for_status()
            with target_path.open("wb") as file_obj:
                for chunk in response.iter_bytes():
                    if not chunk:
                        continue
                    downloaded_bytes += len(chunk)
                    if downloaded_bytes > settings.max_download_bytes:
                        raise ConversionError(
                            f"Downloaded file exceeds {settings.max_download_mb} MB."
                        )
                    file_obj.write(chunk)
    except httpx.HTTPError as exc:
        raise ConversionError(f"Failed to download source file: {exc}") from exc

    if downloaded_bytes == 0:
        raise ConversionError("Downloaded file is empty.")

    return target_path


def convert_to_pdf(input_path: Path, output_dir: Path, office_profile_dir: Path) -> Path:
    soffice_binary = get_soffice_binary()
    uno_python = get_uno_python_binary()
    pdf_path = output_dir / f"{input_path.stem}.pdf"
    listener_port = reserve_tcp_port()
    listener_log_path = output_dir / "libreoffice-listener.log"

    with listener_log_path.open("w+", encoding="utf-8") as log_file:
        listener_process = subprocess.Popen(
            [
                str(soffice_binary),
                f"-env:UserInstallation={office_profile_dir.resolve().as_uri()}",
                "--headless",
                "--nologo",
                "--nodefault",
                "--nofirststartwizard",
                "--norestore",
                "--invisible",
                f"--accept=socket,host=127.0.0.1,port={listener_port};urp;",
            ],
            stdout=log_file,
            stderr=log_file,
            text=True,
        )
        try:
            result = run_command(
                [
                    str(uno_python),
                    str(UNO_EXPORT_SCRIPT),
                    "--host",
                    "127.0.0.1",
                    "--port",
                    str(listener_port),
                    "--input",
                    str(input_path),
                    "--output",
                    str(pdf_path),
                    "--timeout",
                    str(settings.libreoffice_start_timeout_seconds),
                ]
            )
        except ConversionError as exc:
            listener_output = read_text_file(listener_log_path)
            if listener_output:
                raise ConversionError(
                    f"{exc} LibreOffice listener output: {listener_output}"
                ) from exc
            raise
        finally:
            terminate_process(listener_process)

    if not pdf_path.exists():
        details = (
            result.stdout.strip()
            if result.stdout
            else "LibreOffice did not generate a PDF file."
        )
        raise ConversionError(details)
    return pdf_path


def convert_pdf_to_svgs(pdf_path: Path, output_dir: Path) -> list[Path]:
    page_count = get_pdf_page_count(pdf_path)
    max_workers = max(1, min(settings.page_convert_workers, page_count))

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        svg_files = list(
            executor.map(
                lambda page: convert_pdf_page_to_svg(pdf_path, output_dir, page),
                range(1, page_count + 1),
            )
        )

    return sorted(svg_files)


def convert_pdf_page_to_svg(pdf_path: Path, output_dir: Path, page: int) -> Path:
    svg_path = output_dir / f"slide-{page:03d}.svg"
    run_command(
        [
            "pdftocairo",
            "-svg",
            "-f",
            str(page),
            "-l",
            str(page),
            str(pdf_path),
            str(svg_path),
        ]
    )
    if not svg_path.exists():
        raise ConversionError(f"Failed to generate SVG for page {page}.")
    return svg_path


def get_pdf_page_count(pdf_path: Path) -> int:
    result = run_command(["pdfinfo", str(pdf_path)])
    match = re.search(r"^Pages:\s+(\d+)$", result.stdout, re.MULTILINE)
    if match is None:
        raise ConversionError("Unable to determine PDF page count.")

    page_count = int(match.group(1))
    if page_count <= 0:
        raise ConversionError("PDF contains no pages.")
    return page_count


def build_zip_bytes(svg_files: Iterable[Path]) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for index, svg_file in enumerate(svg_files, start=1):
            archive.write(svg_file, arcname=f"slide-{index:03d}.svg")
    buffer.seek(0)
    return buffer.getvalue()


def run_command(command: list[str]) -> subprocess.CompletedProcess[str]:
    try:
        result = subprocess.run(
            command,
            check=True,
            capture_output=True,
            text=True,
            timeout=settings.command_timeout_seconds,
        )
    except subprocess.TimeoutExpired as exc:
        raise ConversionError(f"Command timed out: {' '.join(command)}") from exc
    except subprocess.CalledProcessError as exc:
        stderr = (exc.stderr or "").strip()
        stdout = (exc.stdout or "").strip()
        details = stderr or stdout or "No command output."
        raise ConversionError(f"Command failed: {' '.join(command)}. {details}") from exc
    return result


def reserve_tcp_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        sock.listen(1)
        return sock.getsockname()[1]


def terminate_process(process: subprocess.Popen[str]) -> None:
    if process.poll() is not None:
        return
    process.terminate()
    try:
        process.wait(timeout=10)
    except subprocess.TimeoutExpired:
        process.kill()
        process.wait(timeout=5)


def read_text_file(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8").strip()
    except FileNotFoundError:
        return ""


def sanitize_filename(name: str) -> str:
    normalized = re.sub(r"[^A-Za-z0-9._-]+", "-", name).strip("-._")
    return normalized or "source"
