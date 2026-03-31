from __future__ import annotations

import io
import re
import shutil
import subprocess
import zipfile
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Iterable
from urllib.parse import urlparse
from uuid import uuid4

import httpx

from .config import settings

SUPPORTED_EXTENSIONS = {".ppt", ".pptx"}


class ConversionError(RuntimeError):
    pass


def get_java_command() -> str:
    java_command = settings.java_command
    if Path(java_command).exists():
        return java_command

    fallback = shutil.which(java_command)
    if fallback:
        return fallback

    raise ConversionError(f"Missing Java runtime executable: {java_command}")


def ensure_dependencies() -> None:
    get_java_command()
    if not settings.java_renderer_jar.exists():
        raise ConversionError("Missing native SVG renderer helper jar.")


def convert_ppt_url_to_svg_zip(source_url: str) -> tuple[str, bytes]:
    ensure_dependencies()
    settings.work_root.mkdir(parents=True, exist_ok=True)

    with TemporaryDirectory(dir=settings.work_root) as temp_dir:
        temp_path = Path(temp_dir)
        input_path = download_presentation(source_url, temp_path)
        svg_dir = temp_path / "svg"
        svg_dir.mkdir(parents=True, exist_ok=True)

        svg_files = render_presentation_to_svgs(input_path, svg_dir)
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


def render_presentation_to_svgs(input_path: Path, output_dir: Path) -> list[Path]:
    run_command(
        [
            get_java_command(),
            "-jar",
            str(settings.java_renderer_jar),
            "--input",
            str(input_path),
            "--output-dir",
            str(output_dir),
            "--text-as-shapes",
            "true" if settings.svg_text_as_shapes else "false",
        ]
    )

    svg_files = sorted(output_dir.glob("slide-*.svg"))
    if not svg_files:
        raise ConversionError("Renderer did not generate any SVG slides.")
    return svg_files


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


def sanitize_filename(name: str) -> str:
    normalized = re.sub(r"[^A-Za-z0-9._-]+", "-", name).strip("-._")
    return normalized or "source"


def save_zip_bytes(archive_name: str, archive_bytes: bytes) -> tuple[str, Path]:
    settings.downloads_root.mkdir(parents=True, exist_ok=True)
    base_name = sanitize_filename(Path(archive_name).stem)
    unique_name = f"{base_name}-{uuid4().hex[:12]}.zip"
    target_path = settings.downloads_root / unique_name
    target_path.write_bytes(archive_bytes)
    return unique_name, target_path
