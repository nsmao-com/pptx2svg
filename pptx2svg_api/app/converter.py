# input: Remote PPT/PPTX URL plus LibreOffice and MuPDF commands.
# output: Per-slide SVG ZIP for legacy conversion, PNG page screenshots plus masked vision PNGs/PPTX-derived image metadata for beautify, and progressive page-ready PNG callbacks for pipelined recognition; beautify metadata includes normalized chart type/data and image usage hints.
# pos: Core conversion pipeline behind the PPT/PPTX->SVG FastAPI endpoint.
# update: When this file changes, update this header and pptx2svg_api/app/README.md.

from __future__ import annotations

import base64
import copy
import hashlib
import hmac
import io
import json
import mimetypes
import posixpath
import re
import shutil
import subprocess
import xml.etree.ElementTree as ET
import zipfile
from email.utils import formatdate
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Callable, Iterable, TypeAlias
from urllib.parse import quote, urlparse
from uuid import uuid4

import httpx
from PIL import Image, ImageDraw

from .config import settings

SUPPORTED_EXTENSIONS = {".ppt", ".pptx"}
OOO_NS = "http://xml.openoffice.org/svg/export"
XLINK_NS = "http://www.w3.org/1999/xlink"
PPT_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
STYLE_FONT_FAMILY_RE = re.compile(r"font-family\s*:\s*[^;]+")
OOXML_TYPEFACE_ATTR_RE = re.compile(r'(?<=\btypeface=")[^"]*')
DATA_IMAGE_RE = re.compile(r"^data:(image/[A-Za-z0-9.+-]+);base64,(.*)$", re.DOTALL)
PPTX_SLIDE_RE = re.compile(r"^ppt/slides/slide(\d+)\.xml$")
TRANSFORM_RE = re.compile(r"([a-zA-Z]+)\(([^)]*)\)")
DEFAULT_SLIDE_WIDTH_EMU = 12192000
DEFAULT_SLIDE_HEIGHT_EMU = 6858000
NORMALIZED_SLIDE_WIDTH = 1280.0
NORMALIZED_SLIDE_HEIGHT = 720.0
VISION_MASK_FILL = "#F1F5F9"
VISION_MASK_STROKE = "#CBD5E1"
SCREENSHOT_NAME_MARKERS = ("screenshot", "screen-shot", "screen_shot", "capture", "snip", "截屏", "截图", "屏幕截图")
SUPPORTED_CHART_TYPES = {"bar", "horizontal-bar", "line", "pie", "donut"}
ImageUploadCacheKey: TypeAlias = tuple[str, str]
ImageUploadCache: TypeAlias = dict[ImageUploadCacheKey, str]
CHART_TYPE_LABELS = {
    "bar": "柱形图",
    "horizontal-bar": "条形图",
    "line": "折线图",
    "pie": "饼图",
    "donut": "环形图",
    "unknown": "未知图表",
}
CHART_DEFAULT_PALETTE = [
    "#2563EB",
    "#94A3B8",
    "#22C55E",
    "#F59E0B",
    "#EF4444",
    "#8B5CF6",
    "#14B8A6",
    "#F97316",
]


class ConversionError(RuntimeError):
    pass


def get_required_command(command_name: str) -> str:
    if Path(command_name).exists():
        return command_name

    fallback = shutil.which(command_name)
    if fallback:
        return fallback

    raise ConversionError(f"Missing required executable: {command_name}")


def ensure_dependencies() -> None:
    get_required_command(settings.libreoffice_command)
    get_required_command(settings.mupdf_command)


def convert_ppt_url_to_svg_zip(source_url: str) -> tuple[str, bytes]:
    ensure_dependencies()
    settings.work_root.mkdir(parents=True, exist_ok=True)

    with TemporaryDirectory(dir=settings.work_root) as temp_dir:
        temp_path = Path(temp_dir)
        input_path = download_presentation(source_url, temp_path)
        svg_dir = temp_path / "svg"
        svg_dir.mkdir(parents=True, exist_ok=True)

        asset_prefix = build_asset_prefix(input_path)
        svg_files = render_presentation_to_svgs(input_path, svg_dir, asset_prefix)
        archive_name = f"{input_path.stem}.zip"
        archive_bytes = build_zip_bytes(svg_files)
        return archive_name, archive_bytes


def analyze_ppt_url_for_beautify(source_url: str) -> tuple[str, bytes, dict]:
    ensure_dependencies()
    settings.work_root.mkdir(parents=True, exist_ok=True)

    with TemporaryDirectory(dir=settings.work_root) as temp_dir:
        temp_path = Path(temp_dir)
        input_path = download_presentation(source_url, temp_path)
        png_dir = temp_path / "png"
        png_dir.mkdir(parents=True, exist_ok=True)
        asset_dir = temp_path / "assets"
        asset_dir.mkdir(parents=True, exist_ok=True)

        png_files = render_presentation_to_pngs(input_path, png_dir, temp_path)
        metadata, asset_files = build_beautify_png_metadata(input_path, temp_path, png_files, asset_dir)
        archive_name = f"{input_path.stem}-beautify.zip"
        archive_bytes = build_beautify_png_zip_bytes(png_files, asset_files, metadata)
        return archive_name, archive_bytes, metadata


def analyze_ppt_url_for_beautify_progressive(
    source_url: str,
    work_dir: Path,
    on_page_ready: Callable[[dict, Path], None] | None = None,
) -> tuple[str, bytes, dict, list[Path], list[tuple[Path, str]]]:
    ensure_dependencies()
    work_dir.mkdir(parents=True, exist_ok=True)

    input_path = download_presentation(source_url, work_dir)
    png_dir = work_dir / "png"
    png_dir.mkdir(parents=True, exist_ok=True)
    asset_dir = work_dir / "assets"
    asset_dir.mkdir(parents=True, exist_ok=True)

    profile_dir = work_dir / "libreoffice-png-profile"
    profile_dir.mkdir(parents=True, exist_ok=True)
    pdf_path = export_presentation_to_pdf(input_path, work_dir, profile_dir)
    total_pages = read_pdf_page_count(pdf_path)
    if total_pages <= 0:
        raise ConversionError("PDF page count is empty.")

    slide_texts, chart_map, image_map, asset_files = collect_beautify_source_metadata(
        input_path,
        work_dir,
        asset_dir,
        total_pages,
    )

    png_files: list[Path] = []
    slides: list[dict] = []
    failed_pages: list[int] = []
    for page_number in range(1, total_pages + 1):
        png_file = render_pdf_page_to_png(pdf_path, png_dir, page_number)
        png_files.append(png_file)
        slide_metadata, failed = build_beautify_png_slide_metadata(
            page_number,
            png_file,
            slide_texts,
            image_map,
            chart_map,
        )
        attach_beautify_vision_image_metadata(page_number, png_file, slide_metadata)
        if failed:
            failed_pages.append(page_number)
        slides.append(slide_metadata)
        if on_page_ready is not None:
            page_ready_metadata = copy.deepcopy(slide_metadata)
            page_ready_metadata["total_pages"] = total_pages
            on_page_ready(page_ready_metadata, png_file)

    metadata = {
        "source_filename": input_path.name,
        "total_pages": total_pages,
        "normalized_width": int(NORMALIZED_SLIDE_WIDTH),
        "normalized_height": int(NORMALIZED_SLIDE_HEIGHT),
        "render_mode": "png_vision",
        "failed_pages": failed_pages,
        "slides": slides,
    }
    archive_name = f"{input_path.stem}-beautify.zip"
    archive_bytes = build_beautify_png_zip_bytes(png_files, asset_files, metadata)
    return archive_name, archive_bytes, metadata, png_files, asset_files


def extract_ppt_url_to_text(source_url: str) -> tuple[str, bytes]:
    settings.work_root.mkdir(parents=True, exist_ok=True)

    with TemporaryDirectory(dir=settings.work_root) as temp_dir:
        temp_path = Path(temp_dir)
        input_path = download_presentation(source_url, temp_path)
        text_input = prepare_presentation_for_text_extraction(input_path, temp_path)
        slide_texts = extract_pptx_slide_texts(text_input)
        text_content = build_slide_text_document(slide_texts)
        text_name = f"{input_path.stem}.txt"
        return text_name, text_content.encode("utf-8")


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


def render_presentation_to_svgs(input_path: Path, output_dir: Path, asset_prefix: str) -> list[Path]:
    working_dir = output_dir.parent
    profile_dir = working_dir / "libreoffice-profile"
    profile_dir.mkdir(parents=True, exist_ok=True)

    try:
        text_svg_input = prepare_presentation_for_text_svg(input_path, working_dir)
        full_svg = export_presentation_to_svg(text_svg_input, working_dir, profile_dir)
        final_svg_files = split_libreoffice_svg(full_svg, output_dir, asset_prefix)
    except ConversionError:
        if settings.aliyun_oss_enabled:
            raise
        final_svg_files = render_presentation_to_outline_svgs(input_path, output_dir, working_dir, profile_dir)

    if not final_svg_files:
        raise ConversionError("Renderer did not generate any SVG slides.")
    return final_svg_files


def render_presentation_to_outline_svgs(
    input_path: Path,
    output_dir: Path,
    working_dir: Path,
    profile_dir: Path,
) -> list[Path]:
    pdf_path = export_presentation_to_pdf(input_path, working_dir, profile_dir)
    pdf_svg_dir = working_dir / "mupdf-svg"
    pdf_svg_dir.mkdir(parents=True, exist_ok=True)
    pdf_svg_files = render_pdf_to_svgs(pdf_path, pdf_svg_dir)

    final_svg_files: list[Path] = []
    for page_number, source_svg in enumerate(pdf_svg_files, start=1):
        target_svg = output_dir / format_slide_name(page_number)
        shutil.copy2(source_svg, target_svg)
        final_svg_files.append(target_svg)

    return final_svg_files


def render_presentation_to_pngs(input_path: Path, output_dir: Path, working_dir: Path) -> list[Path]:
    profile_dir = working_dir / "libreoffice-png-profile"
    profile_dir.mkdir(parents=True, exist_ok=True)
    pdf_path = export_presentation_to_pdf(input_path, working_dir, profile_dir)
    png_files = render_pdf_to_pngs(pdf_path, output_dir)
    if not png_files:
        raise ConversionError("Renderer did not generate any PNG slides.")
    return png_files


def prepare_presentation_for_text_svg(input_path: Path, working_dir: Path) -> Path:
    if input_path.suffix.lower() != ".pptx":
        return input_path

    target_path = working_dir / f"{input_path.stem}-fixed-font{input_path.suffix}"
    fixed_font = settings.svg_fixed_font_family

    try:
        with zipfile.ZipFile(input_path, "r") as source_archive:
            with zipfile.ZipFile(target_path, "w", compression=zipfile.ZIP_DEFLATED) as target_archive:
                for item in source_archive.infolist():
                    data = source_archive.read(item.filename)
                    if item.filename.endswith(".xml"):
                        data = replace_ooxml_font_names(data, fixed_font)
                    target_archive.writestr(item, data)
    except zipfile.BadZipFile as exc:
        raise ConversionError("PPTX file is not a valid zip archive.") from exc

    return target_path


def replace_ooxml_font_names(data: bytes, fixed_font: str) -> bytes:
    try:
        text = data.decode("utf-8")
    except UnicodeDecodeError:
        return data

    updated_text = OOXML_TYPEFACE_ATTR_RE.sub(fixed_font, text)
    return updated_text.encode("utf-8")


def prepare_presentation_for_text_extraction(input_path: Path, working_dir: Path) -> Path:
    if input_path.suffix.lower() == ".pptx":
        return input_path

    if input_path.suffix.lower() != ".ppt":
        raise ConversionError("Only .ppt and .pptx files are supported for text extraction.")

    profile_dir = working_dir / "libreoffice-text-profile"
    profile_dir.mkdir(parents=True, exist_ok=True)
    export_dir = working_dir / "pptx-export"
    export_dir.mkdir(parents=True, exist_ok=True)
    run_command(
        build_libreoffice_command(
            profile_dir,
            [
                "--convert-to",
                "pptx",
                "--outdir",
                str(export_dir),
                str(input_path),
            ],
        )
    )
    return collect_single_output_file(export_dir, ".pptx")


def extract_pptx_slide_texts(input_path: Path) -> list[list[str]]:
    try:
        with zipfile.ZipFile(input_path, "r") as archive:
            slide_entries = collect_ordered_slide_entries(archive)
            slide_texts = [
                extract_slide_text_lines(archive.read(slide_entry))
                for slide_entry in slide_entries
            ]
    except zipfile.BadZipFile as exc:
        raise ConversionError("PPTX file is not a valid zip archive.") from exc

    if not slide_texts:
        raise ConversionError("PPTX file does not contain any slides.")
    return slide_texts


def collect_ordered_slide_entries(archive: zipfile.ZipFile) -> list[str]:
    ordered_entries = collect_presentation_ordered_slide_entries(archive)
    if ordered_entries:
        return ordered_entries

    numbered_entries: list[tuple[int, str]] = []
    for name in archive.namelist():
        match = PPTX_SLIDE_RE.match(name)
        if match:
            numbered_entries.append((int(match.group(1)), name))
    return [name for _, name in sorted(numbered_entries)]


def collect_presentation_ordered_slide_entries(archive: zipfile.ZipFile) -> list[str]:
    try:
        presentation = ET.fromstring(archive.read("ppt/presentation.xml"))
        relationships = ET.fromstring(archive.read("ppt/_rels/presentation.xml.rels"))
    except (KeyError, ET.ParseError):
        return []

    relationship_targets: dict[str, str] = {}
    for relationship in relationships:
        if local_name(relationship.tag) != "Relationship":
            continue
        relationship_id = relationship.attrib.get("Id", "")
        target = relationship.attrib.get("Target", "")
        if relationship_id and target:
            relationship_targets[relationship_id] = normalize_pptx_part_path("ppt", target)

    ordered_entries: list[str] = []
    for slide_id in presentation.iter(f"{{{PPT_NS}}}sldId"):
        relationship_id = slide_id.attrib.get(f"{{{REL_NS}}}id", "")
        target = relationship_targets.get(relationship_id)
        if target and target in archive.namelist():
            ordered_entries.append(target)

    return ordered_entries


def normalize_pptx_part_path(base_dir: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(base_dir, target))


def extract_slide_text_lines(slide_xml: bytes) -> list[str]:
    try:
        root = ET.fromstring(slide_xml)
    except ET.ParseError as exc:
        raise ConversionError(f"Slide XML is not valid: {exc}") from exc

    lines: list[str] = []
    for paragraph in root.iter(f"{{{DRAWING_NS}}}p"):
        chunks = [
            text_node.text or ""
            for text_node in paragraph.iter(f"{{{DRAWING_NS}}}t")
        ]
        line = normalize_text_line("".join(chunks))
        if line:
            lines.append(line)
    return lines


def normalize_text_line(text: str) -> str:
    return re.sub(r"[ \t\r\f\v]+", " ", text).strip()


def build_slide_text_document(slide_texts: list[list[str]]) -> str:
    sections: list[str] = []
    for index, lines in enumerate(slide_texts, start=1):
        header = f"#第{format_chinese_number(index)}页"
        body = "\n".join(lines) if lines else "（本页无文字）"
        sections.append(f"{header}\n{body}")
    return "\n\n".join(sections) + "\n"


def build_beautify_metadata(input_path: Path, working_dir: Path, svg_files: list[Path]) -> dict:
    text_input: Path | None = None
    slide_texts: list[list[str]] = []
    chart_map: dict[int, list[dict]] = {}

    try:
        text_input = prepare_presentation_for_text_extraction(input_path, working_dir)
        slide_texts = extract_pptx_slide_texts(text_input)
    except ConversionError:
        slide_texts = [[] for _ in svg_files]

    if text_input is not None and text_input.suffix.lower() == ".pptx":
        try:
            chart_map = extract_pptx_chart_metadata(text_input)
        except ConversionError:
            chart_map = {}

    slides: list[dict] = []
    failed_pages: list[int] = []
    for page_number, svg_file in enumerate(svg_files, start=1):
        page_metadata = {
            "index": page_number,
            "svg_filename": format_slide_name(page_number),
            "text": slide_texts[page_number - 1] if page_number <= len(slide_texts) else [],
            "images": [],
            "charts": copy.deepcopy(chart_map.get(page_number, [])),
            "fallback_as_image": False,
            "parse_error": "",
        }
        try:
            svg_info = analyze_slide_svg(svg_file, page_number)
            page_metadata["images"] = svg_info["images"]
            page_metadata["charts"] = attach_chart_fallback_images(page_metadata["charts"], page_metadata["images"])
            if svg_info["fallback_as_image"]:
                page_metadata["fallback_as_image"] = True
        except ConversionError as exc:
            page_metadata["fallback_as_image"] = True
            page_metadata["parse_error"] = str(exc)
            failed_pages.append(page_number)
        slides.append(page_metadata)

    return {
        "source_filename": input_path.name,
        "total_pages": len(svg_files),
        "normalized_width": int(NORMALIZED_SLIDE_WIDTH),
        "normalized_height": int(NORMALIZED_SLIDE_HEIGHT),
        "failed_pages": failed_pages,
        "slides": slides,
    }


def build_beautify_png_metadata(
    input_path: Path,
    working_dir: Path,
    png_files: list[Path],
    asset_dir: Path,
) -> tuple[dict, list[tuple[Path, str]]]:
    slide_texts, chart_map, image_map, asset_files = collect_beautify_source_metadata(
        input_path,
        working_dir,
        asset_dir,
        len(png_files),
    )

    slides: list[dict] = []
    failed_pages: list[int] = []
    for page_number, png_file in enumerate(png_files, start=1):
        slide_metadata, failed = build_beautify_png_slide_metadata(
            page_number,
            png_file,
            slide_texts,
            image_map,
            chart_map,
        )
        attach_beautify_vision_image_metadata(page_number, png_file, slide_metadata)
        if failed:
            failed_pages.append(page_number)
        slides.append(slide_metadata)

    return (
        {
            "source_filename": input_path.name,
            "total_pages": len(png_files),
            "normalized_width": int(NORMALIZED_SLIDE_WIDTH),
            "normalized_height": int(NORMALIZED_SLIDE_HEIGHT),
            "render_mode": "png_vision",
            "failed_pages": failed_pages,
            "slides": slides,
        },
        asset_files,
    )


def collect_beautify_source_metadata(
    input_path: Path,
    working_dir: Path,
    asset_dir: Path,
    page_count: int,
) -> tuple[list[list[str]], dict[int, list[dict]], dict[int, list[dict]], list[tuple[Path, str]]]:
    text_input: Path | None = None
    slide_texts: list[list[str]] = []
    chart_map: dict[int, list[dict]] = {}
    image_map: dict[int, list[dict]] = {}
    asset_files: list[tuple[Path, str]] = []

    try:
        text_input = prepare_presentation_for_text_extraction(input_path, working_dir)
        slide_texts = extract_pptx_slide_texts(text_input)
    except ConversionError:
        slide_texts = [[] for _ in range(page_count)]

    if text_input is not None and text_input.suffix.lower() == ".pptx":
        try:
            chart_map = extract_pptx_chart_metadata(text_input)
        except ConversionError:
            chart_map = {}
        try:
            image_map, asset_files = extract_pptx_image_metadata(text_input, asset_dir)
        except ConversionError:
            image_map = {}
            asset_files = []

    return slide_texts, chart_map, image_map, asset_files


def build_beautify_png_slide_metadata(
    page_number: int,
    png_file: Path,
    slide_texts: list[list[str]],
    image_map: dict[int, list[dict]],
    chart_map: dict[int, list[dict]],
) -> tuple[dict, bool]:
    width, height = read_image_file_dimensions(png_file)
    failed = width <= 0 or height <= 0
    images = copy.deepcopy(image_map.get(page_number, []))
    charts = attach_chart_fallback_images(copy.deepcopy(chart_map.get(page_number, [])), images)
    return (
        {
            "index": page_number,
            "page_key": f"origin-{page_number:03d}",
            "image_filename": format_slide_png_name(page_number),
            "image_width": width,
            "image_height": height,
            "image_mime": "image/png",
            "text": slide_texts[page_number - 1] if page_number <= len(slide_texts) else [],
            "images": images,
            "charts": charts,
            "fallback_as_image": False,
            "parse_error": "" if not failed else "PNG page image is not readable",
        },
        failed,
    )


def read_image_file_dimensions(image_path: Path) -> tuple[int, int]:
    try:
        with Image.open(image_path) as image:
            return int(image.width), int(image.height)
    except Exception:
        return 0, 0


def extract_pptx_image_metadata(input_path: Path, asset_dir: Path) -> tuple[dict[int, list[dict]], list[tuple[Path, str]]]:
    images_by_slide: dict[int, list[dict]] = {}
    asset_files: list[tuple[Path, str]] = []
    try:
        with zipfile.ZipFile(input_path, "r") as archive:
            slide_width_emu, slide_height_emu = read_pptx_slide_size_emu(archive)
            slide_entries = collect_ordered_slide_entries(archive)
            for page_number, slide_entry in enumerate(slide_entries, start=1):
                images_by_slide[page_number] = extract_slide_images(
                    archive,
                    slide_entry,
                    page_number,
                    slide_width_emu,
                    slide_height_emu,
                    asset_dir,
                    asset_files,
                )
    except zipfile.BadZipFile as exc:
        raise ConversionError("PPTX file is not a valid zip archive.") from exc
    return images_by_slide, asset_files


def extract_slide_images(
    archive: zipfile.ZipFile,
    slide_entry: str,
    page_number: int,
    slide_width_emu: int,
    slide_height_emu: int,
    asset_dir: Path,
    asset_files: list[tuple[Path, str]],
) -> list[dict]:
    try:
        slide_root = ET.fromstring(archive.read(slide_entry))
    except (KeyError, ET.ParseError):
        return []

    relationships = read_relationship_items(archive, relationship_part_for(slide_entry))
    images: list[dict] = []
    image_index = 0
    for picture in iter_slide_image_containers(slide_root):
        rel_id = find_picture_relationship_id(picture)
        if not rel_id:
            continue
        rel = relationships.get(rel_id)
        if not rel:
            continue

        target = rel.get("target", "")
        target_mode = rel.get("target_mode", "")
        bounds = picture_bounds(picture, slide_width_emu, slide_height_emu)
        if bounds["width"] <= 0 or bounds["height"] <= 0:
            continue

        next_image_index = image_index + 1
        image_url = ""
        asset_filename = ""
        pixel_width = 0
        pixel_height = 0
        if target_mode.lower() == "external" or target.startswith(("http://", "https://", "data:")):
            image_url = target
        else:
            media_path = normalize_pptx_part_path(posixpath.dirname(slide_entry), target)
            if media_path not in archive.namelist():
                continue
            image_bytes = archive.read(media_path)
            pixel_width, pixel_height = read_image_bytes_dimensions(image_bytes)
            extension = detect_image_extension(image_bytes, media_path)
            if pixel_width <= 0 and Path(media_path).suffix.lower() not in {".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp"}:
                continue
            asset_filename = f"assets/slide-{page_number:03d}/image-{next_image_index:03d}{extension}"
            asset_path = asset_dir / f"slide-{page_number:03d}" / f"image-{next_image_index:03d}{extension}"
            asset_path.parent.mkdir(parents=True, exist_ok=True)
            asset_path.write_bytes(image_bytes)
            asset_files.append((asset_path, asset_filename))

        aspect_ratio = bounds["width"] / bounds["height"] if bounds["height"] > 0 else 0
        role = classify_svg_image_role(bounds["x"], bounds["y"], bounds["width"], bounds["height"], image_url or target)
        image_meta = {
            "id": f"image-{page_number:03d}-{next_image_index:03d}",
            "url": image_url,
            "filename": asset_filename,
            "x": bounds["x"],
            "y": bounds["y"],
            "width": bounds["width"],
            "height": bounds["height"],
            "aspect_ratio": round(aspect_ratio, 4),
            "role": role,
        }
        if should_ignore_image_ocr(role):
            image_meta["ocr_ignore"] = True
        if pixel_width > 0 and pixel_height > 0:
            image_meta["pixel_width"] = pixel_width
            image_meta["pixel_height"] = pixel_height
        image_index = next_image_index
        images.append(image_meta)
    return images


def iter_slide_image_containers(slide_root: ET.Element) -> list[ET.Element]:
    containers: list[ET.Element] = []
    seen: set[int] = set()
    for tag_name in ("bg", "pic", "sp"):
        for element in slide_root.iter(f"{{{PPT_NS}}}{tag_name}"):
            if find_picture_relationship_id(element):
                marker = id(element)
                if marker in seen:
                    continue
                seen.add(marker)
                containers.append(element)
    return containers


def read_relationship_items(archive: zipfile.ZipFile, rels_entry: str) -> dict[str, dict[str, str]]:
    try:
        root = ET.fromstring(archive.read(rels_entry))
    except (KeyError, ET.ParseError):
        return {}
    relationships: dict[str, dict[str, str]] = {}
    for relationship in root:
        if local_name(relationship.tag) != "Relationship":
            continue
        rel_id = relationship.attrib.get("Id", "")
        target = relationship.attrib.get("Target", "")
        if rel_id and target:
            relationships[rel_id] = {
                "target": target,
                "target_mode": relationship.attrib.get("TargetMode", ""),
            }
    return relationships


def find_picture_relationship_id(picture: ET.Element) -> str:
    for blip in picture.iter(f"{{{DRAWING_NS}}}blip"):
        rel_id = blip.attrib.get(f"{{{REL_NS}}}embed", "") or blip.attrib.get(f"{{{REL_NS}}}link", "")
        if rel_id:
            return rel_id
    return ""


def picture_bounds(picture: ET.Element, slide_width_emu: int, slide_height_emu: int) -> dict:
    if local_name(picture.tag) == "bg":
        return {"x": 0.0, "y": 0.0, "width": NORMALIZED_SLIDE_WIDTH, "height": NORMALIZED_SLIDE_HEIGHT}
    off = None
    ext = None
    for xfrm in picture.iter(f"{{{DRAWING_NS}}}xfrm"):
        off = xfrm.find(f"{{{DRAWING_NS}}}off")
        ext = xfrm.find(f"{{{DRAWING_NS}}}ext")
        break
    if off is None or ext is None:
        return {"x": 0.0, "y": 0.0, "width": 0.0, "height": 0.0}
    x = parse_int(off.attrib.get("x", ""), 0) / slide_width_emu * NORMALIZED_SLIDE_WIDTH
    y = parse_int(off.attrib.get("y", ""), 0) / slide_height_emu * NORMALIZED_SLIDE_HEIGHT
    width = parse_int(ext.attrib.get("cx", ""), 0) / slide_width_emu * NORMALIZED_SLIDE_WIDTH
    height = parse_int(ext.attrib.get("cy", ""), 0) / slide_height_emu * NORMALIZED_SLIDE_HEIGHT
    return {
        "x": round(clamp_number(x, 0, NORMALIZED_SLIDE_WIDTH), 2),
        "y": round(clamp_number(y, 0, NORMALIZED_SLIDE_HEIGHT), 2),
        "width": round(clamp_number(width, 0, NORMALIZED_SLIDE_WIDTH), 2),
        "height": round(clamp_number(height, 0, NORMALIZED_SLIDE_HEIGHT), 2),
    }


def read_image_bytes_dimensions(image_bytes: bytes) -> tuple[int, int]:
    try:
        with Image.open(io.BytesIO(image_bytes)) as image:
            return int(image.width), int(image.height)
    except Exception:
        return 0, 0


def detect_image_extension(image_bytes: bytes, source_path: str) -> str:
    source_ext = Path(source_path).suffix.lower()
    if source_ext in {".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp", ".tif", ".tiff"}:
        return ".jpg" if source_ext == ".jpeg" else source_ext
    try:
        with Image.open(io.BytesIO(image_bytes)) as image:
            fmt = (image.format or "").lower()
    except Exception:
        fmt = ""
    if fmt == "jpeg":
        return ".jpg"
    if fmt in {"png", "gif", "webp", "bmp", "tiff"}:
        return f".{fmt}"
    return ".png"


def analyze_slide_svg(svg_path: Path, page_number: int) -> dict:
    try:
        tree = ET.parse(svg_path)
    except ET.ParseError as exc:
        raise ConversionError(f"SVG parse failed: {exc}") from exc

    root = tree.getroot()
    width, height = get_svg_logical_size(root)
    if width <= 0 or height <= 0:
        width, height = NORMALIZED_SLIDE_WIDTH, NORMALIZED_SLIDE_HEIGHT
    scale_x = NORMALIZED_SLIDE_WIDTH / width
    scale_y = NORMALIZED_SLIDE_HEIGHT / height

    images: list[dict] = []
    fallback_as_image = False

    def walk(node: ET.Element, matrix: tuple[float, float, float, float, float, float]) -> None:
        nonlocal fallback_as_image
        next_matrix = multiply_matrix(matrix, parse_transform(node.attrib.get("transform", "")))
        if local_name(node.tag) == "image":
            image = image_metadata_from_svg_node(node, next_matrix, scale_x, scale_y, page_number, len(images) + 1)
            if image is not None:
                images.append(image)
                if image["role"] == "background" and image["width"] >= 1200 and image["height"] >= 660:
                    fallback_as_image = True
        for child in node:
            walk(child, next_matrix)

    walk(root, identity_matrix())
    return {
        "images": images,
        "fallback_as_image": fallback_as_image and len(images) == 1,
    }


def get_svg_logical_size(root: ET.Element) -> tuple[float, float]:
    view_box = root.attrib.get("viewBox") or root.attrib.get("viewbox")
    if view_box:
        parts = [parse_svg_number(part) for part in re.split(r"[,\s]+", view_box.strip()) if part.strip()]
        if len(parts) == 4 and parts[2] > 0 and parts[3] > 0:
            return parts[2], parts[3]
    width = parse_svg_number(root.attrib.get("width", ""))
    height = parse_svg_number(root.attrib.get("height", ""))
    return width, height


def image_metadata_from_svg_node(
    node: ET.Element,
    matrix: tuple[float, float, float, float, float, float],
    scale_x: float,
    scale_y: float,
    page_number: int,
    image_index: int,
) -> dict | None:
    href = node.attrib.get("href") or node.attrib.get(f"{{{XLINK_NS}}}href") or ""
    x = parse_svg_number(node.attrib.get("x", "0"))
    y = parse_svg_number(node.attrib.get("y", "0"))
    width = parse_svg_number(node.attrib.get("width", "0"))
    height = parse_svg_number(node.attrib.get("height", "0"))
    if width <= 0 or height <= 0:
        return None

    x1, y1 = apply_matrix(matrix, x, y)
    x2, y2 = apply_matrix(matrix, x + width, y + height)
    nx = clamp_number(min(x1, x2) * scale_x, 0, NORMALIZED_SLIDE_WIDTH)
    ny = clamp_number(min(y1, y2) * scale_y, 0, NORMALIZED_SLIDE_HEIGHT)
    nw = clamp_number(abs(x2 - x1) * scale_x, 0, NORMALIZED_SLIDE_WIDTH)
    nh = clamp_number(abs(y2 - y1) * scale_y, 0, NORMALIZED_SLIDE_HEIGHT)
    aspect_ratio = nw / nh if nh > 0 else 0

    role = classify_svg_image_role(nx, ny, nw, nh, href)
    image_meta = {
        "id": f"image-{page_number:03d}-{image_index:03d}",
        "url": href,
        "x": round(nx, 2),
        "y": round(ny, 2),
        "width": round(nw, 2),
        "height": round(nh, 2),
        "aspect_ratio": round(aspect_ratio, 4),
        "role": role,
    }
    if should_ignore_image_ocr(role):
        image_meta["ocr_ignore"] = True
    return image_meta


def classify_svg_image_role(x: float, y: float, width: float, height: float, href: str) -> str:
    href_lower = href.lower()
    if width >= NORMALIZED_SLIDE_WIDTH * 0.85 and height >= NORMALIZED_SLIDE_HEIGHT * 0.85:
        return "background"
    if "logo" in href_lower:
        return "logo"
    if width <= NORMALIZED_SLIDE_WIDTH * 0.2 and height <= NORMALIZED_SLIDE_HEIGHT * 0.2 and y <= NORMALIZED_SLIDE_HEIGHT * 0.2:
        return "logo"
    area_ratio = (width * height) / (NORMALIZED_SLIDE_WIDTH * NORMALIZED_SLIDE_HEIGHT)
    if any(marker in href_lower for marker in SCREENSHOT_NAME_MARKERS):
        return "screenshot"
    if area_ratio >= 0.18 and width >= NORMALIZED_SLIDE_WIDTH * 0.32 and height >= NORMALIZED_SLIDE_HEIGHT * 0.22:
        return "screenshot"
    return "content"


def should_ignore_image_ocr(role: str) -> bool:
    return role.strip().lower() == "screenshot"


def should_mask_image_for_vision(image: dict) -> bool:
    if bool(image.get("ocr_ignore")):
        return True
    return should_ignore_image_ocr(str(image.get("role", "")))


def attach_beautify_vision_image_metadata(page_number: int, png_file: Path, slide_metadata: dict) -> None:
    images = slide_metadata.get("images")
    if not isinstance(images, list) or not images:
        return
    masked_images = [image for image in images if isinstance(image, dict) and should_mask_image_for_vision(image)]
    if not masked_images:
        return
    vision_filename = format_slide_vision_png_name(page_number)
    vision_path = png_file.with_name(vision_filename)
    if not create_masked_vision_png(png_file, vision_path, masked_images):
        return
    slide_metadata["vision_image_filename"] = vision_filename
    slide_metadata["vision_image_mime"] = "image/png"
    slide_metadata["vision_masked_image_ids"] = [
        str(image.get("id", "")).strip()
        for image in masked_images
        if str(image.get("id", "")).strip()
    ]


def create_masked_vision_png(source_path: Path, output_path: Path, masked_images: list[dict]) -> bool:
    try:
        with Image.open(source_path) as source:
            image = source.convert("RGB")
            draw = ImageDraw.Draw(image)
            scale_x = image.width / NORMALIZED_SLIDE_WIDTH
            scale_y = image.height / NORMALIZED_SLIDE_HEIGHT
            for item in masked_images:
                rect = normalized_image_rect_to_pixels(item, scale_x, scale_y, image.width, image.height)
                if rect is None:
                    continue
                draw.rectangle(rect, fill=VISION_MASK_FILL, outline=VISION_MASK_STROKE, width=2)
            image.save(output_path, format="PNG")
            return output_path.is_file()
    except Exception:
        return False


def normalized_image_rect_to_pixels(
    image_meta: dict,
    scale_x: float,
    scale_y: float,
    image_width: int,
    image_height: int,
) -> tuple[int, int, int, int] | None:
    x = parse_svg_number(image_meta.get("x", 0))
    y = parse_svg_number(image_meta.get("y", 0))
    width = parse_svg_number(image_meta.get("width", 0))
    height = parse_svg_number(image_meta.get("height", 0))
    if width <= 0 or height <= 0:
        return None
    pad = 2
    left = max(0, int(x * scale_x) - pad)
    top = max(0, int(y * scale_y) - pad)
    right = min(image_width, int((x + width) * scale_x) + pad)
    bottom = min(image_height, int((y + height) * scale_y) + pad)
    if right <= left or bottom <= top:
        return None
    return left, top, right, bottom


def identity_matrix() -> tuple[float, float, float, float, float, float]:
    return 1, 0, 0, 1, 0, 0


def multiply_matrix(
    left: tuple[float, float, float, float, float, float],
    right: tuple[float, float, float, float, float, float],
) -> tuple[float, float, float, float, float, float]:
    la, lb, lc, ld, le, lf = left
    ra, rb, rc, rd, re, rf = right
    return (
        la * ra + lc * rb,
        lb * ra + ld * rb,
        la * rc + lc * rd,
        lb * rc + ld * rd,
        la * re + lc * rf + le,
        lb * re + ld * rf + lf,
    )


def apply_matrix(matrix: tuple[float, float, float, float, float, float], x: float, y: float) -> tuple[float, float]:
    a, b, c, d, e, f = matrix
    return a * x + c * y + e, b * x + d * y + f


def parse_transform(raw: str) -> tuple[float, float, float, float, float, float]:
    matrix = identity_matrix()
    for name, args_raw in TRANSFORM_RE.findall(raw or ""):
        args = [parse_svg_number(part) for part in re.split(r"[,\s]+", args_raw.strip()) if part.strip()]
        transform = identity_matrix()
        lname = name.lower()
        if lname == "translate":
            tx = args[0] if len(args) > 0 else 0
            ty = args[1] if len(args) > 1 else 0
            transform = (1, 0, 0, 1, tx, ty)
        elif lname == "scale":
            sx = args[0] if len(args) > 0 else 1
            sy = args[1] if len(args) > 1 else sx
            transform = (sx, 0, 0, sy, 0, 0)
        elif lname == "matrix" and len(args) >= 6:
            transform = (args[0], args[1], args[2], args[3], args[4], args[5])
        matrix = multiply_matrix(matrix, transform)
    return matrix


def parse_svg_number(raw: str | float | int) -> float:
    if isinstance(raw, (int, float)):
        return float(raw)
    match = re.search(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?", raw or "")
    if not match:
        return 0.0
    try:
        return float(match.group(0))
    except ValueError:
        return 0.0


def clamp_number(value: float, minimum: float, maximum: float) -> float:
    return max(minimum, min(maximum, value))


def attach_chart_fallback_images(charts: list[dict], images: list[dict]) -> list[dict]:
    if not charts or not images:
        return charts
    for chart in charts:
        if not chart.get("fallback_as_image"):
            continue
        best_image = best_overlapping_image(chart, images)
        if best_image is None:
            continue
        chart["fallback_image"] = {
            "id": best_image.get("id", ""),
            "url": best_image.get("url", ""),
            "x": best_image.get("x", 0),
            "y": best_image.get("y", 0),
            "width": best_image.get("width", 0),
            "height": best_image.get("height", 0),
            "role": best_image.get("role", ""),
            "overlap_ratio": best_image.get("_chart_overlap_ratio", 0),
        }
    return charts


def best_overlapping_image(chart: dict, images: list[dict]) -> dict | None:
    best: dict | None = None
    best_ratio = 0.0
    for image in images:
        ratio = rect_overlap_ratio(chart, image)
        if ratio > best_ratio:
            candidate = dict(image)
            candidate["_chart_overlap_ratio"] = round(ratio, 4)
            best = candidate
            best_ratio = ratio
    if best_ratio <= 0:
        return None
    return best


def rect_overlap_ratio(a: dict, b: dict) -> float:
    ax1 = float(a.get("x", 0) or 0)
    ay1 = float(a.get("y", 0) or 0)
    ax2 = ax1 + float(a.get("width", 0) or 0)
    ay2 = ay1 + float(a.get("height", 0) or 0)
    bx1 = float(b.get("x", 0) or 0)
    by1 = float(b.get("y", 0) or 0)
    bx2 = bx1 + float(b.get("width", 0) or 0)
    by2 = by1 + float(b.get("height", 0) or 0)
    overlap_width = max(0.0, min(ax2, bx2) - max(ax1, bx1))
    overlap_height = max(0.0, min(ay2, by2) - max(ay1, by1))
    chart_area = max(1.0, (ax2 - ax1) * (ay2 - ay1))
    return overlap_width * overlap_height / chart_area


def extract_pptx_chart_metadata(input_path: Path) -> dict[int, list[dict]]:
    charts_by_slide: dict[int, list[dict]] = {}
    try:
        with zipfile.ZipFile(input_path, "r") as archive:
            slide_width_emu, slide_height_emu = read_pptx_slide_size_emu(archive)
            slide_entries = collect_ordered_slide_entries(archive)
            for page_number, slide_entry in enumerate(slide_entries, start=1):
                charts_by_slide[page_number] = extract_slide_charts(
                    archive,
                    slide_entry,
                    page_number,
                    slide_width_emu,
                    slide_height_emu,
                )
    except zipfile.BadZipFile as exc:
        raise ConversionError("PPTX file is not a valid zip archive.") from exc
    return charts_by_slide


def read_pptx_slide_size_emu(archive: zipfile.ZipFile) -> tuple[int, int]:
    try:
        root = ET.fromstring(archive.read("ppt/presentation.xml"))
    except (KeyError, ET.ParseError):
        return DEFAULT_SLIDE_WIDTH_EMU, DEFAULT_SLIDE_HEIGHT_EMU
    slide_size = root.find(f"{{{PPT_NS}}}sldSz")
    if slide_size is None:
        return DEFAULT_SLIDE_WIDTH_EMU, DEFAULT_SLIDE_HEIGHT_EMU
    width = parse_int(slide_size.attrib.get("cx", ""), DEFAULT_SLIDE_WIDTH_EMU)
    height = parse_int(slide_size.attrib.get("cy", ""), DEFAULT_SLIDE_HEIGHT_EMU)
    if width <= 0 or height <= 0:
        return DEFAULT_SLIDE_WIDTH_EMU, DEFAULT_SLIDE_HEIGHT_EMU
    return width, height


def extract_slide_charts(
    archive: zipfile.ZipFile,
    slide_entry: str,
    page_number: int,
    slide_width_emu: int,
    slide_height_emu: int,
) -> list[dict]:
    try:
        slide_root = ET.fromstring(archive.read(slide_entry))
    except (KeyError, ET.ParseError):
        return []

    relationships = read_relationships(archive, relationship_part_for(slide_entry))
    charts: list[dict] = []
    for chart_index, graphic_frame in enumerate(slide_root.iter(f"{{{PPT_NS}}}graphicFrame"), start=1):
        chart_rel_id = find_chart_relationship_id(graphic_frame)
        if not chart_rel_id:
            continue
        chart_target = relationships.get(chart_rel_id, "")
        chart_path = normalize_pptx_part_path(posixpath.dirname(slide_entry), chart_target) if chart_target else ""
        chart_meta = {
            "id": f"chart-{page_number:03d}-{chart_index:03d}",
            "chart_type": "unknown",
            "chart_type_label": CHART_TYPE_LABELS["unknown"],
            "raw_chart_type": "",
            **graphic_frame_bounds(graphic_frame, slide_width_emu, slide_height_emu),
            "data": [],
            "style": {},
            "fallback_as_image": True,
        }
        if chart_path and chart_path in archive.namelist():
            try:
                parsed = parse_chart_part(archive.read(chart_path))
                requires_fallback = bool(parsed.pop("requires_fallback", False))
                chart_meta.update(parsed)
                chart_meta["fallback_as_image"] = (
                    requires_fallback
                    or chart_meta.get("chart_type") not in SUPPORTED_CHART_TYPES
                    or len(parsed.get("data", [])) == 0
                )
            except ConversionError:
                chart_meta["fallback_as_image"] = True
        charts.append(chart_meta)
    return charts


def relationship_part_for(part_name: str) -> str:
    directory = posixpath.dirname(part_name)
    filename = posixpath.basename(part_name)
    return posixpath.join(directory, "_rels", f"{filename}.rels")


def read_relationships(archive: zipfile.ZipFile, rels_entry: str) -> dict[str, str]:
    try:
        root = ET.fromstring(archive.read(rels_entry))
    except (KeyError, ET.ParseError):
        return {}
    relationships: dict[str, str] = {}
    for relationship in root:
        if local_name(relationship.tag) != "Relationship":
            continue
        rel_id = relationship.attrib.get("Id", "")
        target = relationship.attrib.get("Target", "")
        if rel_id and target:
            relationships[rel_id] = target
    return relationships


def find_chart_relationship_id(graphic_frame: ET.Element) -> str:
    for chart in graphic_frame.iter(f"{{{CHART_NS}}}chart"):
        rel_id = chart.attrib.get(f"{{{REL_NS}}}id", "")
        if rel_id:
            return rel_id
    return ""


def graphic_frame_bounds(graphic_frame: ET.Element, slide_width_emu: int, slide_height_emu: int) -> dict:
    off = None
    ext = None
    for xfrm in graphic_frame.iter(f"{{{PPT_NS}}}xfrm"):
        off = xfrm.find(f"{{{DRAWING_NS}}}off")
        ext = xfrm.find(f"{{{DRAWING_NS}}}ext")
        break
    if off is None or ext is None:
        return {"x": 0, "y": 0, "width": 0, "height": 0}
    x = parse_int(off.attrib.get("x", ""), 0) / slide_width_emu * NORMALIZED_SLIDE_WIDTH
    y = parse_int(off.attrib.get("y", ""), 0) / slide_height_emu * NORMALIZED_SLIDE_HEIGHT
    width = parse_int(ext.attrib.get("cx", ""), 0) / slide_width_emu * NORMALIZED_SLIDE_WIDTH
    height = parse_int(ext.attrib.get("cy", ""), 0) / slide_height_emu * NORMALIZED_SLIDE_HEIGHT
    return {
        "x": round(x, 2),
        "y": round(y, 2),
        "width": round(width, 2),
        "height": round(height, 2),
    }


def parse_chart_part(chart_xml: bytes) -> dict:
    try:
        root = ET.fromstring(chart_xml)
    except ET.ParseError as exc:
        raise ConversionError(f"Chart XML is not valid: {exc}") from exc

    chart_containers: list[ET.Element] = []
    for candidate in root.iter():
        lname = local_name(candidate.tag)
        if lname.endswith("Chart") and lname not in {"chart", "plotArea"}:
            chart_containers.append(candidate)

    data_rows: list[dict] = []
    normalized_types: list[str] = []
    raw_chart_types: list[str] = []
    for chart_container in chart_containers:
        raw_chart_type = local_name(chart_container.tag).removesuffix("Chart")
        normalized_type = normalize_chart_type(raw_chart_type, chart_container)
        raw_chart_types.append(raw_chart_type)
        normalized_types.append(normalized_type)
        for series_index, series in enumerate(chart_container.iter(f"{{{CHART_NS}}}ser")):
            data_rows.extend(parse_chart_series(series, series_index, normalized_type))

    chart_type = next((item for item in normalized_types if item in SUPPORTED_CHART_TYPES), "unknown")
    raw_chart_type = raw_chart_types[0] if raw_chart_types else ""
    unique_supported_types = {item for item in normalized_types if item in SUPPORTED_CHART_TYPES}
    requires_fallback = len(unique_supported_types) > 1

    title = extract_chart_title(root)
    style = {"title": title} if title else {}
    return {
        "chart_type": chart_type,
        "chart_type_label": CHART_TYPE_LABELS.get(chart_type, CHART_TYPE_LABELS["unknown"]),
        "raw_chart_type": raw_chart_type,
        "data": data_rows,
        "style": style,
        "requires_fallback": requires_fallback,
    }


def normalize_chart_type(raw_chart_type: str, chart_container: ET.Element) -> str:
    raw = raw_chart_type.strip().lower()
    if raw.startswith("bar"):
        bar_dir = chart_container.find(f"{{{CHART_NS}}}barDir")
        direction = (bar_dir.attrib.get("val", "") if bar_dir is not None else "").lower()
        return "horizontal-bar" if direction == "bar" else "bar"
    if raw.startswith("line"):
        return "line"
    if raw.startswith("pie"):
        return "pie"
    if raw.startswith("doughnut") or raw.startswith("donut"):
        return "donut"
    return "unknown"


def parse_chart_series(series: ET.Element, series_index: int, chart_type: str) -> list[dict]:
    name = ""
    tx = series.find(f"{{{CHART_NS}}}tx")
    if tx is not None:
        name = first_chart_text(tx)

    categories = chart_axis_values(series.find(f"{{{CHART_NS}}}cat"))
    if not categories:
        categories = chart_axis_values(series.find(f"{{{CHART_NS}}}xVal"))
    values = chart_axis_values(series.find(f"{{{CHART_NS}}}val"))
    if not values:
        values = chart_axis_values(series.find(f"{{{CHART_NS}}}yVal"))

    rows: list[dict] = []
    color = CHART_DEFAULT_PALETTE[series_index % len(CHART_DEFAULT_PALETTE)]
    for idx, value in enumerate(values):
        category = categories[idx] if idx < len(categories) else str(idx + 1)
        row = {
            "label": category,
            "value": parse_chart_value(value),
            "color": CHART_DEFAULT_PALETTE[idx % len(CHART_DEFAULT_PALETTE)] if chart_type in {"pie", "donut"} else color,
            "category": category,
        }
        if name:
            row["group"] = name
            row["series"] = name
        rows.append(row)
    return rows


def chart_axis_values(node: ET.Element | None) -> list[str]:
    if node is None:
        return []
    values: list[tuple[int, str]] = []
    for pt in node.iter(f"{{{CHART_NS}}}pt"):
        idx = int(pt.attrib.get("idx", "0") or 0)
        v = pt.find(f"{{{CHART_NS}}}v")
        if v is not None and v.text is not None:
            values.append((idx, v.text))
    return [value for _, value in sorted(values)]


def parse_chart_value(value: str) -> float | str:
    try:
        return float(value)
    except (TypeError, ValueError):
        return value


def parse_int(value: str, fallback: int) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return fallback


def first_chart_text(node: ET.Element) -> str:
    for v in node.iter(f"{{{CHART_NS}}}v"):
        if v.text:
            return v.text
    for t in node.iter(f"{{{DRAWING_NS}}}t"):
        if t.text:
            return t.text
    return ""


def extract_chart_title(root: ET.Element) -> str:
    title = root.find(f".//{{{CHART_NS}}}title")
    if title is None:
        return ""
    return first_chart_text(title)


def format_chinese_number(number: int) -> str:
    digits = "零一二三四五六七八九"
    if number <= 0:
        return str(number)
    if number < 10:
        return digits[number]
    if number < 20:
        return "十" + (digits[number % 10] if number % 10 else "")
    if number < 100:
        tens, ones = divmod(number, 10)
        return digits[tens] + "十" + (digits[ones] if ones else "")
    if number < 1000:
        hundreds, rest = divmod(number, 100)
        if rest == 0:
            return digits[hundreds] + "百"
        connector = "零" if rest < 10 else ""
        return digits[hundreds] + "百" + connector + format_chinese_number(rest)
    return str(number)


def build_zip_bytes(svg_files: Iterable[Path], metadata: dict | None = None) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for index, svg_file in enumerate(svg_files, start=1):
            archive.write(svg_file, arcname=f"slide-{index:03d}.svg")
        if metadata is not None:
            archive.writestr("metadata.json", json.dumps(metadata, ensure_ascii=False, indent=2))
    buffer.seek(0)
    return buffer.getvalue()


def build_beautify_png_zip_bytes(
    png_files: Iterable[Path],
    asset_files: Iterable[tuple[Path, str]],
    metadata: dict,
) -> bytes:
    buffer = io.BytesIO()
    png_file_list = list(png_files)
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        written_names: set[str] = set()
        for index, png_file in enumerate(png_file_list, start=1):
            arcname = format_slide_png_name(index)
            archive.write(png_file, arcname=arcname)
            written_names.add(arcname)
        for slide in metadata.get("slides", []):
            if not isinstance(slide, dict):
                continue
            vision_filename = str(slide.get("vision_image_filename", "")).strip()
            if not vision_filename or vision_filename in written_names:
                continue
            slide_index = int(slide.get("index") or 0)
            if slide_index <= 0 or slide_index > len(png_file_list):
                continue
            vision_path = png_file_list[slide_index - 1].with_name(vision_filename)
            if not vision_path.is_file():
                continue
            archive.write(vision_path, arcname=vision_filename)
            written_names.add(vision_filename)
        for asset_path, arcname in asset_files:
            archive.write(asset_path, arcname=arcname)
        archive.writestr("metadata.json", json.dumps(metadata, ensure_ascii=False, indent=2))
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


def export_presentation_to_pdf(input_path: Path, working_dir: Path, profile_dir: Path) -> Path:
    export_dir = working_dir / "pdf-export"
    export_dir.mkdir(parents=True, exist_ok=True)
    run_command(
        build_libreoffice_command(
            profile_dir,
            [
                "--convert-to",
                settings.libreoffice_pdf_filter,
                "--outdir",
                str(export_dir),
                str(input_path),
            ],
        )
    )

    exported_pdf = collect_single_output_file(export_dir, ".pdf")
    target_pdf = working_dir / f"{input_path.stem}.pdf"
    shutil.move(str(exported_pdf), target_pdf)
    return target_pdf


def export_presentation_to_svg(input_path: Path, working_dir: Path, profile_dir: Path) -> Path:
    export_dir = working_dir / "libreoffice-text-svg"
    export_dir.mkdir(parents=True, exist_ok=True)
    run_command(
        build_libreoffice_command(
            profile_dir,
            [
                "--convert-to",
                settings.libreoffice_svg_filter,
                "--outdir",
                str(export_dir),
                str(input_path),
            ],
        )
    )

    exported_svg = collect_single_output_file(export_dir, ".svg")
    target_svg = working_dir / f"{input_path.stem}.svg"
    shutil.move(str(exported_svg), target_svg)
    return target_svg


def split_libreoffice_svg(source_svg: Path, output_dir: Path, asset_prefix: str) -> list[Path]:
    register_svg_namespaces(source_svg)
    try:
        tree = ET.parse(source_svg)
    except ET.ParseError as exc:
        raise ConversionError(f"LibreOffice SVG is not valid XML: {exc}") from exc

    root = tree.getroot()
    slide_ids = collect_slide_ids(root)
    slide_group = find_child_by_class(root, "SlideGroup")
    if slide_group is None:
        raise ConversionError("LibreOffice SVG did not contain a SlideGroup.")

    slide_containers = collect_slide_containers(slide_group)
    if not slide_ids:
        slide_ids = sorted(slide_containers)

    svg_files: list[Path] = []
    image_upload_cache: ImageUploadCache = {}
    for page_number, slide_id in enumerate(slide_ids, start=1):
        container = slide_containers.get(slide_id)
        if container is None:
            raise ConversionError(f"LibreOffice SVG is missing container for {slide_id}.")

        page_root = build_single_slide_svg(root, slide_group, container)
        normalize_svg_fonts(page_root)
        if settings.strip_svg_embedded_fonts:
            strip_svg_embedded_fonts(page_root)
        externalize_svg_images(page_root, asset_prefix, image_upload_cache)

        target_svg = output_dir / format_slide_name(page_number)
        ET.ElementTree(page_root).write(target_svg, encoding="utf-8", xml_declaration=True)
        if target_svg.stat().st_size <= 0:
            raise ConversionError(f"Generated empty SVG for {slide_id}.")
        svg_files.append(target_svg)

    return svg_files


def register_svg_namespaces(svg_path: Path) -> None:
    for _, namespace in ET.iterparse(svg_path, events=("start-ns",)):
        prefix, uri = namespace
        ET.register_namespace(prefix, uri)


def collect_slide_ids(root: ET.Element) -> list[str]:
    slide_ids: list[str] = []
    slide_attr = f"{{{OOO_NS}}}slide"
    for element in root.iter():
        element_id = element.attrib.get("id", "")
        slide_id = element.attrib.get(slide_attr)
        if element_id.startswith("ooo:meta_slide_") and slide_id:
            slide_ids.append(slide_id)
    return slide_ids


def find_child_by_class(root: ET.Element, class_name: str) -> ET.Element | None:
    for child in root:
        if local_name(child.tag) == "g" and child.attrib.get("class") == class_name:
            return child
    return None


def collect_slide_containers(slide_group: ET.Element) -> dict[str, ET.Element]:
    slide_containers: dict[str, ET.Element] = {}
    for container in slide_group.iter():
        if local_name(container.tag) != "g" or not container.attrib.get("id", "").startswith("container-"):
            continue

        for child in container:
            if local_name(child.tag) != "g":
                continue
            if child.attrib.get("class") != "Slide":
                continue

            slide_id = child.attrib.get("id")
            if slide_id:
                slide_containers[slide_id] = container
            break
    return slide_containers


def build_single_slide_svg(
    original_root: ET.Element,
    original_slide_group: ET.Element,
    slide_container: ET.Element,
) -> ET.Element:
    page_root = ET.Element(original_root.tag, dict(original_root.attrib))
    for child in original_root:
        if should_skip_root_child(child):
            continue

        if child is original_slide_group:
            new_slide_group = ET.Element(child.tag, dict(child.attrib))
            new_slide_group.append(copy.deepcopy(slide_container))
            page_root.append(new_slide_group)
            continue

        page_root.append(copy.deepcopy(child))

    return page_root


def should_skip_root_child(element: ET.Element) -> bool:
    element_class = element.attrib.get("class", "")
    element_id = element.attrib.get("id", "")
    if local_name(element.tag) == "script":
        return True
    if local_name(element.tag) == "g" and element_class == "DummySlide":
        return True
    if local_name(element.tag) == "defs" and element_id == "presentation-animations":
        return True
    if local_name(element.tag) == "defs" and element_class in {"Animations", "TextShapeIndex"}:
        return True
    return False


def normalize_svg_fonts(element: ET.Element) -> None:
    fixed_font = settings.svg_fixed_font_family
    for node in element.iter():
        if "font-family" in node.attrib:
            node.set("font-family", fixed_font)

        style = node.attrib.get("style")
        if style and "font-family" in style:
            node.set("style", STYLE_FONT_FAMILY_RE.sub(f"font-family:{fixed_font}", style))


def strip_svg_embedded_fonts(element: ET.Element) -> None:
    for child in list(element):
        if local_name(child.tag) == "font":
            element.remove(child)
            continue
        strip_svg_embedded_fonts(child)


def externalize_svg_images(element: ET.Element, asset_prefix: str, upload_cache: ImageUploadCache) -> None:
    if not settings.aliyun_oss_enabled:
        return

    href_attrs = ("href", f"{{{XLINK_NS}}}href")
    for node in element.iter():
        if local_name(node.tag) != "image":
            continue

        href_attr = next((attr for attr in href_attrs if attr in node.attrib), "")
        if not href_attr:
            continue

        href = node.attrib[href_attr]
        match = DATA_IMAGE_RE.match(href)
        if not match:
            continue

        content_type = match.group(1).lower()
        try:
            image_bytes = base64.b64decode(match.group(2), validate=True)
        except ValueError as exc:
            raise ConversionError("SVG contains an invalid base64 image.") from exc

        extension = guess_image_extension(content_type)
        digest = hashlib.sha256(image_bytes).hexdigest()
        cache_key = (content_type, digest)
        public_url = upload_cache.get(cache_key)
        if public_url is None:
            object_key = f"{asset_prefix}/images/{digest[:24]}{extension}"
            public_url = upload_bytes_to_aliyun_oss(object_key, image_bytes, content_type)
            upload_cache[cache_key] = public_url
        node.set(href_attr, public_url)


def upload_bytes_to_aliyun_oss(object_key: str, data: bytes, content_type: str) -> str:
    endpoint = settings.aliyun_oss_endpoint.removeprefix("https://").removeprefix("http://").strip("/")
    bucket_name = settings.aliyun_oss_bucket_name
    encoded_key = quote(object_key, safe="/")
    upload_url = f"https://{bucket_name}.{endpoint}/{encoded_key}"
    date_header = formatdate(usegmt=True)
    content_md5 = base64.b64encode(hashlib.md5(data).digest()).decode("ascii")
    resource = f"/{bucket_name}/{object_key}"
    string_to_sign = f"PUT\n{content_md5}\n{content_type}\n{date_header}\n{resource}"
    signature = base64.b64encode(
        hmac.new(
            settings.aliyun_oss_access_key_secret.encode("utf-8"),
            string_to_sign.encode("utf-8"),
            hashlib.sha1,
        ).digest()
    ).decode("ascii")
    headers = {
        "Authorization": f"OSS {settings.aliyun_oss_access_key_id}:{signature}",
        "Content-MD5": content_md5,
        "Content-Type": content_type,
        "Date": date_header,
    }

    try:
        response = httpx.put(upload_url, content=data, headers=headers, timeout=settings.command_timeout_seconds)
        response.raise_for_status()
    except httpx.HTTPError as exc:
        raise ConversionError(f"Failed to upload SVG image to OSS: {exc}") from exc

    return f"{settings.aliyun_oss_base_url}/{encoded_key}"


def guess_image_extension(content_type: str) -> str:
    if content_type == "image/jpeg":
        return ".jpg"
    if content_type == "image/svg+xml":
        return ".svg"
    return mimetypes.guess_extension(content_type) or ".bin"


def build_asset_prefix(input_path: Path) -> str:
    base_prefix = settings.aliyun_oss_prefix
    presentation_name = sanitize_filename(input_path.stem)
    unique_id = uuid4().hex[:12]
    if base_prefix:
        return f"{base_prefix}/{presentation_name}-{unique_id}"
    return f"{presentation_name}-{unique_id}"


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def render_pdf_to_svgs(pdf_path: Path, output_dir: Path) -> list[Path]:
    run_command(
        [
            get_required_command(settings.mupdf_command),
            "draw",
            "-q",
            "-F",
            "svg",
            "-o",
            str(output_dir / "slide-%03d.svg"),
            str(pdf_path),
        ]
    )

    svg_files = sorted(output_dir.glob("slide-*.svg"))
    if not svg_files:
        raise ConversionError("MuPDF did not generate any SVG slides from PDF fallback.")
    return svg_files


def render_pdf_to_pngs(pdf_path: Path, output_dir: Path) -> list[Path]:
    dpi = max(72, min(settings.png_render_dpi, 240))
    run_command(
        [
            get_required_command(settings.mupdf_command),
            "draw",
            "-q",
            "-r",
            str(dpi),
            "-F",
            "png",
            "-o",
            str(output_dir / "slide-%03d.png"),
            str(pdf_path),
        ]
    )

    png_files = sorted(output_dir.glob("slide-*.png"))
    if not png_files:
        raise ConversionError("MuPDF did not generate any PNG slides from PDF.")
    return png_files


def render_pdf_page_to_png(pdf_path: Path, output_dir: Path, page_number: int) -> Path:
    dpi = max(72, min(settings.png_render_dpi, 240))
    target_path = output_dir / format_slide_png_name(page_number)
    run_command(
        [
            get_required_command(settings.mupdf_command),
            "draw",
            "-q",
            "-r",
            str(dpi),
            "-F",
            "png",
            "-o",
            str(target_path),
            str(pdf_path),
            str(page_number),
        ]
    )
    if not target_path.exists():
        raise ConversionError(f"MuPDF did not generate PNG for page {page_number}.")
    return target_path


def read_pdf_page_count(pdf_path: Path) -> int:
    result = run_command([get_required_command(settings.mupdf_command), "info", str(pdf_path)])
    text = "\n".join([result.stdout or "", result.stderr or ""])
    match = re.search(r"Pages:\s*(\d+)", text, re.IGNORECASE)
    if match:
        return int(match.group(1))
    raise ConversionError("MuPDF did not report PDF page count.")


def build_libreoffice_command(profile_dir: Path, extra_args: list[str]) -> list[str]:
    profile_uri = profile_dir.resolve().as_uri()
    return [
        get_required_command(settings.libreoffice_command),
        "--headless",
        "--invisible",
        "--nodefault",
        "--nolockcheck",
        "--nologo",
        "--nofirststartwizard",
        "--norestore",
        f"-env:UserInstallation={profile_uri}",
        *extra_args,
    ]


def collect_single_output_file(export_dir: Path, suffix: str) -> Path:
    exported_files = sorted(path for path in export_dir.iterdir() if path.suffix.lower() == suffix)
    if len(exported_files) != 1:
        raise ConversionError(
            f"Expected exactly one {suffix} file in {export_dir}, got {len(exported_files)}."
        )
    return exported_files[0]


def format_slide_name(index: int) -> str:
    return f"slide-{index:03d}.svg"


def format_slide_png_name(index: int) -> str:
    return f"slide-{index:03d}.png"


def format_slide_vision_png_name(index: int) -> str:
    return f"slide-{index:03d}-vision.png"


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


def save_text_bytes(text_name: str, text_bytes: bytes) -> tuple[str, Path]:
    settings.downloads_root.mkdir(parents=True, exist_ok=True)
    base_name = sanitize_filename(Path(text_name).stem)
    unique_name = f"{base_name}-{uuid4().hex[:12]}.txt"
    target_path = settings.downloads_root / unique_name
    target_path.write_bytes(text_bytes)
    return unique_name, target_path


def save_json_bytes(json_name: str, payload: dict) -> tuple[str, Path]:
    settings.downloads_root.mkdir(parents=True, exist_ok=True)
    base_name = sanitize_filename(Path(json_name).stem)
    unique_name = f"{base_name}-{uuid4().hex[:12]}.json"
    target_path = settings.downloads_root / unique_name
    target_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return unique_name, target_path
