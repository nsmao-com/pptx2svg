# input: HTTP requests carrying remote PPT/PPTX URLs.
# output: Health, SVG ZIP conversion, beautify PNG analysis, text extraction, and download responses.
# pos: FastAPI routing layer for the copied pptx2svg subsystem.
# update: When this file changes, update this header and pptx2svg_api/app/README.md.

from __future__ import annotations

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse, JSONResponse, Response
from pydantic import BaseModel, HttpUrl

from .config import settings
from .converter import (
    ConversionError,
    analyze_ppt_url_for_beautify,
    convert_ppt_url_to_svg_zip,
    ensure_dependencies,
    extract_ppt_url_to_text,
    save_json_bytes,
    save_text_bytes,
    save_zip_bytes,
)


app = FastAPI(title=settings.app_name, version="0.1.0")


class ConvertRequest(BaseModel):
    ppt_url: HttpUrl
    url: bool = False


@app.get("/healthz")
def healthz() -> JSONResponse:
    try:
        ensure_dependencies()
    except ConversionError as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc
    return JSONResponse({"status": "ok"})


@app.post("/api/v1/convert/ppt-to-svg")
def convert_ppt_to_svg(payload: ConvertRequest) -> Response:
    try:
        archive_name, archive_bytes = convert_ppt_url_to_svg_zip(str(payload.ppt_url))
    except ConversionError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    if payload.url:
        saved_name, _ = save_zip_bytes(archive_name, archive_bytes)
        return JSONResponse(
            {
                "filename": saved_name,
                "url": f"/downloads/{saved_name}",
            }
        )

    headers = {
        "Content-Disposition": f'attachment; filename="{archive_name}"',
    }
    return Response(content=archive_bytes, media_type="application/zip", headers=headers)


@app.post("/api/v1/analyze/pptx-beautify")
def analyze_pptx_beautify(payload: ConvertRequest) -> Response:
    try:
        archive_name, archive_bytes, metadata = analyze_ppt_url_for_beautify(str(payload.ppt_url))
    except ConversionError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    if payload.url:
        saved_name, _ = save_zip_bytes(archive_name, archive_bytes)
        metadata_name, _ = save_json_bytes(f"{archive_name}.metadata.json", metadata)
        return JSONResponse(
            {
                "filename": saved_name,
                "url": f"/downloads/{saved_name}",
                "metadata_url": f"/downloads/{metadata_name}",
                "total_pages": metadata.get("total_pages", 0),
            }
        )

    headers = {
        "Content-Disposition": f'attachment; filename="{archive_name}"',
    }
    return Response(content=archive_bytes, media_type="application/zip", headers=headers)


@app.post("/api/v1/extract/ppt-text")
def extract_ppt_text(payload: ConvertRequest) -> Response:
    try:
        text_name, text_bytes = extract_ppt_url_to_text(str(payload.ppt_url))
    except ConversionError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    if payload.url:
        saved_name, _ = save_text_bytes(text_name, text_bytes)
        return JSONResponse(
            {
                "filename": saved_name,
                "url": f"/downloads/{saved_name}",
            }
        )

    headers = {
        "Content-Disposition": f'attachment; filename="{text_name}"',
    }
    return Response(content=text_bytes, media_type="text/plain; charset=utf-8", headers=headers)


@app.get("/downloads/{filename}")
def download_generated_archive(filename: str) -> FileResponse:
    safe_name = filename.rsplit("/", 1)[-1].rsplit("\\", 1)[-1]
    if safe_name != filename:
        raise HTTPException(status_code=400, detail="Invalid filename.")

    archive_path = settings.downloads_root / safe_name
    if not archive_path.is_file():
        raise HTTPException(status_code=404, detail="File not found.")

    media_type = "text/plain; charset=utf-8" if archive_path.suffix.lower() == ".txt" else "application/zip"

    return FileResponse(archive_path, media_type=media_type, filename=archive_path.name)
