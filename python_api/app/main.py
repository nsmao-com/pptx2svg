from __future__ import annotations

from fastapi import FastAPI, HTTPException
from fastapi.responses import JSONResponse, Response
from pydantic import BaseModel, HttpUrl

from .config import settings
from .converter import ConversionError, convert_ppt_url_to_svg_zip, ensure_dependencies


app = FastAPI(title=settings.app_name, version="0.1.0")


class ConvertRequest(BaseModel):
    ppt_url: HttpUrl


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

    headers = {
        "Content-Disposition": f'attachment; filename="{archive_name}"',
    }
    return Response(content=archive_bytes, media_type="application/zip", headers=headers)
