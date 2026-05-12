# input: SVG列表请求、PPT/PPTX URL请求与转换参数HTTP请求
# output: 统一的2222 API，覆盖队列化SVG->PPTX、SVG->PNG、PPT/PPTX->SVG ZIP、beautify PNG解析/遮挡识图图、逐页ready流水线、文本提取、异步任务与下载
# pos: 2pptxsvg 统一API入口
# 一旦我被更新，务必更新我的开头注释，以及所属的文件夹README。

from __future__ import annotations

import re
import shutil
import tempfile
import threading
import time
import uuid
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from os import getenv
from pathlib import Path
from typing import Dict, List, Literal, Optional

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse, JSONResponse, Response
from pydantic import BaseModel, ConfigDict, Field, HttpUrl

from pptx2svg_api.app.config import settings as pptx_settings
from pptx2svg_api.app.converter import (
    ConversionError as PPTXConversionError,
    analyze_ppt_url_for_beautify,
    analyze_ppt_url_for_beautify_progressive,
    convert_ppt_url_to_svg_zip,
    ensure_dependencies as ensure_pptx_dependencies,
    extract_ppt_url_to_text,
    save_json_bytes,
    save_text_bytes,
    save_zip_bytes,
)
from svg_to_editable_pptx import (
    convert_svg_files_to_pptx,
    parse_svg_canvas_size,
    rasterize_svg_to_png,
)


def _sanitize_file_stem(name: str) -> str:
    stem = re.sub(r"[^a-zA-Z0-9_-]+", "_", name.strip())
    stem = stem.strip("._-")
    if not stem:
        return "editable_output"
    return stem[:80]


def _now_ts() -> float:
    return time.time()


class ConvertRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    svgs: List[str] = Field(..., min_length=1, max_length=200)
    title: str = Field(default="editable_output", min_length=1, max_length=120)
    use_bg_image: bool = True
    strict: bool = True


class PPTXConvertRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    ppt_url: HttpUrl
    url: bool = False


class PPTXJobCreateRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    ppt_url: HttpUrl


class RasterizeRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    svg: str = Field(..., min_length=1)
    width: Optional[int] = Field(default=None, ge=1, le=16384)
    height: Optional[int] = Field(default=None, ge=1, le=16384)
    scale: float = Field(default=1.0, gt=0.0, le=8.0)


class JobResponse(BaseModel):
    model_config = ConfigDict(extra="forbid")

    job_id: str
    status: Literal["queued", "running", "succeeded", "failed"]
    created_at: int
    updated_at: int
    total: int
    success: int
    failed: int
    error: Optional[str] = None
    download_url: Optional[str] = None


class PPTXJobResponse(BaseModel):
    model_config = ConfigDict(extra="forbid")

    job_id: str
    kind: Literal["convert", "beautify", "beautify_pages", "extract"]
    status: Literal["queued", "running", "succeeded", "failed"]
    created_at: int
    updated_at: int
    error: Optional[str] = None
    download_url: Optional[str] = None
    metadata_url: Optional[str] = None
    total_pages: Optional[int] = None
    ready_pages: Optional[int] = None
    current_page: Optional[int] = None
    phase: Optional[str] = None
    pages: Optional[List[dict]] = None


@dataclass
class JobRecord:
    job_id: str
    status: Literal["queued", "running", "succeeded", "failed"]
    created_at: float
    updated_at: float
    total: int
    title: str
    use_bg_image: bool
    strict: bool
    work_dir: Path
    output_path: Optional[Path] = None
    success_count: int = 0
    failed_count: int = 0
    error: Optional[str] = None


class ConversionJobManager:
    def __init__(
        self,
        work_root: Path,
        max_workers: int,
        ttl_seconds: int,
        convert_timeout_seconds: Optional[float] = None,
    ) -> None:
        self.work_root = work_root
        self.max_workers = max_workers
        self.ttl_seconds = ttl_seconds
        self.convert_timeout_seconds = convert_timeout_seconds
        self._lock = threading.Lock()
        self._jobs: Dict[str, JobRecord] = {}
        self._executor = ThreadPoolExecutor(max_workers=max_workers)
        self.work_root.mkdir(parents=True, exist_ok=True)

    def shutdown(self) -> None:
        self._executor.shutdown(wait=False, cancel_futures=False)

    def _cleanup_expired_locked(self) -> None:
        now = _now_ts()
        stale_ids: List[str] = []
        for job_id, job in self._jobs.items():
            if job.status in {"queued", "running"}:
                continue
            if now - job.updated_at > self.ttl_seconds:
                stale_ids.append(job_id)

        for job_id in stale_ids:
            job = self._jobs.pop(job_id, None)
            if job is None:
                continue
            shutil.rmtree(job.work_dir, ignore_errors=True)

    def submit(self, req: ConvertRequest) -> JobRecord:
        with self._lock:
            self._cleanup_expired_locked()
            job_id = uuid.uuid4().hex
            created = _now_ts()
            title = _sanitize_file_stem(req.title)
            work_dir = self.work_root / job_id
            work_dir.mkdir(parents=True, exist_ok=True)

            job = JobRecord(
                job_id=job_id,
                status="queued",
                created_at=created,
                updated_at=created,
                total=len(req.svgs),
                title=title,
                use_bg_image=req.use_bg_image,
                strict=req.strict,
                work_dir=work_dir,
            )
            self._jobs[job_id] = job

        # 在线程池中执行，支持多个任务并发转换。
        self._executor.submit(self._run_job, job_id, req)
        return job

    def _run_job(self, job_id: str, req: ConvertRequest) -> None:
        with self._lock:
            job = self._jobs.get(job_id)
            if job is None:
                return
            job.status = "running"
            job.updated_at = _now_ts()

        try:
            with self._lock:
                job = self._jobs.get(job_id)
                if job is None:
                    return
                work_dir = job.work_dir

            input_dir = work_dir / "input"
            input_dir.mkdir(parents=True, exist_ok=True)

            svg_paths: List[Path] = []
            for index, svg in enumerate(req.svgs, start=1):
                svg_path = input_dir / f"slide_{index:04d}.svg"
                svg_path.write_text(svg, encoding="utf-8")
                svg_paths.append(svg_path)

            output_path = work_dir / f"{_sanitize_file_stem(req.title)}.pptx"
            result = convert_svg_files_to_pptx(
                svg_files=svg_paths,
                output_path=output_path,
                use_bg_image=req.use_bg_image,
                strict=req.strict,
                log_fn=None,
                timeout_seconds=self.convert_timeout_seconds,
            )

            with self._lock:
                job = self._jobs.get(job_id)
                if job is None:
                    return
                job.status = "succeeded"
                job.updated_at = _now_ts()
                job.output_path = output_path
                job.success_count = int(result["success"])
                job.failed_count = len(result["failed"])
                job.error = None

        except Exception as exc:
            with self._lock:
                job = self._jobs.get(job_id)
                if job is None:
                    return
                job.status = "failed"
                job.updated_at = _now_ts()
                job.error = str(exc)

    def run_sync_convert(self, req: ConvertRequest) -> bytes:
        with tempfile.TemporaryDirectory(prefix="svg2ppt_sync_", dir=str(self.work_root)) as work_dir:
            work_path = Path(work_dir)
            input_dir = work_path / "input"
            input_dir.mkdir(parents=True, exist_ok=True)

            svg_paths: List[Path] = []
            for index, svg in enumerate(req.svgs, start=1):
                svg_path = input_dir / f"slide_{index:04d}.svg"
                svg_path.write_text(svg, encoding="utf-8")
                svg_paths.append(svg_path)

            title = _sanitize_file_stem(req.title)
            output_path = work_path / f"{title}.pptx"

            def _do_convert() -> None:
                convert_svg_files_to_pptx(
                    svg_files=svg_paths,
                    output_path=output_path,
                    use_bg_image=req.use_bg_image,
                    strict=req.strict,
                    log_fn=None,
                    timeout_seconds=self.convert_timeout_seconds,
                )

            future = self._executor.submit(_do_convert)
            try:
                future.result()
            except TimeoutError:
                raise
            except Exception as exc:
                raise HTTPException(status_code=400, detail=str(exc)) from exc

            if not output_path.exists():
                raise HTTPException(status_code=500, detail="output file missing")

            return output_path.read_bytes()

    def get(self, job_id: str) -> Optional[JobRecord]:
        with self._lock:
            self._cleanup_expired_locked()
            job = self._jobs.get(job_id)
            if job is None:
                return None
            return JobRecord(**job.__dict__)

    def delete(self, job_id: str) -> bool:
        with self._lock:
            job = self._jobs.get(job_id)
            if job is None:
                return False
            if job.status in {"queued", "running"}:
                raise ValueError("job is still running")
            self._jobs.pop(job_id, None)
            work_dir = job.work_dir

        shutil.rmtree(work_dir, ignore_errors=True)
        return True


@dataclass
class PPTXJobRecord:
    job_id: str
    kind: Literal["convert", "beautify", "beautify_pages", "extract"]
    status: Literal["queued", "running", "succeeded", "failed"]
    created_at: float
    updated_at: float
    source_url: str
    work_dir: Optional[Path] = None
    download_name: Optional[str] = None
    download_path: Optional[Path] = None
    metadata_name: Optional[str] = None
    metadata_path: Optional[Path] = None
    total_pages: Optional[int] = None
    ready_pages: int = 0
    current_page: int = 0
    phase: Optional[str] = None
    pages: List[dict] = field(default_factory=list)
    error: Optional[str] = None


class PPTXTaskManager:
    def __init__(self, work_root: Path, max_workers: int, ttl_seconds: int) -> None:
        self.work_root = work_root
        self.max_workers = max_workers
        self.ttl_seconds = ttl_seconds
        self._lock = threading.Lock()
        self._jobs: Dict[str, PPTXJobRecord] = {}
        self._executor = ThreadPoolExecutor(max_workers=max_workers)
        self.work_root.mkdir(parents=True, exist_ok=True)

    def shutdown(self) -> None:
        self._executor.shutdown(wait=False, cancel_futures=False)

    def _cleanup_job_files(self, job: PPTXJobRecord) -> None:
        for artifact_path in (job.download_path, job.metadata_path):
            if artifact_path is None:
                continue
            try:
                artifact_path.unlink(missing_ok=True)
            except OSError:
                continue
        if job.work_dir is not None:
            shutil.rmtree(job.work_dir, ignore_errors=True)

    def _cleanup_expired_locked(self) -> None:
        now = _now_ts()
        stale_ids: List[str] = []
        for job_id, job in self._jobs.items():
            if job.status in {"queued", "running"}:
                continue
            if now - job.updated_at > self.ttl_seconds:
                stale_ids.append(job_id)

        for job_id in stale_ids:
            job = self._jobs.pop(job_id, None)
            if job is not None:
                self._cleanup_job_files(job)

    def execute_blocking(
        self,
        kind: Literal["convert", "beautify", "beautify_pages", "extract"],
        source_url: str,
        save_outputs: bool,
    ) -> dict:
        future = self._executor.submit(self._execute, kind, source_url, save_outputs)
        return future.result()

    def submit(
        self,
        kind: Literal["convert", "beautify", "beautify_pages", "extract"],
        req: PPTXJobCreateRequest,
    ) -> PPTXJobRecord:
        with self._lock:
            self._cleanup_expired_locked()
            job_id = uuid.uuid4().hex
            created = _now_ts()
            work_dir = self.work_root / job_id
            job = PPTXJobRecord(
                job_id=job_id,
                kind=kind,
                status="queued",
                created_at=created,
                updated_at=created,
                source_url=str(req.ppt_url),
                work_dir=work_dir if kind == "beautify_pages" else None,
                phase="queued" if kind == "beautify_pages" else None,
            )
            self._jobs[job_id] = job

        self._executor.submit(self._run_job, job_id, kind, str(req.ppt_url))
        return job

    def _run_job(
        self,
        job_id: str,
        kind: Literal["convert", "beautify", "beautify_pages", "extract"],
        source_url: str,
    ) -> None:
        with self._lock:
            job = self._jobs.get(job_id)
            if job is None:
                return
            job.status = "running"
            job.updated_at = _now_ts()

        try:
            if kind == "beautify_pages":
                result = self._execute_progressive_beautify(job_id, source_url)
            else:
                result = self._execute(kind, source_url, save_outputs=True)
            with self._lock:
                job = self._jobs.get(job_id)
                if job is None:
                    return
                job.status = "succeeded"
                job.updated_at = _now_ts()
                job.phase = "completed" if kind == "beautify_pages" else job.phase
                job.download_name = result.get("download_name")
                job.download_path = result.get("download_path")
                job.metadata_name = result.get("metadata_name")
                job.metadata_path = result.get("metadata_path")
                job.total_pages = result.get("total_pages")
                job.error = None
        except Exception as exc:
            with self._lock:
                job = self._jobs.get(job_id)
                if job is None:
                    return
                job.status = "failed"
                job.updated_at = _now_ts()
                job.phase = "failed" if kind == "beautify_pages" else job.phase
                job.error = str(exc)

    def _execute(
        self,
        kind: Literal["convert", "beautify", "beautify_pages", "extract"],
        source_url: str,
        save_outputs: bool,
    ) -> dict:
        if kind == "convert":
            archive_name, archive_bytes = convert_ppt_url_to_svg_zip(source_url)
            if save_outputs:
                saved_name, saved_path = save_zip_bytes(archive_name, archive_bytes)
                return {
                    "name": archive_name,
                    "download_name": saved_name,
                    "download_path": saved_path,
                }
            return {"name": archive_name, "data": archive_bytes}

        if kind == "beautify":
            archive_name, archive_bytes, metadata = analyze_ppt_url_for_beautify(source_url)
            total_pages = int(metadata.get("total_pages", 0) or 0)
            if save_outputs:
                saved_name, saved_path = save_zip_bytes(archive_name, archive_bytes)
                metadata_name, metadata_path = save_json_bytes(f"{archive_name}.metadata.json", metadata)
                return {
                    "name": archive_name,
                    "download_name": saved_name,
                    "download_path": saved_path,
                    "metadata_name": metadata_name,
                    "metadata_path": metadata_path,
                    "total_pages": total_pages,
                }
            return {
                "name": archive_name,
                "data": archive_bytes,
                "metadata": metadata,
                "total_pages": total_pages,
            }

        if kind == "extract":
            text_name, text_bytes = extract_ppt_url_to_text(source_url)
            if save_outputs:
                saved_name, saved_path = save_text_bytes(text_name, text_bytes)
                return {
                    "name": text_name,
                    "download_name": saved_name,
                    "download_path": saved_path,
                }
            return {"name": text_name, "data": text_bytes}

        raise ValueError(f"unsupported PPTX job kind: {kind}")

    def _execute_progressive_beautify(self, job_id: str, source_url: str) -> dict:
        with self._lock:
            job = self._jobs.get(job_id)
            if job is None or job.work_dir is None:
                raise ValueError("job not found")
            work_dir = job.work_dir
            job.phase = "rendering"
            job.updated_at = _now_ts()

        def on_page_ready(slide_metadata: dict, png_path: Path) -> None:
            page_index = int(slide_metadata.get("index") or 0)
            total_pages = int(slide_metadata.get("total_pages") or 0)
            page_record = {
                "index": page_index,
                "status": "ready",
                "image_filename": slide_metadata.get("image_filename"),
                "metadata": slide_metadata,
                "image_url": f"/api/v1/pptx/jobs/{job_id}/pages/{page_index}/image",
            }
            if str(slide_metadata.get("vision_image_filename") or "").strip():
                page_record["vision_image_filename"] = slide_metadata.get("vision_image_filename")
                page_record["vision_image_url"] = f"/api/v1/pptx/jobs/{job_id}/pages/{page_index}/vision-image"
            with self._lock:
                current = self._jobs.get(job_id)
                if current is None:
                    return
                existing = [p for p in current.pages if int(p.get("index") or 0) != page_index]
                existing.append(page_record)
                existing.sort(key=lambda p: int(p.get("index") or 0))
                current.pages = existing
                current.ready_pages = len([p for p in existing if p.get("status") == "ready"])
                current.current_page = page_index
                current.total_pages = max(current.total_pages or 0, total_pages, page_index)
                current.phase = "rendering"
                current.updated_at = _now_ts()

        archive_name, archive_bytes, metadata, _png_files, _asset_files = analyze_ppt_url_for_beautify_progressive(
            source_url,
            work_dir,
            on_page_ready=on_page_ready,
        )
        total_pages = int(metadata.get("total_pages", 0) or 0)
        saved_name, saved_path = save_zip_bytes(archive_name, archive_bytes)
        metadata_name, metadata_path = save_json_bytes(f"{archive_name}.metadata.json", metadata)
        with self._lock:
            job = self._jobs.get(job_id)
            if job is not None:
                job.total_pages = total_pages
                job.ready_pages = max(job.ready_pages, len(job.pages))
                job.current_page = total_pages
                job.phase = "packaging"
                job.updated_at = _now_ts()
        return {
            "name": archive_name,
            "download_name": saved_name,
            "download_path": saved_path,
            "metadata_name": metadata_name,
            "metadata_path": metadata_path,
            "total_pages": total_pages,
        }

    def get(self, job_id: str) -> Optional[PPTXJobRecord]:
        with self._lock:
            self._cleanup_expired_locked()
            job = self._jobs.get(job_id)
            if job is None:
                return None
            return PPTXJobRecord(**job.__dict__)

    def delete(self, job_id: str) -> bool:
        with self._lock:
            job = self._jobs.get(job_id)
            if job is None:
                return False
            if job.status in {"queued", "running"}:
                raise ValueError("job is still running")
            self._jobs.pop(job_id, None)

        self._cleanup_job_files(job)
        return True


def _job_to_response(job: JobRecord) -> JobResponse:
    download_url = None
    if job.status == "succeeded" and job.output_path is not None and job.output_path.exists():
        download_url = f"/v1/jobs/{job.job_id}/download"

    return JobResponse(
        job_id=job.job_id,
        status=job.status,
        created_at=int(job.created_at),
        updated_at=int(job.updated_at),
        total=job.total,
        success=job.success_count,
        failed=job.failed_count,
        error=job.error,
        download_url=download_url,
    )


def _build_manager() -> ConversionJobManager:
    work_root = Path(__file__).resolve().parent / "_jobs"
    max_workers = 4
    ttl_seconds = 3600
    convert_timeout_seconds = 900.0

    raw_root = getenv("JOB_WORKDIR")
    raw_workers = getenv("MAX_CONCURRENT_JOBS")
    raw_ttl = getenv("JOB_TTL_SECONDS")
    raw_timeout = getenv("SVG2PPTX_CONVERT_TIMEOUT_SECONDS")

    env_work_root = Path(raw_root) if raw_root else work_root
    env_max_workers = max_workers
    env_ttl_seconds = ttl_seconds
    env_convert_timeout_seconds = convert_timeout_seconds

    if raw_workers:
        try:
            env_max_workers = max(1, int(raw_workers))
        except ValueError:
            env_max_workers = max_workers

    if raw_ttl:
        try:
            env_ttl_seconds = max(60, int(raw_ttl))
        except ValueError:
            env_ttl_seconds = ttl_seconds

    if raw_timeout:
        try:
            env_convert_timeout_seconds = max(0.0, float(raw_timeout))
        except ValueError:
            env_convert_timeout_seconds = convert_timeout_seconds

    return ConversionJobManager(
        work_root=env_work_root,
        max_workers=env_max_workers,
        ttl_seconds=env_ttl_seconds,
        convert_timeout_seconds=env_convert_timeout_seconds if env_convert_timeout_seconds > 0 else None,
    )


def _build_pptx_manager() -> PPTXTaskManager:
    work_root = Path(__file__).resolve().parent / "_pptx_jobs"
    max_workers = 2
    ttl_seconds = 3600

    raw_root = getenv("PPTX_JOB_WORKDIR") or getenv("JOB_WORKDIR")
    raw_workers = getenv("PPTX_MAX_CONCURRENT_JOBS") or getenv("MAX_CONCURRENT_JOBS")
    raw_ttl = getenv("PPTX_JOB_TTL_SECONDS") or getenv("JOB_TTL_SECONDS")

    env_work_root = Path(raw_root) / "pptx" if raw_root else work_root
    env_max_workers = max_workers
    env_ttl_seconds = ttl_seconds

    if raw_workers:
        try:
            env_max_workers = max(1, int(raw_workers))
        except ValueError:
            env_max_workers = max_workers

    if raw_ttl:
        try:
            env_ttl_seconds = max(60, int(raw_ttl))
        except ValueError:
            env_ttl_seconds = ttl_seconds

    return PPTXTaskManager(work_root=env_work_root, max_workers=env_max_workers, ttl_seconds=env_ttl_seconds)


def _resolve_raster_output_size(req: RasterizeRequest) -> tuple[int, int]:
    try:
        src_w, src_h = parse_svg_canvas_size(req.svg)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"invalid svg: {exc}") from exc

    if src_w <= 0 or src_h <= 0:
        src_w, src_h = 1280.0, 720.0

    if req.width is not None and req.height is not None:
        base_w = float(req.width)
        base_h = float(req.height)
    elif req.width is not None:
        base_w = float(req.width)
        base_h = base_w * (src_h / src_w)
    elif req.height is not None:
        base_h = float(req.height)
        base_w = base_h * (src_w / src_h)
    else:
        base_w = float(src_w)
        base_h = float(src_h)

    out_w = int(round(base_w * req.scale))
    out_h = int(round(base_h * req.scale))

    if out_w <= 0 or out_h <= 0:
        raise HTTPException(status_code=400, detail="invalid raster size")
    if out_w > 16384 or out_h > 16384:
        raise HTTPException(status_code=400, detail="raster size too large (max 16384)")
    return out_w, out_h


def _pptx_job_to_response(job: PPTXJobRecord) -> PPTXJobResponse:
    download_url = None
    if job.download_name and job.download_path is not None and job.download_path.exists():
        download_url = f"/api/v1/pptx/jobs/{job.job_id}/download"

    metadata_url = None
    if job.metadata_name and job.metadata_path is not None and job.metadata_path.exists():
        metadata_url = f"/api/v1/pptx/jobs/{job.job_id}/metadata"

    return PPTXJobResponse(
        job_id=job.job_id,
        kind=job.kind,
        status=job.status,
        created_at=int(job.created_at),
        updated_at=int(job.updated_at),
        error=job.error,
        download_url=download_url,
        metadata_url=metadata_url,
        total_pages=job.total_pages,
        ready_pages=job.ready_pages if job.kind == "beautify_pages" else None,
        current_page=job.current_page if job.kind == "beautify_pages" else None,
        phase=job.phase,
        pages=job.pages if job.kind == "beautify_pages" else None,
    )


app = FastAPI(title="2PPTXSVG API", version="1.0.0")
manager = _build_manager()
pptx_manager = _build_pptx_manager()


@app.on_event("shutdown")
def on_shutdown() -> None:
    manager.shutdown()
    pptx_manager.shutdown()


@app.get("/healthz")
def healthz() -> dict:
    try:
        ensure_pptx_dependencies()
    except PPTXConversionError as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc

    return {
        "status": "ok",
        "maxConcurrentJobs": manager.max_workers,
        "pptxMaxConcurrentJobs": pptx_manager.max_workers,
        "jobTtlSeconds": manager.ttl_seconds,
        "pptxJobTtlSeconds": pptx_manager.ttl_seconds,
        "svgToPptxTimeoutSeconds": manager.convert_timeout_seconds,
        "pptx2svg": "ok",
    }


@app.post("/api/v1/convert/ppt-to-svg")
def convert_ppt_to_svg(payload: PPTXConvertRequest) -> Response:
    try:
        result = pptx_manager.execute_blocking("convert", str(payload.ppt_url), save_outputs=payload.url)
    except PPTXConversionError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    if payload.url:
        saved_name = result["download_name"]
        return JSONResponse({"filename": saved_name, "url": f"/downloads/{saved_name}"})

    archive_name = result["name"]
    archive_bytes = result["data"]
    headers = {
        "Content-Disposition": f'attachment; filename="{archive_name}"',
    }
    return Response(content=archive_bytes, media_type="application/zip", headers=headers)


@app.post("/api/v1/analyze/pptx-beautify")
def analyze_pptx_beautify(payload: PPTXConvertRequest) -> Response:
    try:
        result = pptx_manager.execute_blocking("beautify", str(payload.ppt_url), save_outputs=payload.url)
    except PPTXConversionError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    if payload.url:
        saved_name = result["download_name"]
        metadata_name = result["metadata_name"]
        return JSONResponse(
            {
                "filename": saved_name,
                "url": f"/downloads/{saved_name}",
                "metadata_url": f"/downloads/{metadata_name}",
                "total_pages": result.get("total_pages", 0),
            }
        )

    archive_name = result["name"]
    archive_bytes = result["data"]
    headers = {
        "Content-Disposition": f'attachment; filename="{archive_name}"',
    }
    return Response(content=archive_bytes, media_type="application/zip", headers=headers)


@app.post("/api/v1/extract/ppt-text")
def extract_ppt_text(payload: PPTXConvertRequest) -> Response:
    try:
        result = pptx_manager.execute_blocking("extract", str(payload.ppt_url), save_outputs=payload.url)
    except PPTXConversionError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    if payload.url:
        saved_name = result["download_name"]
        return JSONResponse({"filename": saved_name, "url": f"/downloads/{saved_name}"})

    text_name = result["name"]
    text_bytes = result["data"]
    headers = {
        "Content-Disposition": f'attachment; filename="{text_name}"',
    }
    return Response(content=text_bytes, media_type="text/plain; charset=utf-8", headers=headers)


@app.post("/api/v1/pptx/jobs/convert", response_model=PPTXJobResponse, status_code=202)
def create_pptx_convert_job(req: PPTXJobCreateRequest) -> PPTXJobResponse:
    job = pptx_manager.submit("convert", req)
    return _pptx_job_to_response(job)


@app.post("/api/v1/pptx/jobs/beautify", response_model=PPTXJobResponse, status_code=202)
def create_pptx_beautify_job(req: PPTXJobCreateRequest) -> PPTXJobResponse:
    job = pptx_manager.submit("beautify", req)
    return _pptx_job_to_response(job)


@app.post("/api/v1/pptx/jobs/beautify-pages", response_model=PPTXJobResponse, status_code=202)
def create_pptx_beautify_pages_job(req: PPTXJobCreateRequest) -> PPTXJobResponse:
    job = pptx_manager.submit("beautify_pages", req)
    return _pptx_job_to_response(job)


@app.post("/api/v1/pptx/jobs/extract", response_model=PPTXJobResponse, status_code=202)
def create_pptx_extract_job(req: PPTXJobCreateRequest) -> PPTXJobResponse:
    job = pptx_manager.submit("extract", req)
    return _pptx_job_to_response(job)


@app.get("/api/v1/pptx/jobs/{job_id}", response_model=PPTXJobResponse)
def get_pptx_job(job_id: str) -> PPTXJobResponse:
    job = pptx_manager.get(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="job not found or expired")
    return _pptx_job_to_response(job)


@app.get("/api/v1/pptx/jobs/{job_id}/pages/{page_index}/image")
def download_pptx_job_page_image(job_id: str, page_index: int) -> FileResponse:
    job = pptx_manager.get(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="job not found or expired")
    if job.kind != "beautify_pages" or job.work_dir is None:
        raise HTTPException(status_code=409, detail="job does not expose page images")
    if page_index <= 0:
        raise HTTPException(status_code=400, detail="invalid page index")
    page_path = job.work_dir / "png" / f"slide-{page_index:03d}.png"
    if not page_path.is_file():
        raise HTTPException(status_code=404, detail="page image not ready")
    return FileResponse(page_path, media_type="image/png", filename=page_path.name)


@app.get("/api/v1/pptx/jobs/{job_id}/pages/{page_index}/vision-image")
def download_pptx_job_page_vision_image(job_id: str, page_index: int) -> FileResponse:
    job = pptx_manager.get(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="job not found or expired")
    if job.kind != "beautify_pages" or job.work_dir is None:
        raise HTTPException(status_code=409, detail="job does not expose page images")
    if page_index <= 0:
        raise HTTPException(status_code=400, detail="invalid page index")
    page_path = job.work_dir / "png" / f"slide-{page_index:03d}-vision.png"
    if not page_path.is_file():
        raise HTTPException(status_code=404, detail="page vision image not ready")
    return FileResponse(page_path, media_type="image/png", filename=page_path.name)


@app.get("/api/v1/pptx/jobs/{job_id}/assets/{asset_path:path}")
def download_pptx_job_asset(job_id: str, asset_path: str) -> FileResponse:
    job = pptx_manager.get(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="job not found or expired")
    if job.kind != "beautify_pages" or job.work_dir is None:
        raise HTTPException(status_code=409, detail="job does not expose assets")
    safe_parts = [part for part in asset_path.replace("\\", "/").split("/") if part not in {"", ".", ".."}]
    if not safe_parts or safe_parts[0] != "assets":
        raise HTTPException(status_code=400, detail="invalid asset path")
    target = (job.work_dir / Path(*safe_parts)).resolve()
    root = job.work_dir.resolve()
    if root not in target.parents and target != root:
        raise HTTPException(status_code=400, detail="invalid asset path")
    if not target.is_file():
        raise HTTPException(status_code=404, detail="asset not found")
    media_type = "application/octet-stream"
    suffix = target.suffix.lower()
    if suffix in {".png", ".jpg", ".jpeg", ".webp", ".gif"}:
        media_type = f"image/{'jpeg' if suffix in {'.jpg', '.jpeg'} else suffix[1:]}"
    return FileResponse(target, media_type=media_type, filename=target.name)


@app.get("/api/v1/pptx/jobs/{job_id}/download")
def download_pptx_job_result(job_id: str) -> FileResponse:
    job = pptx_manager.get(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="job not found or expired")
    if job.status != "succeeded":
        raise HTTPException(status_code=409, detail=f"job status is {job.status}")
    if job.download_path is None or not job.download_path.exists():
        raise HTTPException(status_code=500, detail="output file missing")

    suffix = job.download_path.suffix.lower()
    if suffix == ".txt":
        media_type = "text/plain; charset=utf-8"
    elif suffix == ".json":
        media_type = "application/json"
    else:
        media_type = "application/zip"
    return FileResponse(job.download_path, media_type=media_type, filename=job.download_path.name)


@app.get("/api/v1/pptx/jobs/{job_id}/metadata")
def download_pptx_job_metadata(job_id: str) -> FileResponse:
    job = pptx_manager.get(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="job not found or expired")
    if job.kind not in {"beautify", "beautify_pages"}:
        raise HTTPException(status_code=409, detail="job does not have metadata")
    if job.metadata_path is None or not job.metadata_path.exists():
        raise HTTPException(status_code=404, detail="metadata file missing")

    return FileResponse(job.metadata_path, media_type="application/json", filename=job.metadata_path.name)


@app.delete("/api/v1/pptx/jobs/{job_id}", status_code=204)
def delete_pptx_job(job_id: str) -> Response:
    try:
        deleted = pptx_manager.delete(job_id)
    except ValueError as exc:
        raise HTTPException(status_code=409, detail=str(exc)) from exc

    if not deleted:
        raise HTTPException(status_code=404, detail="job not found")

    return Response(status_code=204)


@app.get("/downloads/{filename}")
def download_generated_archive(filename: str) -> FileResponse:
    safe_name = filename.rsplit("/", 1)[-1].rsplit("\\", 1)[-1]
    if safe_name != filename:
        raise HTTPException(status_code=400, detail="Invalid filename.")

    archive_path = pptx_settings.downloads_root / safe_name
    if not archive_path.is_file():
        raise HTTPException(status_code=404, detail="File not found.")

    suffix = archive_path.suffix.lower()
    if suffix == ".txt":
        media_type = "text/plain; charset=utf-8"
    elif suffix == ".json":
        media_type = "application/json"
    else:
        media_type = "application/zip"
    return FileResponse(archive_path, media_type=media_type, filename=archive_path.name)


@app.post("/v1/jobs/convert", response_model=JobResponse, status_code=202)
def create_convert_job(req: ConvertRequest) -> JobResponse:
    job = manager.submit(req)
    return _job_to_response(job)


@app.get("/v1/jobs/{job_id}", response_model=JobResponse)
def get_convert_job(job_id: str) -> JobResponse:
    job = manager.get(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="job not found or expired")
    return _job_to_response(job)


@app.get("/v1/jobs/{job_id}/download")
def download_result(job_id: str) -> FileResponse:
    job = manager.get(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="job not found or expired")
    if job.status != "succeeded":
        raise HTTPException(status_code=409, detail=f"job status is {job.status}")
    if job.output_path is None or not job.output_path.exists():
        raise HTTPException(status_code=500, detail="output file missing")

    return FileResponse(
        path=job.output_path,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename=f"{job.title}.pptx",
    )


@app.delete("/v1/jobs/{job_id}", status_code=204)
def delete_job(job_id: str) -> Response:
    try:
        deleted = manager.delete(job_id)
    except ValueError as exc:
        raise HTTPException(status_code=409, detail=str(exc)) from exc

    if not deleted:
        raise HTTPException(status_code=404, detail="job not found")

    return Response(status_code=204)


@app.post("/v1/rasterize")
def rasterize_sync(req: RasterizeRequest) -> Response:
    out_w, out_h = _resolve_raster_output_size(req)
    try:
        png_data = rasterize_svg_to_png(
            svg_content=req.svg,
            output_width=out_w,
            output_height=out_h,
        )
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    return Response(
        content=png_data,
        media_type="image/png",
        headers={
            "X-Output-Width": str(out_w),
            "X-Output-Height": str(out_h),
        },
    )


@app.post("/v1/convert")
def convert_sync(req: ConvertRequest) -> Response:
    try:
        title = _sanitize_file_stem(req.title)
        pptx_data = manager.run_sync_convert(req)
    except TimeoutError as exc:
        raise HTTPException(status_code=504, detail=str(exc)) from exc
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    headers = {
        "Content-Disposition": f"attachment; filename={title}.pptx",
    }
    return Response(
        content=pptx_data,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers=headers,
    )
