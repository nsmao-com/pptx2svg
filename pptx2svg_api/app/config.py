# input: Environment variables that tune runtime paths, limits, and commands.
# output: A single Settings object used by the API and converter modules.
# pos: Local configuration boundary for the PPT/PPTX->SVG/PNG subsystem.
# update: When this file changes, update this header and pptx2svg_api/app/README.md.

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path


@dataclass(slots=True)
class Settings:
    app_name: str = "ppt-to-svg-api"
    work_root: Path = Path(os.getenv("WORK_ROOT", "/tmp/ppt-to-svg"))
    downloads_subdir: str = os.getenv("DOWNLOADS_SUBDIR", "downloads")
    download_timeout_seconds: int = int(os.getenv("DOWNLOAD_TIMEOUT_SECONDS", "120"))
    command_timeout_seconds: int = int(os.getenv("COMMAND_TIMEOUT_SECONDS", "240"))
    max_download_mb: int = int(os.getenv("MAX_DOWNLOAD_MB", "100"))
    libreoffice_command: str = os.getenv("LIBREOFFICE_COMMAND", "soffice")
    mupdf_command: str = os.getenv("MUPDF_COMMAND", "mutool")
    libreoffice_pdf_filter: str = os.getenv(
        "LIBREOFFICE_PDF_FILTER",
        "pdf:impress_pdf_Export",
    )
    libreoffice_svg_filter: str = os.getenv(
        "LIBREOFFICE_SVG_FILTER",
        "svg:impress_svg_Export",
    )
    png_render_dpi: int = int(os.getenv("PPTX_PNG_RENDER_DPI", "144"))
    svg_fixed_font_family: str = os.getenv(
        "SVG_FIXED_FONT_FAMILY",
        "Noto Sans CJK SC",
    )
    strip_svg_embedded_fonts: bool = (
        os.getenv("STRIP_SVG_EMBEDDED_FONTS", "true").lower()
        in {"1", "true", "yes", "on"}
    )
    aliyun_oss_endpoint: str = os.getenv("ALIYUN_OSS_ENDPOINT", "")
    aliyun_oss_access_key_id: str = os.getenv("ALIYUN_OSS_ACCESS_KEY_ID", "")
    aliyun_oss_access_key_secret: str = os.getenv("ALIYUN_OSS_ACCESS_KEY_SECRET", "")
    aliyun_oss_bucket_name: str = os.getenv("ALIYUN_OSS_BUCKET_NAME", "")
    aliyun_oss_base_url: str = os.getenv("ALIYUN_OSS_BASE_URL", "").rstrip("/")
    aliyun_oss_prefix: str = os.getenv("ALIYUN_OSS_PREFIX", "pptx2svg/assets").strip("/")
    upload_svg_images_to_oss: bool = (
        os.getenv("UPLOAD_SVG_IMAGES_TO_OSS", "false").lower()
        in {"1", "true", "yes", "on"}
    )

    @property
    def aliyun_oss_enabled(self) -> bool:
        return (
            self.upload_svg_images_to_oss
            and bool(self.aliyun_oss_endpoint)
            and bool(self.aliyun_oss_access_key_id)
            and bool(self.aliyun_oss_access_key_secret)
            and bool(self.aliyun_oss_bucket_name)
            and bool(self.aliyun_oss_base_url)
        )

    @property
    def max_download_bytes(self) -> int:
        return self.max_download_mb * 1024 * 1024

    @property
    def downloads_root(self) -> Path:
        return self.work_root / self.downloads_subdir


settings = Settings()
