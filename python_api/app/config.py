from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path


def env_bool(name: str, default: bool) -> bool:
    value = os.getenv(name)
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "on"}


@dataclass(slots=True)
class Settings:
    app_name: str = "ppt-to-svg-api"
    work_root: Path = Path(os.getenv("WORK_ROOT", "/tmp/ppt-to-svg"))
    downloads_subdir: str = os.getenv("DOWNLOADS_SUBDIR", "downloads")
    download_timeout_seconds: int = int(os.getenv("DOWNLOAD_TIMEOUT_SECONDS", "120"))
    command_timeout_seconds: int = int(os.getenv("COMMAND_TIMEOUT_SECONDS", "240"))
    max_download_mb: int = int(os.getenv("MAX_DOWNLOAD_MB", "100"))
    java_command: str = os.getenv("JAVA_COMMAND", "java")
    java_renderer_jar: Path = Path(
        os.getenv(
            "JAVA_RENDERER_JAR",
            "/opt/pptx2svg-renderer/pptx2svg-renderer.jar",
        )
    )
    svg_text_as_shapes: bool = env_bool("SVG_TEXT_AS_SHAPES", True)

    @property
    def max_download_bytes(self) -> int:
        return self.max_download_mb * 1024 * 1024

    @property
    def downloads_root(self) -> Path:
        return self.work_root / self.downloads_subdir


settings = Settings()
