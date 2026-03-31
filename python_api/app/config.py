from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path


@dataclass(slots=True)
class Settings:
    app_name: str = "ppt-to-svg-api"
    work_root: Path = Path(os.getenv("WORK_ROOT", "/tmp/ppt-to-svg"))
    download_timeout_seconds: int = int(os.getenv("DOWNLOAD_TIMEOUT_SECONDS", "120"))
    command_timeout_seconds: int = int(os.getenv("COMMAND_TIMEOUT_SECONDS", "240"))
    max_download_mb: int = int(os.getenv("MAX_DOWNLOAD_MB", "100"))
    libreoffice_start_timeout_seconds: int = int(
        os.getenv("LIBREOFFICE_START_TIMEOUT_SECONDS", "45")
    )
    libreoffice_program_dir: Path = Path(
        os.getenv("LIBREOFFICE_PROGRAM_DIR", "/opt/libreoffice26.2/program")
    )

    @property
    def max_download_bytes(self) -> int:
        return self.max_download_mb * 1024 * 1024


settings = Settings()
