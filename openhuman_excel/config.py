from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()


@dataclass
class Settings:
    base_url: str = os.getenv("OPENH_EXCEL_BASE_URL", "https://api.deepseek.com")
    api_key: str = os.getenv("OPENH_EXCEL_API_KEY", "")
    model: str = os.getenv("OPENH_EXCEL_MODEL", "deepseek-chat")
    request_timeout: int = int(os.getenv("OPENH_EXCEL_REQUEST_TIMEOUT", "90"))
    daily_limit_hours: float = min(10.0, max(0.25, float(os.getenv("OPENH_EXCEL_DAILY_LIMIT_HOURS", "10"))))
    data_dir: Path = Path.home() / ".openhuman-excel"

    @property
    def database_path(self) -> Path:
        return self.data_dir / "agent.sqlite3"

    @property
    def checkpoint_dir(self) -> Path:
        return self.data_dir / "checkpoints"


settings = Settings()
