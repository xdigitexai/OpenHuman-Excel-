from __future__ import annotations

from dataclasses import dataclass, field
from enum import StrEnum
from pathlib import Path
from typing import Any


class SessionState(StrEnum):
    IDLE = "idle"
    RUNNING = "running"
    PAUSED = "paused"
    STOPPING = "stopping"
    COMPLETED = "completed"
    FAILED = "failed"


@dataclass
class WorkbookSnapshot:
    path: Path
    sheets: list[dict[str, Any]] = field(default_factory=list)
    active_sheet: str | None = None


@dataclass
class TaskRequest:
    instruction: str
    workbook_path: Path
    task_id: str


@dataclass
class Action:
    action_type: str
    target: str
    value: Any = None
    rationale: str = ""
    idempotency_key: str = ""
    requires_approval: bool = False
