from __future__ import annotations

import json
import sqlite3
import threading
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


class Ledger:
    """Persistent task/action ledger; uniqueness prevents accidental repeat execution."""

    def __init__(self, path: Path):
        path.parent.mkdir(parents=True, exist_ok=True)
        self.path = path
        self._lock = threading.RLock()
        with self._connect() as db:
            db.executescript("""
                CREATE TABLE IF NOT EXISTS tasks (
                    task_id TEXT PRIMARY KEY, instruction TEXT NOT NULL,
                    workbook TEXT NOT NULL, state TEXT NOT NULL,
                    created_at TEXT NOT NULL, updated_at TEXT NOT NULL
                );
                CREATE TABLE IF NOT EXISTS actions (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    task_id TEXT NOT NULL, action_key TEXT NOT NULL,
                    action_type TEXT NOT NULL, target TEXT NOT NULL,
                    payload TEXT, status TEXT NOT NULL, created_at TEXT NOT NULL,
                    UNIQUE(task_id, action_key)
                );
                CREATE TABLE IF NOT EXISTS events (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    task_id TEXT, event TEXT NOT NULL, details TEXT,
                    created_at TEXT NOT NULL
                );
            """)

    def _connect(self):
        db = sqlite3.connect(self.path, timeout=20)
        db.row_factory = sqlite3.Row
        return db

    @staticmethod
    def now() -> str:
        return datetime.now(timezone.utc).isoformat()

    def create_task(self, task_id: str, instruction: str, workbook: str) -> None:
        now = self.now()
        with self._lock, self._connect() as db:
            db.execute("INSERT INTO tasks VALUES (?, ?, ?, 'queued', ?, ?)",
                       (task_id, instruction, workbook, now, now))

    def set_task_state(self, task_id: str, state: str) -> None:
        with self._lock, self._connect() as db:
            db.execute("UPDATE tasks SET state=?, updated_at=? WHERE task_id=?",
                       (state, self.now(), task_id))

    def record_event(self, event: str, task_id: str | None = None,
                     details: dict[str, Any] | None = None) -> None:
        with self._lock, self._connect() as db:
            db.execute("INSERT INTO events(task_id,event,details,created_at) VALUES(?,?,?,?)",
                       (task_id, event, json.dumps(details or {}, default=str), self.now()))

    def begin_action(self, task_id: str, action_key: str, action_type: str,
                     target: str, payload: Any = None) -> bool:
        """Return False if this exact action was already recorded for this task."""
        with self._lock, self._connect() as db:
            try:
                db.execute("""INSERT INTO actions
                    (task_id,action_key,action_type,target,payload,status,created_at)
                    VALUES(?,?,?,?,?,'started',?)""",
                    (task_id, action_key, action_type, target,
                     json.dumps(payload, default=str), self.now()))
                return True
            except sqlite3.IntegrityError:
                return False

    def finish_action(self, task_id: str, action_key: str, status: str) -> None:
        with self._lock, self._connect() as db:
            db.execute("UPDATE actions SET status=? WHERE task_id=? AND action_key=?",
                       (status, task_id, action_key))

    def actions_for_task(self, task_id: str) -> list[dict[str, Any]]:
        with self._connect() as db:
            return [dict(row) for row in db.execute(
                "SELECT * FROM actions WHERE task_id=? ORDER BY id", (task_id,))]
