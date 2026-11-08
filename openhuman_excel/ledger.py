from __future__ import annotations

import json
import sqlite3
import threading
from collections.abc import Callable
from datetime import UTC, datetime
from pathlib import Path
from typing import Any


class Ledger:
    """Persistent task/action ledger; uniqueness prevents accidental repeat execution."""

    def __init__(self, path: Path):
        path.parent.mkdir(parents=True, exist_ok=True)
        self.path = path
        self._lock = threading.RLock()
        db = self._connect()
        try:
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
            db.commit()
        finally:
            db.close()

    def _connect(self):
        db = sqlite3.connect(self.path, timeout=20)
        db.row_factory = sqlite3.Row
        return db

    def _run(self, work: Callable[[sqlite3.Connection], Any]) -> Any:
        """Run a unit of work on its own connection and always close it."""
        with self._lock:
            db = self._connect()
            try:
                result = work(db)
                db.commit()
                return result
            except Exception:
                db.rollback()
                raise
            finally:
                db.close()

    @staticmethod
    def now() -> str:
        return datetime.now(UTC).isoformat()

    def create_task(self, task_id: str, instruction: str, workbook: str) -> None:
        now = self.now()
        def work(db):
            db.execute("INSERT INTO tasks VALUES (?, ?, ?, 'queued', ?, ?)",
                       (task_id, instruction, workbook, now, now))
        self._run(work)

    def set_task_state(self, task_id: str, state: str) -> None:
        def work(db):
            db.execute("UPDATE tasks SET state=?, updated_at=? WHERE task_id=?",
                       (state, self.now(), task_id))
        self._run(work)

    def get_task_state(self, task_id: str) -> str | None:
        def work(db):
            row = db.execute("SELECT state FROM tasks WHERE task_id=?", (task_id,)).fetchone()
            return row["state"] if row else None
        return self._run(work)

    def record_event(self, event: str, task_id: str | None = None,
                     details: dict[str, Any] | None = None) -> None:
        def work(db):
            db.execute("INSERT INTO events(task_id,event,details,created_at) VALUES(?,?,?,?)",
                       (task_id, event, json.dumps(details or {}, default=str), self.now()))
        self._run(work)

    def begin_action(self, task_id: str, action_key: str, action_type: str,
                     target: str, payload: Any = None) -> bool:
        """Return False only if this exact action already completed for this task.

        Actions left in a non-completed state (started before a crash, failed,
        interrupted) are retryable: the workbook state for them is unknown, so
        re-running them is the only safe recovery.
        """
        def work(db):
            existing = db.execute(
                "SELECT status FROM actions WHERE task_id=? AND action_key=?",
                (task_id, action_key)).fetchone()
            if existing is not None and existing["status"] == "completed":
                return False
            if existing is None:
                db.execute("""INSERT INTO actions
                    (task_id,action_key,action_type,target,payload,status,created_at)
                    VALUES(?,?,?,?,?,'started',?)""",
                    (task_id, action_key, action_type, target,
                     json.dumps(payload, default=str), self.now()))
            else:
                db.execute("""UPDATE actions SET status='started'
                    WHERE task_id=? AND action_key=?""", (task_id, action_key))
            return True
        return self._run(work)

    def finish_action(self, task_id: str, action_key: str, status: str) -> None:
        def work(db):
            db.execute("UPDATE actions SET status=? WHERE task_id=? AND action_key=?",
                       (status, task_id, action_key))
        self._run(work)

    def actions_for_task(self, task_id: str) -> list[dict[str, Any]]:
        def work(db):
            return [dict(row) for row in db.execute(
                "SELECT * FROM actions WHERE task_id=? ORDER BY id", (task_id,))]
        return self._run(work)
