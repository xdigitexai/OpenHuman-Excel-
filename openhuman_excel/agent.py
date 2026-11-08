from __future__ import annotations

import json
import uuid
from pathlib import Path
from typing import Callable

from openhuman_excel.config import settings
from openhuman_excel.ledger import Ledger
from openhuman_excel.provider import OpenAICompatibleProvider
from openhuman_excel.workbook import ExcelCom, inspect_with_openpyxl, make_working_copy

SYSTEM_PROMPT = """You are OpenHuman Excel, an accurate Excel work assistant.
Analyze the user's specific request and workbook structure. Never invent extra work.
Do not repeat an operation already completed. If the task is ambiguous, ask a question.
Return concise JSON with keys: summary, steps (array of concrete steps), risks (array).
Do not claim to have edited a workbook; this stage only plans and inspects.
"""


class AgentService:
    def __init__(self, log: Callable[[str], None] = print):
        self.log = log
        self.ledger = Ledger(settings.database_path)
        self.provider = OpenAICompatibleProvider(settings.base_url, settings.api_key,
                                                  settings.model, settings.request_timeout)
        self._stop = False
        self._pause = False
        self.current_task_id: str | None = None

    def pause(self):
        self._pause = True

    def resume(self):
        self._pause = False

    def stop(self):
        self._stop = True
        self._pause = False

    def _checkpoint_control(self):
        import time
        while self._pause and not self._stop:
            time.sleep(0.25)
        if self._stop:
            raise InterruptedError("Stopped by user.")

    def inspect_and_plan(self, source: Path, instruction: str) -> dict:
        self._stop = False
        task_id = str(uuid.uuid4())
        working_copy = make_working_copy(source, settings.checkpoint_dir)
        snapshot = inspect_with_openpyxl(working_copy)
        self.current_task_id = task_id
        self.ledger.create_task(task_id, instruction, str(working_copy))
        self.ledger.record_event("workbook_inspected", task_id, snapshot)
        self.log(f"Working copy created: {working_copy}")
        self.log(f"Workbook sheets: {', '.join(snapshot['sheet_names']) or '(none)'}")
        prompt = (f"User task: {instruction}\nWorkbook structure:\n"
                  f"{json.dumps(snapshot, ensure_ascii=False)}")
        plan_text = self.provider.complete(SYSTEM_PROMPT, prompt)
        self.ledger.record_event("plan_received", task_id, {"plan": plan_text})
        try:
            plan = json.loads(plan_text)
        except json.JSONDecodeError:
            plan = {"summary": "Review the model's proposed plan", "steps": [],
                    "risks": ["The model did not return valid JSON; no edits were made."],
                    "raw_response": plan_text}
        self.ledger.set_task_state(task_id, "planned")
        return {"task_id": task_id, "working_copy": str(working_copy),
                "snapshot": snapshot, "plan": plan}
