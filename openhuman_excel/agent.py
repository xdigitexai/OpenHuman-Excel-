from __future__ import annotations

import json
import time
import uuid
from collections.abc import Callable
from pathlib import Path

from openhuman_excel.config import settings
from openhuman_excel.executor import WorkbookExecutor
from openhuman_excel.ledger import Ledger
from openhuman_excel.planning import PlanValidationError, validate_plan
from openhuman_excel.provider import OpenAICompatibleProvider
from openhuman_excel.workbook import (
    inspect_with_openpyxl,
    make_working_copy,
    verify_formulas_in_excel,
)

SYSTEM_PROMPT = """You are OpenHuman Excel, an accurate Excel work assistant.
Analyze the user's specific request and workbook structure. Never invent extra work.
Do not repeat an operation already completed. If the task is ambiguous, ask a question.
Return concise JSON with keys: summary, steps (array of concrete steps), risks (array).
Each step must be an object with a "type" chosen from: set_cell, write_formula,
rename_sheet, add_sheet, remove_duplicates, set_header_style, plus the fields that
type needs (sheet, cell, value, range, new_name, row).
Do not claim to have edited a workbook; this stage only plans and inspects.
"""


class SessionBudgetExceeded(RuntimeError):
    pass


class AgentService:
    def __init__(self, log: Callable[[str], None] = print,
                 ledger: Ledger | None = None,
                 provider: OpenAICompatibleProvider | None = None):
        self.log = log
        self.ledger = ledger if ledger is not None else Ledger(settings.database_path)
        self.provider = provider if provider is not None else OpenAICompatibleProvider(
            settings.base_url, settings.api_key, settings.model, settings.request_timeout)
        self._stop = False
        self._pause = False
        self.current_task_id: str | None = None
        self.executor: WorkbookExecutor | None = None
        self._session_started_monotonic: float | None = None

    # ---- controls -------------------------------------------------------

    def pause(self):
        self._pause = True
        if self.executor is not None:
            self.executor.pause()

    def resume(self):
        self._pause = False
        if self.executor is not None:
            self.executor.resume()

    def stop(self):
        self._stop = True
        self._pause = False
        if self.executor is not None:
            self.executor.stop()

    def reset_controls(self):
        """Clear stop/pause before a deliberate new run (GUI Start / Apply)."""
        self._stop = False
        self._pause = False

    def _checkpoint_control(self):
        import time
        while self._pause and not self._stop:
            time.sleep(0.25)
        if self._stop:
            raise InterruptedError("Stopped by user.")

    def _check_budget(self):
        """Stop the session when the configured daily work allowance is used up.

        The budget only limits continuous agent work time; the agent never
        invents extra tasks to fill the window.
        """
        if self._session_started_monotonic is None:
            self._session_started_monotonic = time.monotonic()
            return
        elapsed_hours = (time.monotonic() - self._session_started_monotonic) / 3600.0
        if elapsed_hours >= settings.daily_limit_hours:
            raise SessionBudgetExceeded(
                f"Daily agent-work limit reached ({settings.daily_limit_hours} h). "
                "The session stopped instead of inventing more work.")

    # ---- planning -------------------------------------------------------

    def inspect_and_plan(self, source: Path, instruction: str) -> dict:
        self._check_budget()
        task_id = str(uuid.uuid4())
        working_copy = make_working_copy(source, settings.checkpoint_dir)
        snapshot = inspect_with_openpyxl(working_copy)
        self.current_task_id = task_id
        self.ledger.create_task(task_id, instruction, str(working_copy))
        self.ledger.record_event("workbook_inspected", task_id, snapshot)
        self.log(f"Working copy created: {working_copy}")
        self.log(f"Workbook sheets: {', '.join(snapshot['sheet_names']) or '(none)'}")
        self._checkpoint_control()
        prompt = (f"User task: {instruction}\nWorkbook structure:\n"
                  f"{json.dumps(snapshot, ensure_ascii=False)}")
        plan_text = self.provider.complete(SYSTEM_PROMPT, prompt)
        self.ledger.record_event("plan_received", task_id, {"plan": plan_text})
        self._checkpoint_control()
        try:
            plan = json.loads(plan_text)
        except json.JSONDecodeError:
            plan = {"summary": "Review the model's proposed plan", "steps": [],
                    "risks": ["The model did not return valid JSON; no edits were made."],
                    "raw_response": plan_text}
        self.ledger.set_task_state(task_id, "planned")
        return {"task_id": task_id, "working_copy": str(working_copy),
                "snapshot": snapshot, "plan": plan}

    # ---- execution ------------------------------------------------------

    def run_plan(self, plan_result: dict,
                 approve: Callable[[dict], bool] | None = None) -> dict:
        """Validate a plan and execute it against the working copy.

        Destructive actions call `approve` (or are skipped when no approver is
        wired). Every action goes through the ledger's duplicate guard, so
        re-running the same plan never repeats completed work, and actions
        interrupted by a crash are retried rather than lost.
        """
        if approve is None:
            approve = lambda action: False
        task_id = plan_result["task_id"]
        working_copy = Path(plan_result["working_copy"])
        plan = plan_result.get("plan", {})
        self._check_budget()
        try:
            actions, issues = validate_plan(plan)
        except PlanValidationError as exc:
            self.ledger.set_task_state(task_id, "failed")
            self.ledger.record_event("plan_validation_failed", task_id, {"error": str(exc)})
            raise
        for issue in issues:
            self.log(f"Plan issue: {issue}")
        if issues:
            self.ledger.record_event("plan_validation_issues", task_id, {"issues": issues})
        if not actions:
            self.ledger.set_task_state(task_id, "completed")
            return {"task_id": task_id, "executed": 0, "issues": issues,
                    "message": "No executable steps in this plan; nothing was edited."}

        self.ledger.set_task_state(task_id, "running")
        self.executor = WorkbookExecutor(self.ledger, log=self.log)
        completed = 0
        try:
            completed = self.executor.execute(working_copy, task_id, actions, approve)
        except InterruptedError:
            self.ledger.set_task_state(task_id, "paused")
            self.ledger.record_event("execution_stopped", task_id, {"completed": completed})
            raise
        except Exception as exc:
            self.ledger.set_task_state(task_id, "failed")
            self.ledger.record_event("execution_failed", task_id,
                                     {"error": f"{type(exc).__name__}: {exc}",
                                      "completed": completed})
            raise
        finally:
            self.executor = None
        self.ledger.set_task_state(task_id, "completed")
        self.ledger.record_event("execution_completed", task_id, {"completed": completed})

        verification = self._verify_in_excel(working_copy, actions)
        return {"task_id": task_id, "executed": completed, "issues": issues,
                "verification": verification}

    def _verify_in_excel(self, working_copy: Path, actions: list[dict]) -> dict | None:
        """Verify formula results with real Excel when it is available."""
        formulas = [a for a in actions if a.get("type") == "write_formula"]
        if not formulas:
            return None
        checks: dict[str, list[str]] = {}
        for action in formulas:
            checks.setdefault(action["sheet"], []).append(action["cell"])
        try:
            result = verify_formulas_in_excel(working_copy, checks)
        except Exception as exc:
            self.log(f"Excel verification unavailable: {type(exc).__name__}: {exc}")
            return None
        if result.get("ok"):
            self.log("Excel verification (recalculated values): "
                     f"{json.dumps(result['values'], ensure_ascii=False, default=str)}")
        else:
            self.log(f"Excel verification failed: {result.get('error')}")
        return result
