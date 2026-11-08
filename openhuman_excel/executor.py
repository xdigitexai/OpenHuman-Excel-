from __future__ import annotations

import hashlib
import time
from pathlib import Path
from typing import Any, Callable

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils.cell import range_boundaries

from openhuman_excel.ledger import Ledger


class ExecutionError(RuntimeError):
    pass


ALLOWED_ACTIONS = {"set_cell", "write_formula", "rename_sheet", "add_sheet",
                   "remove_duplicates", "set_header_style"}


def _key(action: dict[str, Any]) -> str:
    import json
    raw = json.dumps(action, sort_keys=True, ensure_ascii=False, default=str)
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


class WorkbookExecutor:
    """Constrained workbook executor. Every action is logged and duplicate guarded."""

    def __init__(self, ledger: Ledger, log: Callable[[str], None] = print,
                 pace_seconds: float = 0.35):
        self.ledger = ledger
        self.log = log
        self.pace_seconds = max(0.0, min(float(pace_seconds), 5.0))
        self.paused = False
        self.stopped = False

    def pause(self):
        self.paused = True

    def resume(self):
        self.paused = False

    def stop(self):
        self.stopped = True
        self.paused = False

    def _checkpoint(self):
        while self.paused and not self.stopped:
            time.sleep(0.2)
        if self.stopped:
            raise InterruptedError("Execution stopped by user.")

    def execute(self, path: Path, task_id: str, actions: list[dict[str, Any]],
                approve: Callable[[dict[str, Any]], bool]) -> int:
        keep_vba = path.suffix.lower() == ".xlsm"
        wb = load_workbook(path, keep_vba=keep_vba)
        completed = 0
        try:
            for index, action in enumerate(actions):
                self._checkpoint()
                kind = action.get("type")
                if kind not in ALLOWED_ACTIONS:
                    raise ExecutionError(f"Unsupported action type at step {index + 1}: {kind!r}")
                if action.get("requires_approval") and not approve(action):
                    self.log(f"Skipped pending approval: {kind}")
                    continue
                key = action.get("idempotency_key") or _key(action)
                target = str(action.get("target", "workbook"))
                if not self.ledger.begin_action(task_id, key, kind, target, action):
                    self.log(f"Skipped already-recorded action: {kind} → {target}")
                    continue
                try:
                    self._apply(wb, action)
                    wb.save(path)
                    self.ledger.finish_action(task_id, key, "completed")
                    completed += 1
                    self.log(f"Completed {kind}: {target}")
                except Exception:
                    self.ledger.finish_action(task_id, key, "failed")
                    raise
                if self.pace_seconds:
                    time.sleep(self.pace_seconds)
            return completed
        finally:
            wb.close()

    @staticmethod
    def _apply(wb, action: dict[str, Any]) -> None:
        kind = action["type"]
        if kind in {"set_cell", "write_formula"}:
            sheet = action.get("sheet")
            cell = action.get("cell")
            if sheet not in wb.sheetnames or not isinstance(cell, str):
                raise ExecutionError("Action must reference an existing sheet and cell.")
            value = action.get("value")
            if kind == "write_formula" and (not isinstance(value, str) or not value.startswith("=")):
                raise ExecutionError("Formula action value must begin with '='.")
            wb[sheet][cell] = value
        elif kind == "rename_sheet":
            old, new = action.get("sheet"), action.get("new_name")
            if old not in wb.sheetnames or not isinstance(new, str) or not new.strip():
                raise ExecutionError("Invalid worksheet rename.")
            if new in wb.sheetnames:
                raise ExecutionError(f"Worksheet already exists: {new}")
            wb[old].title = new.strip()[:31]
        elif kind == "add_sheet":
            name = action.get("new_name")
            if not isinstance(name, str) or not name.strip() or name in wb.sheetnames:
                raise ExecutionError("Invalid or duplicate worksheet name.")
            wb.create_sheet(name.strip()[:31])
        elif kind == "remove_duplicates":
            sheet_name, cell_range = action.get("sheet"), action.get("range")
            if sheet_name not in wb.sheetnames or not isinstance(cell_range, str):
                raise ExecutionError("Deduplication requires an existing sheet and range.")
            ws = wb[sheet_name]
            min_col, min_row, max_col, max_row = range_boundaries(cell_range)
            seen, remove = set(), []
            for row in range(min_row, max_row + 1):
                values = tuple(ws.cell(row, col).value for col in range(min_col, max_col + 1))
                if values in seen:
                    remove.append(row)
                else:
                    seen.add(values)
            for row in reversed(remove):
                ws.delete_rows(row, 1)
        elif kind == "set_header_style":
            sheet_name, row = action.get("sheet"), int(action.get("row", 1))
            if sheet_name not in wb.sheetnames or row < 1:
                raise ExecutionError("Invalid header style target.")
            ws = wb[sheet_name]
            for cell in ws[row]:
                if cell.value is not None:
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.fill = PatternFill("solid", fgColor="244A73")
