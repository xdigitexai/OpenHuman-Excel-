"""Validate model-produced plans into constrained, executable actions.

The model output is untrusted input. Every step is checked against a strict
per-type schema here; anything that fails validation is dropped and reported,
never executed. Workbook content is treated as data, never as instructions.
"""
from __future__ import annotations

import re
from typing import Any

from openhuman_excel.executor import ALLOWED_ACTIONS, DESTRUCTIVE_ACTIONS, ExecutionError


class PlanValidationError(ExecutionError):
    pass


# Characters allowed in a worksheet name per Excel, minus the reserved set.
_INVALID_SHEET_CHARS = re.compile(r"[\[\]:*?/\\]")
_CELL_RE = re.compile(r"^\$?[A-Za-z]{1,3}\$?[0-9]{1,7}$")
_RANGE_RE = re.compile(r"^\$?[A-Za-z]{1,3}\$?[0-9]{1,7}:\$?[A-Za-z]{1,3}\$?[0-9]{1,7}$")
# Formula guard: block DDE/external-object payloads and network lookups.
# A formula such as =cmd|'/c calc'!A1 or =WEBSERVICE("http://...") must never
# reach Excel. Plain in-workbook references (Sheet2!A1) remain allowed.
_FORMULA_DANGEROUS = re.compile(
    r"\||\b(WEBSERVICE|FILTERXML|CALL|REGISTER)\b|://", re.IGNORECASE)


def sanitize_sheet_name(name: Any) -> str:
    if not isinstance(name, str):
        raise PlanValidationError("Sheet name must be a string.")
    cleaned = _INVALID_SHEET_CHARS.sub("", name.strip())[:31]
    if not cleaned:
        raise PlanValidationError("Sheet name is empty after sanitization.")
    return cleaned


def validate_formula(value: Any) -> str:
    if not isinstance(value, str) or not value.startswith("="):
        raise PlanValidationError("Formula value must be a string beginning with '='.")
    if _FORMULA_DANGEROUS.search(value):
        raise PlanValidationError(
            "Formula rejected: DDE/external data/network references are not allowed.")
    return value


def validate_step(step: Any) -> dict[str, Any]:
    """Validate a single plan step; returns a sanitized action dict."""
    if not isinstance(step, dict):
        raise PlanValidationError("Step must be an object.")
    kind = step.get("type")
    if kind not in ALLOWED_ACTIONS:
        raise PlanValidationError(f"Unsupported action type: {kind!r}")

    action: dict[str, Any] = {"type": kind}
    if kind in {"set_cell", "write_formula"}:
        action["sheet"] = sanitize_sheet_name(step.get("sheet"))
        cell = step.get("cell")
        if not isinstance(cell, str) or not _CELL_RE.match(cell.strip()):
            raise PlanValidationError(f"Invalid cell reference: {cell!r}")
        action["cell"] = cell.strip().replace("$", "")
        value = step.get("value")
        if kind == "write_formula":
            action["value"] = validate_formula(value)
        else:
            if not isinstance(value, (str, int, float, bool)) or value is None:
                raise PlanValidationError("Cell value must be a string, number or boolean.")
            action["value"] = value
    elif kind == "rename_sheet":
        action["sheet"] = sanitize_sheet_name(step.get("sheet"))
        action["new_name"] = sanitize_sheet_name(step.get("new_name"))
        if action["new_name"].lower() == action["sheet"].lower():
            raise PlanValidationError("Rename target equals the current name.")
    elif kind == "add_sheet":
        action["new_name"] = sanitize_sheet_name(step.get("new_name"))
    elif kind == "remove_duplicates":
        action["sheet"] = sanitize_sheet_name(step.get("sheet"))
        cell_range = step.get("range")
        if not isinstance(cell_range, str) or not _RANGE_RE.match(cell_range.strip()):
            raise PlanValidationError(f"Invalid range: {cell_range!r}")
        action["range"] = cell_range.strip().replace("$", "")
    elif kind == "set_header_style":
        action["sheet"] = sanitize_sheet_name(step.get("sheet"))
        row = step.get("row", 1)
        if isinstance(row, bool) or not isinstance(row, int) or not 1 <= row <= 1_048_576:
            raise PlanValidationError(f"Invalid header row: {row!r}")
        action["row"] = row

    if kind in DESTRUCTIVE_ACTIONS:
        action["requires_approval"] = True
    return action


def validate_plan(plan: Any) -> tuple[list[dict[str, Any]], list[str]]:
    """Validate a whole plan. Returns (actions, issues). Invalid steps are skipped."""
    if not isinstance(plan, dict):
        raise PlanValidationError("Plan must be a JSON object.")
    raw_steps = plan.get("steps")
    if raw_steps is None:
        return [], []
    if not isinstance(raw_steps, list):
        raise PlanValidationError("Plan 'steps' must be an array.")
    actions: list[dict[str, Any]] = []
    issues: list[str] = []
    for index, step in enumerate(raw_steps):
        try:
            actions.append(validate_step(step))
        except PlanValidationError as exc:
            issues.append(f"Step {index + 1} rejected: {exc}")
    return actions, issues
