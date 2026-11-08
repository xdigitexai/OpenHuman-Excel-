from __future__ import annotations

import shutil
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any

SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm"}


def make_working_copy(source: Path, destination_dir: Path) -> Path:
    source = source.expanduser().resolve()
    if not source.is_file() or source.suffix.lower() not in SUPPORTED_EXTENSIONS:
        raise ValueError("Choose an existing .xlsx or .xlsm workbook. The legacy .xls format is not yet supported.")
    destination_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    # A short uuid suffix keeps copies unique even when several are made in
    # the same second (for example, a retry right after a crash).
    target = destination_dir / f"{source.stem}.openh-{stamp}-{uuid.uuid4().hex[:6]}{source.suffix}"
    shutil.copy2(source, target)
    return target


def inspect_with_openpyxl(path: Path) -> dict[str, Any]:
    from openpyxl import load_workbook

    keep_vba = path.suffix.lower() == ".xlsm"
    wb = load_workbook(path, read_only=True, data_only=False, keep_vba=keep_vba)
    try:
        sheets = []
        for ws in wb.worksheets:
            sheets.append({
                "name": ws.title,
                "max_row": ws.max_row,
                "max_column": ws.max_column,
                "state": ws.sheet_state,
            })
        return {"path": str(path), "sheet_names": wb.sheetnames, "sheets": sheets}
    finally:
        wb.close()


class ExcelCom:
    """Visible Microsoft Excel automation adapter. All COM work stays on its worker thread."""

    def __init__(self, visible: bool = True):
        self.visible = visible
        self.app = None
        self.workbook = None
        self._com_initialized = False

    def open(self, path: Path) -> dict[str, Any]:
        if __import__("os").name != "nt":
            raise RuntimeError("Microsoft Excel COM automation is available on Windows only.")
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        self._com_initialized = True
        try:
            self.app = win32com.client.DispatchEx("Excel.Application")
            self.app.Visible = self.visible
            self.app.DisplayAlerts = False
            self.workbook = self.app.Workbooks.Open(str(path.resolve()), ReadOnly=False)
        except Exception:
            self.close(save=False)
            raise
        sheets = []
        for sheet in self.workbook.Worksheets:
            used = sheet.UsedRange
            sheets.append({"name": sheet.Name, "rows": used.Rows.Count,
                           "columns": used.Columns.Count})
        return {"sheets": sheets, "active_sheet": self.app.ActiveSheet.Name}

    def close(self, save: bool = True) -> None:
        try:
            try:
                if self.workbook is not None:
                    self.workbook.Close(SaveChanges=save)
            finally:
                if self.app is not None:
                    try:
                        self.app.Quit()
                    finally:
                        self.app = None
        finally:
            self.workbook = None
            if self._com_initialized:
                import pythoncom
                # Balanced against CoInitialize() in open().
                pythoncom.CoUninitialize()
                self._com_initialized = False

    def recalculate(self) -> None:
        if self.app is None:
            raise RuntimeError("Excel is not open.")
        self.app.CalculateFullRebuild()

    def cell_values(self, sheet_name: str, cells: list[str]) -> dict[str, Any]:
        """Read displayed (calculated) values after recalculation."""
        if self.workbook is None:
            raise RuntimeError("No workbook is open.")
        sheet = self.workbook.Worksheets(sheet_name)
        return {cell: sheet.Range(cell).Value for cell in cells}


def verify_formulas_in_excel(path: Path, checks: dict[str, list[str]]) -> dict[str, Any]:
    """Open the workbook in real Excel, force a full rebuild, and read calculated cells.

    openpyxl writes formulas without cached results, so only Excel itself can
    confirm that a formula evaluates to the expected value. The workbook is
    closed without saving.
    """
    com = ExcelCom(visible=False)
    com.open(path)
    try:
        com.recalculate()
        results = {sheet: com.cell_values(sheet, cells) for sheet, cells in checks.items()}
        return {"ok": True, "values": results}
    except Exception as exc:
        return {"ok": False, "error": f"{type(exc).__name__}: {exc}"}
    finally:
        com.close(save=False)
