from __future__ import annotations

import shutil
from datetime import datetime
from pathlib import Path
from typing import Any


SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}


def make_working_copy(source: Path, destination_dir: Path) -> Path:
    source = source.expanduser().resolve()
    if not source.is_file() or source.suffix.lower() not in SUPPORTED_EXTENSIONS:
        raise ValueError("Choose an existing .xlsx, .xlsm, or .xls workbook.")
    destination_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    target = destination_dir / f"{source.stem}.openh-{stamp}{source.suffix}"
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

    def open(self, path: Path) -> dict[str, Any]:
        if __import__("os").name != "nt":
            raise RuntimeError("Microsoft Excel COM automation is available on Windows only.")
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        self.app = win32com.client.DispatchEx("Excel.Application")
        self.app.Visible = self.visible
        self.app.DisplayAlerts = False
        self.workbook = self.app.Workbooks.Open(str(path.resolve()), ReadOnly=False)
        sheets = []
        for sheet in self.workbook.Worksheets:
            used = sheet.UsedRange
            sheets.append({"name": sheet.Name, "rows": used.Rows.Count,
                           "columns": used.Columns.Count})
        return {"sheets": sheets, "active_sheet": self.app.ActiveSheet.Name}

    def close(self, save: bool = True) -> None:
        if self.workbook is not None:
            self.workbook.Close(SaveChanges=save)
            self.workbook = None
        if self.app is not None:
            self.app.Quit()
            self.app = None
