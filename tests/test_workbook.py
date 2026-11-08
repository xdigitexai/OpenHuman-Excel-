import pytest
from openpyxl import Workbook

from openhuman_excel.workbook import (
    ExcelCom,
    inspect_with_openpyxl,
    make_working_copy,
)

EXCEL_AVAILABLE = False
if __import__("os").name == "nt":
    try:
        import win32com.client  # noqa: F401
        EXCEL_AVAILABLE = True
    except ImportError:
        EXCEL_AVAILABLE = False


def make_xlsx(path, sheet="Data"):
    wb = Workbook()
    wb.active.title = sheet
    wb.save(path)


def test_copy_preserves_source(tmp_path):
    source = tmp_path / "source.xlsx"
    wb = Workbook()
    wb.active["A1"] = "hello"
    wb.save(source)
    copy = make_working_copy(source, tmp_path / "copies")
    assert copy.exists() and copy != source
    assert inspect_with_openpyxl(copy)["sheets"][0]["name"] == "Sheet"
    assert source.exists()


def test_copies_are_unique_within_the_same_second(tmp_path):
    source = tmp_path / "source.xlsx"
    make_xlsx(source)
    copies = {make_working_copy(source, tmp_path / "copies") for _ in range(5)}
    assert len(copies) == 5


def test_unsupported_extension_rejected(tmp_path):
    source = tmp_path / "old.xls"
    source.write_bytes(b"legacy")
    with pytest.raises(ValueError):
        make_working_copy(source, tmp_path / "copies")


def test_inspect_reports_sheet_dimensions(tmp_path):
    source = tmp_path / "source.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Grid"
    for r in range(1, 4):
        for c in range(1, 3):
            ws.cell(r, c, r * c)
    wb.save(source)
    info = inspect_with_openpyxl(source)
    grid = next(s for s in info["sheets"] if s["name"] == "Grid")
    assert grid["max_row"] == 3 and grid["max_column"] == 2


@pytest.mark.skipif(not EXCEL_AVAILABLE, reason="Windows + pywin32 required")
def test_excel_com_opens_and_closes_cleanly(tmp_path):
    source = tmp_path / "source.xlsx"
    make_xlsx(source)
    copy = make_working_copy(source, tmp_path / "copies")
    com = ExcelCom(visible=False)
    info = com.open(copy)
    try:
        assert info["sheets"][0]["name"] == "Data"
        com.workbook.Worksheets("Data").Range("A1").Value = "from-com"
    finally:
        com.close(save=True)
    wb = __import__("openpyxl").load_workbook(copy)
    assert wb["Data"]["A1"].value == "from-com"
    wb.close()
