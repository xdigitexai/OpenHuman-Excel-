from openhuman_excel.workbook import make_working_copy, inspect_with_openpyxl
from openpyxl import Workbook


def test_copy_preserves_source(tmp_path):
    source = tmp_path / "source.xlsx"
    wb = Workbook()
    wb.active["A1"] = "hello"
    wb.save(source)
    copy = make_working_copy(source, tmp_path / "copies")
    assert copy.exists() and copy != source
    assert inspect_with_openpyxl(copy)["sheets"][0]["name"] == "Sheet"
    assert source.exists()
