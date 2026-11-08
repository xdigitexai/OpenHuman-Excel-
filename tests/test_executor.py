from openpyxl import Workbook, load_workbook
from openhuman_excel.executor import WorkbookExecutor
from openhuman_excel.ledger import Ledger


def test_executor_writes_cell_and_blocks_repeat(tmp_path):
    book = tmp_path / "book.xlsx"
    wb = Workbook()
    wb.active.title = "Data"
    wb.save(book)
    ledger = Ledger(tmp_path / "ledger.sqlite3")
    ledger.create_task("task", "write", str(book))
    executor = WorkbookExecutor(ledger, pace_seconds=0)
    action = {"type": "set_cell", "sheet": "Data", "cell": "A1", "value": "Hello"}
    assert executor.execute(book, "task", [action], lambda _: True) == 1
    assert executor.execute(book, "task", [action], lambda _: True) == 0
    check = load_workbook(book, data_only=False)
    assert check["Data"]["A1"].value == "Hello"
    check.close()


def test_executor_requires_approval(tmp_path):
    book = tmp_path / "book.xlsx"
    wb = Workbook()
    wb.active.title = "Data"
    wb.save(book)
    ledger = Ledger(tmp_path / "ledger.sqlite3")
    ledger.create_task("task", "rename", str(book))
    executor = WorkbookExecutor(ledger, pace_seconds=0)
    action = {"type": "rename_sheet", "sheet": "Data", "new_name": "Archive",
              "requires_approval": True}
    assert executor.execute(book, "task", [action], lambda _: False) == 0
    check = load_workbook(book)
    assert "Data" in check.sheetnames
    check.close()
