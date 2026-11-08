import io
import threading
import time

import pytest
from openpyxl import Workbook, load_workbook

from openhuman_excel.executor import ExecutionError, WorkbookExecutor
from openhuman_excel.ledger import Ledger


def make_book(tmp_path, name="book.xlsx", rows=None):
    book = tmp_path / name
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    for row in (rows or []):
        ws.append(row)
    wb.save(book)
    return book


def make_executor(tmp_path, task="task"):
    ledger = Ledger(tmp_path / "ledger.sqlite3")
    ledger.create_task(task, "test", "book.xlsx")
    logs: list[str] = []
    return WorkbookExecutor(ledger, log=logs.append, pace_seconds=0), logs


def test_executor_writes_cell_and_blocks_repeat(tmp_path):
    book = make_book(tmp_path)
    executor, _ = make_executor(tmp_path)
    action = {"type": "set_cell", "sheet": "Data", "cell": "A1", "value": "Hello"}
    assert executor.execute(book, "task", [action], lambda _: True) == 1
    assert executor.execute(book, "task", [action], lambda _: True) == 0
    check = load_workbook(book, data_only=False)
    assert check["Data"]["A1"].value == "Hello"
    check.close()


def test_executor_requires_approval(tmp_path):
    book = make_book(tmp_path)
    executor, _ = make_executor(tmp_path)
    action = {"type": "rename_sheet", "sheet": "Data", "new_name": "Archive",
              "requires_approval": True}
    assert executor.execute(book, "task", [action], lambda _: False) == 0
    check = load_workbook(book)
    assert "Data" in check.sheetnames
    check.close()


def test_failed_action_can_be_retried(tmp_path):
    """A failed action must not poison its idempotency key forever."""
    book = make_book(tmp_path)
    executor, _ = make_executor(tmp_path)
    bad = {"type": "set_cell", "sheet": "Missing", "cell": "A1", "value": "x"}
    with pytest.raises(ExecutionError):
        executor.execute(book, "task", [bad], lambda _: True)
    good = {"type": "set_cell", "sheet": "Data", "cell": "A1", "value": "ok"}
    assert executor.execute(book, "task", [good], lambda _: True) == 1


def test_action_interrupted_by_simulated_crash_is_retried(tmp_path):
    book = make_book(tmp_path)
    ledger = Ledger(tmp_path / "ledger.sqlite3")
    ledger.create_task("task", "test", str(book))
    action = {"type": "set_cell", "sheet": "Data", "cell": "A1", "value": "v"}
    ledger.begin_action("task", "crash-key", "set_cell", "Data!A1", action)
    executor = WorkbookExecutor(ledger, log=lambda m: None, pace_seconds=0)
    assert executor.execute(book, "task",
                            [dict(action, idempotency_key="crash-key")],
                            lambda _: True) == 1
    check = load_workbook(book)
    assert check["Data"]["A1"].value == "v"
    check.close()


def test_completed_action_not_repeated_after_restart(tmp_path):
    book = make_book(tmp_path)
    ledger = Ledger(tmp_path / "ledger.sqlite3")
    ledger.create_task("task", "test", str(book))
    action = {"type": "set_cell", "sheet": "Data", "cell": "A1", "value": "once"}
    executor = WorkbookExecutor(ledger, log=lambda m: None, pace_seconds=0)
    assert executor.execute(book, "task", [action], lambda _: True) == 1
    # A brand-new executor over the same ledger simulates an app restart.
    restarted = WorkbookExecutor(Ledger(tmp_path / "ledger.sqlite3"),
                                 log=lambda m: None, pace_seconds=0)
    assert restarted.execute(book, "task", [action], lambda _: True) == 0
    check = load_workbook(book)
    assert check["Data"]["A1"].value == "once"
    check.close()


def test_remove_duplicates_removes_repeats_and_keeps_first(tmp_path):
    rows = [["id", "city"], [1, "Oslo"], [2, "Rome"], [1, "Oslo"], [3, "Oslo"]]
    book = make_book(tmp_path, rows=rows)
    executor, _ = make_executor(tmp_path)
    action = {"type": "remove_duplicates", "sheet": "Data", "range": "A1:B5"}
    assert executor.execute(book, "task", [action], lambda _: True) == 1
    check = load_workbook(book)["Data"]
    values = [[check.cell(r, 1).value, check.cell(r, 2).value] for r in range(1, 5)]
    assert values == [["id", "city"], [1, "Oslo"], [2, "Rome"], [3, "Oslo"]]


def test_destructive_action_forces_approval_even_if_plan_forgets(tmp_path):
    book = make_book(tmp_path, rows=[["a"], [1], [1]])
    executor, _logs = make_executor(tmp_path)
    action = {"type": "remove_duplicates", "sheet": "Data", "range": "A1:A3"}
    # Plan forgot requires_approval; executor must demand it anyway.
    assert executor.execute(book, "task", [action], lambda _: False) == 0
    check = load_workbook(book)["Data"]
    assert check.cell(3, 1).value == 1  # duplicate still present


def test_unsupported_action_type_aborts(tmp_path):
    book = make_book(tmp_path)
    executor, _ = make_executor(tmp_path)
    with pytest.raises(ExecutionError):
        executor.execute(book, "task",
                         [{"type": "delete_all", "sheet": "Data"}], lambda _: True)


def test_formula_must_start_with_equals(tmp_path):
    book = make_book(tmp_path)
    executor, _ = make_executor(tmp_path)
    action = {"type": "write_formula", "sheet": "Data", "cell": "B1", "value": "1+1"}
    with pytest.raises(ExecutionError):
        executor.execute(book, "task", [action], lambda _: True)


def test_dde_formula_rejected_in_executor(tmp_path):
    book = make_book(tmp_path)
    executor, _ = make_executor(tmp_path)
    action = {"type": "write_formula", "sheet": "Data", "cell": "B1",
              "value": "=cmd|'/c calc'!A1"}
    with pytest.raises(ExecutionError):
        executor.execute(book, "task", [action], lambda _: True)


def test_log_output_is_ascii_safe(tmp_path):
    """Logs must survive cp1252 Windows consoles (no arrows, dashes etc.)."""
    book = make_book(tmp_path, rows=[["a"], [1], [1]])
    stream = io.StringIO()
    # Encode every log line as cp1252 exactly like a redirected console would.
    ledger = Ledger(tmp_path / "ledger.sqlite3")
    ledger.create_task("task", "test", str(book))
    executor = WorkbookExecutor(ledger, log=lambda m: stream.write(m + "\n"),
                                pace_seconds=0)
    action = {"type": "remove_duplicates", "sheet": "Data", "range": "A1:A3",
              "requires_approval": True}
    executor.execute(book, "task", [action], lambda _: False)
    executor.execute(book, "task", [action], lambda _: False)  # duplicate-skip path
    encoded = stream.getvalue().encode("cp1252")  # must not raise
    assert b"remove_duplicates" in encoded


def test_pause_blocks_until_resume(tmp_path):
    book = make_book(tmp_path)
    executor, _ = make_executor(tmp_path)
    executor.pause()
    started = threading.Event()

    def run():
        started.set()
        executor.execute(book, "task",
                         [{"type": "set_cell", "sheet": "Data", "cell": "A1",
                           "value": "paused-test"}], lambda _: True)

    thread = threading.Thread(target=run, daemon=True)
    thread.start()
    started.wait(2)
    time.sleep(0.5)
    check = load_workbook(book)
    untouched = check["Data"]["A1"].value is None
    check.close()
    executor.resume()
    thread.join(5)
    assert not thread.is_alive()
    assert untouched, "execute() must not apply actions while paused"
    check = load_workbook(book)
    assert check["Data"]["A1"].value == "paused-test"
    check.close()


def test_stop_interrupts_execution(tmp_path):
    book = make_book(tmp_path)
    executor, _ = make_executor(tmp_path)
    executor.stop()
    with pytest.raises(InterruptedError):
        executor.execute(book, "task",
                         [{"type": "set_cell", "sheet": "Data", "cell": "A1",
                           "value": "x"}], lambda _: True)
    check = load_workbook(book)
    assert check["Data"]["A1"].value is None
    check.close()


def test_xlsm_roundtrip_keeps_extension(tmp_path):
    book = make_book(tmp_path, name="book.xlsm")
    executor, _ = make_executor(tmp_path)
    action = {"type": "set_cell", "sheet": "Data", "cell": "A1", "value": "macro-book"}
    assert executor.execute(book, "task", [action], lambda _: True) == 1
    check = load_workbook(book, keep_vba=True)
    assert check["Data"]["A1"].value == "macro-book"
    check.close()
