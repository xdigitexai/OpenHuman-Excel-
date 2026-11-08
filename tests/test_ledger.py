from openhuman_excel.ledger import Ledger


def test_completed_action_is_rejected(tmp_path):
    ledger = Ledger(tmp_path / "agent.sqlite3")
    ledger.create_task("t1", "clean rows", "book.xlsx")
    assert ledger.begin_action("t1", "sheet1:remove-duplicates:A1:D100",
                               "remove_duplicates", "Sheet1!A1:D100")
    ledger.finish_action("t1", "sheet1:remove-duplicates:A1:D100", "completed")
    # A completed action must never run twice, not even after an app restart.
    assert not ledger.begin_action("t1", "sheet1:remove-duplicates:A1:D100",
                                   "remove_duplicates", "Sheet1!A1:D100")


def test_started_action_is_retryable_after_crash(tmp_path):
    ledger = Ledger(tmp_path / "agent.sqlite3")
    ledger.create_task("t1", "clean rows", "book.xlsx")
    assert ledger.begin_action("t1", "key", "set_cell", "Sheet1!A1")
    # Simulates a crash between begin_action and finish_action: the workbook
    # state is unknown, so the action must be re-runnable.
    assert ledger.begin_action("t1", "key", "set_cell", "Sheet1!A1")


def test_failed_action_is_retryable(tmp_path):
    ledger = Ledger(tmp_path / "agent.sqlite3")
    ledger.create_task("t1", "clean rows", "book.xlsx")
    assert ledger.begin_action("t1", "key", "write_formula", "Sheet1!B1")
    ledger.finish_action("t1", "key", "failed")
    assert ledger.begin_action("t1", "key", "write_formula", "Sheet1!B1")


def test_duplicate_guard_survives_restart(tmp_path):
    path = tmp_path / "agent.sqlite3"
    ledger = Ledger(path)
    ledger.create_task("t1", "task", "book.xlsx")
    assert ledger.begin_action("t1", "key-1", "set_cell", "Sheet1!A1")
    ledger.finish_action("t1", "key-1", "completed")
    # Fresh instance over the same database = app restart.
    reopened = Ledger(path)
    assert not reopened.begin_action("t1", "key-1", "set_cell", "Sheet1!A1")


def test_task_state_roundtrip(tmp_path):
    ledger = Ledger(tmp_path / "agent.sqlite3")
    ledger.create_task("t1", "task", "book.xlsx")
    assert ledger.get_task_state("t1") == "queued"
    for state in ("planned", "running", "completed"):
        ledger.set_task_state("t1", state)
        assert ledger.get_task_state("t1") == state
