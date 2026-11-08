from openhuman_excel.ledger import Ledger


def test_duplicate_action_is_rejected(tmp_path):
    ledger = Ledger(tmp_path / "agent.sqlite3")
    ledger.create_task("t1", "clean rows", "book.xlsx")
    assert ledger.begin_action("t1", "sheet1:remove-duplicates:A1:D100",
                               "remove_duplicates", "Sheet1!A1:D100")
    assert not ledger.begin_action("t1", "sheet1:remove-duplicates:A1:D100",
                                   "remove_duplicates", "Sheet1!A1:D100")
