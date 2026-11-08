import json

import pytest
from openpyxl import Workbook, load_workbook

from openhuman_excel.agent import AgentService, SessionBudgetExceeded
from openhuman_excel.ledger import Ledger
from openhuman_excel.provider import ProviderError

PLAN = {
    "summary": "Add a greeting and a formula",
    "steps": [
        {"type": "set_cell", "sheet": "Data", "cell": "A1", "value": "Hello"},
        {"type": "write_formula", "sheet": "Data", "cell": "B1", "value": "=LEN(A1)"},
    ],
    "risks": [],
}


class FakeProvider:
    def __init__(self, response=None, error=None):
        self.response = json.dumps(response) if isinstance(response, dict) else response
        self.error = error
        self.calls = 0

    def complete(self, system, user):
        self.calls += 1
        if self.error:
            raise self.error
        return self.response


def make_service(tmp_path, provider):
    ledger = Ledger(tmp_path / "agent.sqlite3")
    return AgentService(log=lambda m: None, ledger=ledger, provider=provider)


def make_source(tmp_path):
    source = tmp_path / "source.xlsx"
    wb = Workbook()
    wb.active.title = "Data"
    wb.save(source)
    return source


def test_inspect_and_plan_creates_task_and_keeps_source(tmp_path, monkeypatch):
    monkeypatch.setattr("openhuman_excel.config.settings.data_dir",
                        tmp_path)
    source = make_source(tmp_path)
    before = source.read_bytes()
    service = make_service(tmp_path, FakeProvider(PLAN))
    result = service.inspect_and_plan(source, "greet and count")
    assert service.ledger.get_task_state(result["task_id"]) == "planned"
    copy = result["working_copy"]
    assert str(tmp_path) in copy
    assert source.read_bytes() == before  # original untouched
    check = load_workbook(copy)
    assert check.sheetnames == ["Data"]
    check.close()


def test_run_plan_executes_actions_end_to_end(tmp_path, monkeypatch):
    monkeypatch.setattr("openhuman_excel.config.settings.data_dir",
                        tmp_path)
    source = make_source(tmp_path)
    before = source.read_bytes()
    service = make_service(tmp_path, FakeProvider(PLAN))
    result = service.inspect_and_plan(source, "greet and count")
    applied = service.run_plan(result, approve=lambda a: True)
    assert applied["executed"] == 2
    assert service.ledger.get_task_state(result["task_id"]) == "completed"
    check = load_workbook(result["working_copy"])
    assert check["Data"]["A1"].value == "Hello"
    assert check["Data"]["B1"].value == "=LEN(A1)"
    check.close()
    assert source.read_bytes() == before  # original still untouched


def test_run_plan_skips_unapproved_destructive_action(tmp_path, monkeypatch):
    monkeypatch.setattr("openhuman_excel.config.settings.data_dir",
                        tmp_path)
    source = make_source(tmp_path)
    plan = {"summary": "s", "steps": [
        {"type": "remove_duplicates", "sheet": "Data", "range": "A1:A2"}]}
    service = make_service(tmp_path, FakeProvider(plan))
    result = service.inspect_and_plan(source, "dedupe")
    applied = service.run_plan(result, approve=lambda a: False)
    assert applied["executed"] == 0
    assert service.ledger.get_task_state(result["task_id"]) == "completed"


def test_malformed_provider_json_does_not_crash(tmp_path, monkeypatch):
    monkeypatch.setattr("openhuman_excel.config.settings.data_dir",
                        tmp_path)
    source = make_source(tmp_path)
    service = make_service(tmp_path, FakeProvider(response="not json at all"))
    result = service.inspect_and_plan(source, "do something")
    assert result["plan"]["steps"] == []
    assert "raw_response" in result["plan"]
    # Applying a stepless plan edits nothing and completes cleanly.
    applied = service.run_plan(result, approve=lambda a: True)
    assert applied["executed"] == 0


def test_provider_missing_key_raises_clean_error(tmp_path):
    from openhuman_excel.provider import OpenAICompatibleProvider
    provider = OpenAICompatibleProvider("https://example.invalid", "", "m")
    with pytest.raises(ProviderError, match="No API key"):
        provider.complete("sys", "user")


def test_provider_timeout_maps_to_provider_error(monkeypatch):
    import httpx

    from openhuman_excel.provider import OpenAICompatibleProvider

    def boom(*args, **kwargs):
        raise httpx.ConnectTimeout("timed out")

    monkeypatch.setattr(httpx, "post", boom)
    provider = OpenAICompatibleProvider("https://example.invalid", "key", "m")
    with pytest.raises(ProviderError, match="ConnectTimeout"):
        provider.complete("sys", "user")


def test_provider_never_leaks_api_key_in_errors(monkeypatch):
    import httpx

    from openhuman_excel.provider import OpenAICompatibleProvider

    class FakeResponse:
        def raise_for_status(self):
            request = httpx.Request("POST", "https://example.invalid/v1/chat/completions")
            raise httpx.HTTPStatusError("boom", request=request,
                                        response=httpx.Response(500, request=request))

    monkeypatch.setattr(httpx, "post", lambda *a, **k: FakeResponse())
    secret = "sk-super-secret-key-123"
    provider = OpenAICompatibleProvider("https://example.invalid", secret, "m")
    with pytest.raises(ProviderError) as excinfo:
        provider.complete("sys", "user")
    assert secret not in str(excinfo.value)


def test_provider_failure_marks_error_and_keeps_workbook(tmp_path, monkeypatch):
    monkeypatch.setattr("openhuman_excel.config.settings.data_dir",
                        tmp_path)
    source = make_source(tmp_path)
    before = source.read_bytes()
    service = make_service(tmp_path, FakeProvider(error=ProviderError("down")))
    with pytest.raises(ProviderError):
        service.inspect_and_plan(source, "anything")
    assert source.read_bytes() == before


def test_duplicate_plan_never_repeats_completed_actions(tmp_path, monkeypatch):
    """Re-applying the same plan (e.g. after a restart) must not redo work."""
    monkeypatch.setattr("openhuman_excel.config.settings.data_dir",
                        tmp_path)
    source = make_source(tmp_path)
    service = make_service(tmp_path, FakeProvider(PLAN))
    result = service.inspect_and_plan(source, "greet")
    assert service.run_plan(result, approve=lambda a: True)["executed"] == 2
    # Same task, same actions: all completed, so nothing runs again.
    second = service.run_plan(result, approve=lambda a: True)
    assert second["executed"] == 0


def test_daily_budget_stops_instead_of_inventing_work(tmp_path):
    service = make_service(tmp_path, FakeProvider(PLAN))
    # Pretend the session started 99 hours ago against a 10-hour cap.
    import time
    service._session_started_monotonic = time.monotonic() - 99 * 3600
    with pytest.raises(SessionBudgetExceeded):
        service._check_budget()


def test_stop_during_planning_interrupts(tmp_path, monkeypatch):
    monkeypatch.setattr("openhuman_excel.config.settings.data_dir",
                        tmp_path)
    source = make_source(tmp_path)
    service = make_service(tmp_path, FakeProvider(PLAN))
    service.stop()
    with pytest.raises(InterruptedError):
        service.inspect_and_plan(source, "anything")
