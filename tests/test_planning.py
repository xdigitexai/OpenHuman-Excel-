import pytest

from openhuman_excel.planning import (
    PlanValidationError,
    sanitize_sheet_name,
    validate_formula,
    validate_plan,
    validate_step,
)


def test_valid_set_cell_step_is_sanitized():
    action = validate_step({"type": "set_cell", "sheet": " Data ", "cell": "$A$1",
                            "value": "hello"})
    assert action == {"type": "set_cell", "sheet": "Data", "cell": "A1", "value": "hello"}


def test_valid_formula_step_passes():
    action = validate_step({"type": "write_formula", "sheet": "Data", "cell": "B2",
                            "value": "=SUM(A1:A10)"})
    assert action["value"] == "=SUM(A1:A10)"


def test_unsupported_type_is_rejected():
    with pytest.raises(PlanValidationError):
        validate_step({"type": "run_macro", "sheet": "Data"})


def test_non_dict_step_is_rejected():
    with pytest.raises(PlanValidationError):
        validate_step("set_cell")


def test_bad_cell_reference_is_rejected():
    with pytest.raises(PlanValidationError):
        validate_step({"type": "set_cell", "sheet": "Data", "cell": "not-a-cell",
                       "value": 1})


def test_formula_must_start_with_equals():
    with pytest.raises(PlanValidationError):
        validate_step({"type": "write_formula", "sheet": "Data", "cell": "B1",
                       "value": "SUM(A1:A10)"})


def test_dde_formula_is_rejected():
    # =cmd|'/c calc'!A1 style DDE payloads must never reach Excel.
    with pytest.raises(PlanValidationError):
        validate_formula("=cmd|'/c calc'!A1")


def test_network_formula_is_rejected():
    with pytest.raises(PlanValidationError):
        validate_formula('=WEBSERVICE("http://evil.example/x")')


def test_cross_sheet_formula_is_allowed():
    assert validate_formula("=Sheet2!A1+1") == "=Sheet2!A1+1"


def test_rename_sanitizes_reserved_characters():
    action = validate_step({"type": "rename_sheet", "sheet": "Data",
                            "new_name": "bad:name[here]"})
    assert action["new_name"] == "badnamehere"


def test_rename_to_same_name_is_rejected():
    with pytest.raises(PlanValidationError):
        validate_step({"type": "rename_sheet", "sheet": "Data", "new_name": "data"})


def test_remove_duplicates_requires_approval():
    action = validate_step({"type": "remove_duplicates", "sheet": "Data",
                            "range": "A1:D100"})
    assert action["requires_approval"] is True


def test_bad_range_is_rejected():
    with pytest.raises(PlanValidationError):
        validate_step({"type": "remove_duplicates", "sheet": "Data", "range": "A1"})


def test_header_row_must_be_positive_int():
    with pytest.raises(PlanValidationError):
        validate_step({"type": "set_header_style", "sheet": "Data", "row": 0})
    with pytest.raises(PlanValidationError):
        validate_step({"type": "set_header_style", "sheet": "Data", "row": "1"})


def test_sheet_name_length_is_capped():
    assert len(sanitize_sheet_name("x" * 100)) == 31


def test_validate_plan_mixed_steps():
    plan = {"summary": "s", "steps": [
        {"type": "set_cell", "sheet": "Data", "cell": "A1", "value": 1},
        {"type": "launch_missiles"},
        {"type": "set_cell", "sheet": "Data", "cell": "Zzz", "value": 2},
    ]}
    actions, issues = validate_plan(plan)
    assert len(actions) == 1
    assert len(issues) == 2
    assert all(issue.startswith("Step") for issue in issues)


def test_validate_plan_requires_object():
    with pytest.raises(PlanValidationError):
        validate_plan([{"type": "set_cell"}])


def test_validate_plan_without_steps_is_empty():
    actions, issues = validate_plan({"summary": "just a question"})
    assert actions == [] and issues == []
