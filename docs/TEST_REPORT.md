# OpenHuman Excel — Testing & Verification Report

Date: 2026-11-08 · Branch: `feature/desktop-agent` · Platform: Windows 10 (10.0.19045), Python 3.14.8, Microsoft Excel 2016 (COM verified)

## Summary

| Suite | Result |
|---|---|
| Unit/integration tests (`pytest`) | **53 passed, 0 failed, 0 skipped** (baseline before audit: 4) |
| Live Excel verification (`scripts/live_excel_verification.py`) | **16 passed, 0 failed** — real Excel via COM |
| GUI smoke test (`scripts/gui_smoke_test.py`) | **9 passed, 0 failed** — real Tkinter app driven end-to-end |
| Lint (`ruff`) | clean |

## What was verified working (with evidence)

- **Working-copy safety.** Every run edits a timestamped copy; SHA-256 of the source workbook was identical before/after planning and execution (live log). Copies are unique even within the same second (retry-safe).
- **Plan → validated actions → execution.** The full chain now exists end-to-end: model plan → strict schema validation (`planning.py`) → `WorkbookExecutor` → ledger. Verified in pytest and live.
- **Cell edits, formulas, dedup, header styling.** Executed live: `set_cell`, `write_formula`, `remove_duplicates` (7 rows → 4), `set_header_style` (bold/white on dark blue).
- **Real-Excel formula verification.** openpyxl-written `=SUM(B2:B4)` was recalculated by Excel itself via `CalculateFullRebuild` and read back as 650 (live, visible window screenshot `docs/screenshots/01_excel_visible_working_copy.png`).
- **Original preservation.** Source file byte-identical (SHA-256) after the entire pipeline.
- **Approval requirements.** `remove_duplicates` always demands approval (defense in depth — forced even when the plan forgets to flag it). GUI shows an approval dialog; tests cover deny and approve paths.
- **Duplicate-action prevention across restarts.** Completed actions are rejected on re-run, including with a freshly opened ledger (simulated app restart). Interrupted/failed actions ARE retryable (crash recovery) — the old code blocked them forever (fixed).
- **Pause / Resume / Stop.** Executor pauses between actions without applying partial work (unit test with a real thread), Stop raises `InterruptedError` and marks the task `paused`; the GUI buttons work and are enabled/disabled correctly during runs.
- **Error recovery.** Provider failures (missing key, timeouts, HTTP errors, malformed JSON) produce clean errors; the workbook and source stay untouched; the task state reflects the failure.
- **GUI ↔ executor integration.** GUI smoke test drove the real app: Start → plan → Apply plan → executed → verified → ledger `completed`. Screenshots 02–04.
- **.xlsm handling.** `keep_vba` is used on load and save; a `vbaProject.bin` carried through an openpyxl edit byte-identical. (See "not yet verified" below.)
- **Daily budget.** Session work time is checked against `OPENH_EXCEL_DAILY_LIMIT_HOURS`; exceeding it raises `SessionBudgetExceeded` instead of inventing work (tested).
- **Security.** API key never appears in logs or error strings (tested); no `eval`/`exec`/`subprocess` in the package; model formulas with DDE payloads (`=cmd|'/c calc'!A1`), `WEBSERVICE`/`FILTERXML`/network references are rejected before reaching Excel (tested); workbook cell contents are never sent to the model (only sheet names/dimensions).

## Defects found and fixed (commit `c665c38`, follow-ups in final commit)

| # | Defect | Fix |
|---|---|---|
| D1 | GUI never invoked `WorkbookExecutor` — plans could not become actions at all | `AgentService.run_plan()` + GUI "Apply plan" |
| D2 | Pause/Stop/Resume buttons were no-ops (flags never checked) | Checkpoints in planning; executor wired to buttons |
| D3 | Actions blocked from retry after a crash (non-completed ledger rows rejected forever) | `begin_action` retries unless the action completed |
| D4 | `UnicodeEncodeError` on cp1252 consoles (log used `→`), aborting runs | ASCII-safe logging |
| D5 | No plan schema validation; untrusted model output flowed unchecked | `planning.py` strict per-type validation + sanitization |
| D6 | Destructive actions did not require approval | Forced for `remove_duplicates` (plan + executor) |
| D7 | `ExcelCom`: no `CoUninitialize`, unsafe close, never used anywhere | Balanced COM lifecycle, failure-safe close, used for verification |
| D8 | Task lifecycle never left `planned` | `queued → planned → running → completed/failed/paused` |
| D9 | Daily limit configured but never enforced | `SessionBudgetExceeded` enforcement |
| D10 | Formula injection (DDE/WEBSERVICE) unguarded | Formula allow-guards in validator and executor |
| D11 | `inspect_and_plan` reset `_stop` at entry, wiping a user's Stop | Reset moved to explicit `reset_controls()` on new run |
| D12 | Ledger kept sqlite connections open (relied on GC) | Explicit close on every operation |

## Not yet verified / outstanding (honest gaps)

- **Macro execution after edit.** openpyxl-level `vbaProject.bin` retention is verified, and Excel opens the result, but a live "macro still runs in Excel after the agent edits the file" check requires a real macro workbook or Excel's "Trust access to the VBA project object model" (denied on this machine). Fixture creation programmatically was blocked by that trust setting.
- **Ten-hour endurance.** NOT tested. Short automated tests cover persistence, duplicate prevention, and budget checks; no simulated-or-real 10-hour run has been performed. Sleep/power-state handling and provider mid-run interruption recovery are implemented at the code level (pause/stop/retry) but not endurance-tested.
- **Undo/rollback**, PivotTables/charts/filters/named-ranges actions, UI-Automation fallback, and action-level approval streaming are roadmap milestones 1–3 and are not implemented yet.
- **Plan step that requires Excel to evaluate before planning** (e.g., verifying existing formula results pre-execution) works only via the post-execution verification pass.
- **Model-generated plans in production** were exercised with a stub provider (deterministic). A live LLM plan was not executed because no API key is configured in this environment; the provider adapter is unit-tested against failures.

## Reproducing

```powershell
py -3.11 -m venv .venv; .\.venv\Scripts\Activate.ps1
pip install -e ".[dev]"
python -m pytest                       # 53 tests
python scripts/live_excel_verification.py   # needs Excel installed
python scripts/gui_smoke_test.py            # launches the real app
```

Artifacts: `docs/screenshots/*.png` (captured during runs), `docs/live_verification_log.txt`, `test_results/` (gitignored).
