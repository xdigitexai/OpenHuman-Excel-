# OpenHuman Excel

A Windows desktop agent that performs real work in Microsoft Excel through a visible, reviewable workflow.

## Status

Working end-to-end on Windows with real Excel (verified 2026-11-08, see `docs/TEST_REPORT.md`):
inspection, model planning with strict validation, plan execution in a working copy, Excel-side
formula verification, approval gates for destructive actions, pause/resume/stop, and a persistent
action ledger that prevents repeats across restarts while allowing crash recovery.
The project is intentionally Excel-only; CAD belongs in a separate future application.

## Requirements
- Windows 10/11
- Python 3.11+
- Microsoft Excel desktop (recommended for full COM automation)
- An API key for an OpenAI-compatible model provider (optional for local workbook inspection and deterministic actions)

## Install
```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -e ".[dev]"
```

## Run
```powershell
openh-excel
# or
python -m openhuman_excel
```

Choose a workbook, describe a concrete task, and press **Start** to inspect and plan.
Review the plan and press **Apply plan** to execute it against the working copy.
Use **Pause**, **Stop**, and **Resume** to control the session; destructive actions pop an approval dialog.

## Supported actions
`set_cell`, `write_formula`, `rename_sheet`, `add_sheet`, `remove_duplicates` (approval required),
`set_header_style`. Plans produced by the model are validated step-by-step; invalid or unsupported
steps are dropped and reported, never executed. Formulas that reference DDE/external objects or
network functions (`WEBSERVICE`, `FILTERXML`, URLs) are rejected.

## Safety and operating model
- Original workbooks are preserved by default; edits happen in a timestamped working copy.
- Every action is logged to a local SQLite ledger; completed actions never repeat, even after an app restart, while interrupted ones can be retried after a crash.
- Destructive or ambiguous actions require approval.
- The agent stops when the requested work is complete; it never invents tasks to fill the daily window.
- The configured daily session cap (`OPENH_EXCEL_DAILY_LIMIT_HOURS`, max 10 h) bounds agent work time.
- Formula results are verified by recalculating in real Excel when it is available.
- Human-paced mode controls visible interaction timing; it is not intended to deceive monitoring systems.

## Configuration
Copy `.env.example` to `.env` and set an OpenAI-compatible endpoint, model, and API key. Secrets are read from environment variables and are never written to task logs.

## Tests
```powershell
python -m pytest                             # unit + integration (53 tests)
python scripts/live_excel_verification.py    # end-to-end against real Excel
python scripts/gui_smoke_test.py             # drives the real desktop app
```

## Current implementation scope
Workbook discovery, safe working-copy creation, openpyxl inspection, model planning with strict
schema validation, a constrained action executor wired to the GUI, visible Excel COM connection
with balanced lifecycle, Excel-side recalculation verification, task controls (pause/resume/stop),
persistent task/action ledger with crash-safe duplicate guard, daily budget enforcement, and an
OpenAI-compatible provider adapter. See `docs/ROADMAP.md` for what comes next.
