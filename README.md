# OpenHuman Excel

A Windows desktop agent that performs real work in Microsoft Excel through a visible, reviewable workflow.

## Status
Initial working foundation. The project is intentionally Excel-only; CAD belongs in a separate future application.

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

Choose a workbook, describe a concrete task, and press **Start**. The agent works on a copy by default. Use **Pause**, **Stop**, and **Resume** to control the session.

## Safety and operating model
- Original workbooks are preserved by default; edits happen in a timestamped working copy.
- Every action is logged to a local SQLite ledger.
- Destructive or ambiguous actions require approval.
- The agent stops when the requested work is complete; it does not invent tasks to fill a 10-hour window.
- A configured daily session may run for up to 10 hours, but the laptop must remain awake and Excel available.
- Human-paced mode controls visible interaction timing; it is not intended to deceive monitoring systems.

## Configuration
Copy `.env.example` to `.env` and set an OpenAI-compatible endpoint, model, and API key. Secrets are read from environment variables and are never written to task logs.

## Current implementation scope
Workbook discovery, safe working-copy creation, visible Excel COM connection, workbook/sheet inspection, task controls, persistent task/action ledger, duplicate-action guard, checkpointing, and a provider interface. Advanced workbook skills are added incrementally and validated with tests.
