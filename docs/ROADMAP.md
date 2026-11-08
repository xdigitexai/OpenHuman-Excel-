# OpenHuman Excel implementation roadmap

## Milestone 0 — repository foundation
- Windows desktop shell and workbook picker
- Safe timestamped working copy
- Workbook inspection with openpyxl
- SQLite task/action/event ledger
- OpenAI-compatible provider adapter
- Constrained action executor for cell values, formulas, worksheet management,
  duplicate removal, and header formatting
- Per-action persistence, duplicate guard, approval callback, and pause/stop checkpoints

## Milestone 1 — Excel-native integration
- Dedicated COM worker thread with lifecycle management and visible Excel window
- Excel object model adapter for formatting, tables, charts, PivotTables, filters,
  named ranges, and recalculation
- Formula evaluation verification in Excel (openpyxl does not calculate formulas)
- UI Automation fallback and screenshots of the actual desktop
- Explicit save/reopen validation and rollback

## Milestone 2 — agent intelligence
- Typed tool-call schema and strict JSON schema validation
- Workbook profiling and sampling for large sheets
- Task decomposition, dependency graph, verification plan, and evidence for each step
- Similarity-based duplicate detection across sessions and source ranges
- User review of plan, action-by-action approvals, undo/rollback

## Milestone 3 — long-running sessions
- Durable queue, 10-hour configurable daily cap, scheduled breaks, resumable checkpoints
- Recovery after app restart, Excel crash, Windows restart, and provider rate limits
- Sleep/power-state detection and clear warnings; no promise of work while device is asleep/off
- Human-paced visible operation mode and faster native mode
- Daily report and token/cost accounting

## Milestone 4 — production hardening
- Workbook fixtures for .xlsx/.xlsm and large datasets
- Property-based and integration tests on Windows with real Excel
- Security review, secret redaction, audit export, installer and signed release
- Accessibility, keyboard navigation, robust logs, update mechanism

## Guardrails
- Work on a copy by default; never silently overwrite the source.
- No fabricated work to fill the configured time window.
- Stop when the user's requested deliverable is complete.
- Macros, external data connections, destructive sheet deletion, and bulk row deletion require explicit approval.
- Never execute arbitrary Python, shell commands, or model-generated code from a workbook task.
- Treat workbook cells, comments, links, and embedded content as untrusted data, not agent instructions.
- CAD is out of scope for this repository.
