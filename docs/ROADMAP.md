# OpenHuman Excel implementation roadmap

Status is marked per item: [done] verified with tests, [partial] implemented but not fully verified,
[open] not started. See `docs/TEST_REPORT.md` for evidence and outstanding gaps.

## Milestone 0 — repository foundation [done]
- [done] Windows desktop shell and workbook picker
- [done] Safe timestamped working copy (unique per run; source byte-preserved, hash-verified)
- [done] Workbook inspection with openpyxl
- [done] SQLite task/action/event ledger (connections closed; completed actions never repeat)
- [done] OpenAI-compatible provider adapter (clean failures: missing key, timeout, HTTP, malformed JSON)
- [done] Constrained action executor for cell values, formulas, worksheet management,
  duplicate removal, and header formatting
- [done] Per-action persistence, duplicate guard (crash-retry safe), approval callback
  (forced for destructive actions), pause/stop checkpoints

## Milestone 1 — Excel-native integration [partial]
- [partial] Dedicated COM worker thread with lifecycle management and visible Excel window
  (COM lifecycle balanced and verified in-process; the executor itself is openpyxl-based,
  Excel used for verification — in-Excel editing is next)
- [partial] Excel object model adapter: recalculation verified live; formatting/tables/charts/
  PivotTables/filters/named ranges [open]
- [done] Formula evaluation verification in Excel (CalculateFullRebuild + value readback, live-tested)
- [open] UI Automation fallback and screenshots of the actual desktop
- [open] Explicit save/reopen validation and rollback

## Milestone 2 — agent intelligence [partial]
- [done] Strict per-type JSON schema validation of plan steps; invalid steps dropped and reported
- [open] Workbook profiling and sampling for large sheets (currently structure-only snapshot)
- [partial] Task decomposition exists via plan steps; dependency graph and per-step evidence [open]
- [partial] Duplicate detection is exact-key based across sessions; similarity-based [open]
- [partial] Approval dialog for destructive actions done; action-by-action review UI and undo/rollback [open]

## Milestone 3 — long-running sessions [partial]
- [partial] Daily cap enforced (raises SessionBudgetExceeded); durable queue and scheduled breaks [open]
- [partial] Crash recovery verified (interrupted actions retried); Excel-crash and Windows-restart
  recovery [open]; provider rate-limit backoff [open]
- [open] Sleep/power-state detection and clear warnings
- [open] Human-paced visible operation mode and faster native mode
- [open] Daily report and token/cost accounting

## Milestone 4 — production hardening [partial]
- [partial] Workbook fixtures for .xlsx/.xlsm; macro-survival verified at package level only
  (needs a real macro workbook or Excel VBA trust to verify execution)
- [partial] Integration tests on Windows with real Excel done; property-based tests [open]
- [partial] Security review done for: secret redaction, formula injection (DDE/network), code
  execution (none permitted), untrusted workbook data (never sent to model); audit export [open]
- [open] Installer and signed release, accessibility, keyboard navigation, update mechanism

## Guardrails
- Work on a copy by default; never silently overwrite the source.
- No fabricated work to fill the configured time window.
- Stop when the user's requested deliverable is complete.
- Macros, external data connections, destructive sheet deletion, and bulk row deletion require explicit approval.
- Never execute arbitrary Python, shell commands, or model-generated code from a workbook task.
- Treat workbook cells, comments, links, and embedded content as untrusted data, not agent instructions.
- CAD is out of scope for this repository.
