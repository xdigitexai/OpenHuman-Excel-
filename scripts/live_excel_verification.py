"""Live Windows/Excel verification for OpenHuman Excel.

Runs the real agent -> plan validation -> executor pipeline against real
Microsoft Excel via COM, in its own Excel instance (DispatchEx), and captures
screenshots. Never touches any workbook the user already has open.

Run with the project venv:  ./.venv/Scripts/python.exe scripts/live_excel_verification.py
"""
from __future__ import annotations

import hashlib
import json
import shutil
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from openpyxl import Workbook, load_workbook

from openhuman_excel.agent import AgentService
from openhuman_excel.ledger import Ledger
from openhuman_excel.workbook import ExcelCom, verify_formulas_in_excel

RESULTS = ROOT / "test_results"
RESULTS.mkdir(exist_ok=True)
LOG_LINES: list[str] = []


def log(message: str) -> None:
    line = str(message)
    LOG_LINES.append(line)
    try:
        print(line)
    except UnicodeEncodeError:
        print(line.encode("ascii", "replace").decode())


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def focus_window(title_substring: str) -> bool:
    """Bring the first top-level window whose title matches to the foreground."""
    try:
        import win32con
        import win32gui
    except ImportError:
        return False
    found: list[tuple[int, str]] = []

    def callback(hwnd: int, _):
        title = win32gui.GetWindowText(hwnd)
        if title and win32gui.IsWindowVisible(hwnd):
            found.append((hwnd, title))

    win32gui.EnumWindows(callback, None)
    for hwnd, title in found:
        if title_substring.lower() in title.lower():
            try:
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                try:
                    win32gui.SetForegroundWindow(hwnd)
                except Exception:
                    # Windows blocks SetForegroundWindow for background
                    # processes; forcing topmost is the reliable fallback.
                    win32gui.SetWindowPos(hwnd, -1, 0, 0, 0, 0,
                                          0x0001 | 0x0002)  # TOPMOST, NOMOVE|NOSIZE
                time.sleep(1.0)
                return True
            except Exception as exc:
                log(f"Focus failed for '{title}': {type(exc).__name__}: {exc}")
    log(f"No window titled like '{title_substring}' found; "
        f"visible windows: {[t for _, t in found]}")
    return False


def screenshot(name: str, window_title: str | None = None) -> Path | None:
    target = RESULTS / name
    if window_title:
        focus_window(window_title)
    ps = (
        "Add-Type -AssemblyName System.Windows.Forms,System.Drawing;"
        "$b = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds;"
        "$bmp = New-Object System.Drawing.Bitmap $b.Width, $b.Height;"
        "$g = [System.Drawing.Graphics]::FromImage($bmp);"
        "$g.CopyFromScreen($b.Location, [System.Drawing.Point]::Empty, $b.Size);"
        f"$bmp.Save('{target.as_posix()}');"
        "$g.Dispose(); $bmp.Dispose()"
    )
    try:
        subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                       check=True, timeout=30)
        log(f"Screenshot saved: {target}")
        return target
    except Exception as exc:
        log(f"Screenshot failed: {type(exc).__name__}: {exc}")
        return None


PASS, FAIL = [], []


def check(label: str, ok: bool, detail: str = "") -> None:
    (PASS if ok else FAIL).append(f"{label}{(' :: ' + detail) if detail else ''}")
    log(f"[{'PASS' if ok else 'FAIL'}] {label}" + (f" :: {detail}" if detail else ""))


def main() -> int:
    workdir = RESULTS / "live"
    if workdir.exists():
        shutil.rmtree(workdir)
    workdir.mkdir(parents=True)

    # ---------------------------------------------------------------- fixture
    source = workdir / "sales.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"
    ws.append(["region", "amount"])
    for row in [("north", 100), ("south", 250), ("north", 100),
                ("east", 300), ("south", 250)]:
        ws.append(row)
    wb.save(source)
    source_hash = sha256(source)
    log(f"Fixture workbook: {source}")
    log(f"Source SHA256: {source_hash}")

    # ------------------------------------------------ agent pipeline (stubbed LLM)
    # Dedupe FIRST: if a formula column existed inside the range, its unique
    # value would defeat duplicate detection (tuple-based comparison).
    plan = {
        "summary": "Remove duplicate rows, label totals column, add SUM formula, style header",
        "steps": [
            {"type": "remove_duplicates", "sheet": "Sales", "range": "A1:C6"},
            {"type": "set_cell", "sheet": "Sales", "cell": "C1", "value": "total"},
            {"type": "write_formula", "sheet": "Sales", "cell": "C2", "value": "=SUM(B2:B4)"},
            {"type": "set_header_style", "sheet": "Sales", "row": 1},
        ],
        "risks": ["remove_duplicates deletes repeated rows"],
    }
    approvals: list[dict] = []
    ledger = Ledger(workdir / "ledger.sqlite3")
    service = AgentService(log=log, ledger=ledger,
                           provider=type("StubProvider", (), {
                               "complete": lambda self, s, u: json.dumps(plan)})())

    # Keep checkpoints inside test_results by mutating the shared settings
    # object (agent.py holds a reference to it, not to the module attribute).
    from openhuman_excel import config
    real_data_dir = config.settings.data_dir
    config.settings.data_dir = workdir

    result = service.inspect_and_plan(source, "prepare the sales sheet")
    copy = Path(result["working_copy"])
    log(f"Working copy: {copy}")
    check("working copy created next to results", workdir in copy.parents)
    check("original workbook untouched after plan (hash equal)",
          sha256(source) == source_hash)

    applied = service.run_plan(result, approve=lambda a: (approvals.append(a) or True))
    log(f"Executed actions: {applied['executed']}; approvals requested: {len(approvals)}")
    check("4 actions executed", applied["executed"] == 4, f"executed={applied['executed']}")
    check("destructive action asked for approval",
          any(a.get("type") == "remove_duplicates" for a in approvals))
    check("task completed in ledger",
          service.ledger.get_task_state(result["task_id"]) == "completed")
    check("original workbook untouched after execution (hash equal)",
          sha256(source) == source_hash)

    # ------------------------------------------------- openpyxl verification
    wb2 = load_workbook(copy)
    ws2 = wb2["Sales"]
    check("cell C1 written", ws2["C1"].value == "total")
    check("formula written", ws2["C2"].value == "=SUM(B2:B4)")
    rows = [[ws2.cell(r, 1).value, ws2.cell(r, 2).value] for r in range(1, 5)]
    log(f"Rows after dedupe: {rows}")
    check("duplicates removed (7 rows -> 4 rows with header)",
          ws2.max_row == 4 and rows[1] == ["north", 100] and rows[2] == ["south", 250])
    check("header styled bold", ws2["A1"].font.bold is True)
    wb2.close()

    # ------------------------------------------------- real Excel verification
    excel_ok = True
    try:
        verification = verify_formulas_in_excel(copy, {"Sales": ["C2"]})
        excel_ok = bool(verification.get("ok"))
        log(f"Excel verification: {verification}")
        check("Excel recalculated formula", excel_ok and
              verification["values"]["Sales"]["C2"] == 650,
              f"value={verification.get('values', {}).get('Sales', {}).get('C2')}")
    except Exception as exc:
        excel_ok = False
        log(f"Excel verification unavailable: {type(exc).__name__}: {exc}")
        check("Excel recalculated formula", False, str(exc))

    if excel_ok:
        com = ExcelCom(visible=True)
        com.open(copy)
        com.recalculate()
        value = com.cell_values("Sales", ["C2"])["C2"]
        log(f"Visible Excel shows C2 = {value}")
        check("visible Excel shows recalculated total", value == 650, f"value={value}")
        time.sleep(2)
        screenshot("01_excel_visible_working_copy.png",
                   window_title=copy.stem)
        com.close(save=False)
        log("Visible Excel instance closed without saving.")

    # ------------------------------------------------- xlsm / VBA retention
    # NOTE: a live "macro survives and still runs in Excel" test needs either a
    # real macro workbook or Excel's "Trust access to the VBA project object
    # model" (denied on this machine), so VBA retention is verified at the
    # package level: openpyxl must carry vbaProject.bin through an edit cycle.
    xlsm_source = workdir / "macros.xlsm"
    import zipfile
    stub_xlsx = workdir / "vba-base.xlsx"
    make_stub = Workbook()
    make_stub.active.title = "Sheet1"
    make_stub.active["A1"] = "vba-source"
    make_stub.save(stub_xlsx)
    marker = b"PK-MARKER-FAKE-VBA-PROJECT-BIN"
    with zipfile.ZipFile(stub_xlsx) as zin, \
            zipfile.ZipFile(xlsm_source, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            zout.writestr(item, zin.read(item.filename))
        zout.writestr("xl/vbaProject.bin", marker)
    log(f"xlsm fixture with stub vbaProject.bin: {xlsm_source}")
    xlsm_hash = sha256(xlsm_source)

    xlsm_copy = workdir / "macros-working.xlsm"
    shutil.copy2(xlsm_source, xlsm_copy)
    from openhuman_excel.executor import WorkbookExecutor
    executor = WorkbookExecutor(ledger, log=log, pace_seconds=0)
    ledger.create_task("xlsm-task", "edit macro workbook", str(xlsm_copy))
    n = executor.execute(xlsm_copy, "xlsm-task",
                         [{"type": "set_cell", "sheet": "Sheet1", "cell": "A2",
                           "value": "edited-by-agent"}], lambda a: True)
    check("xlsm edited by executor", n == 1)

    with zipfile.ZipFile(xlsm_copy) as zf:
        retained = "xl/vbaProject.bin" in zf.namelist() and \
            zf.read("xl/vbaProject.bin") == marker
    check("vbaProject.bin carried through openpyxl edit unchanged", retained)
    check("original xlsm untouched", sha256(xlsm_source) == xlsm_hash)

    # ------------------------------------------------- final integrity
    check("source workbook hash unchanged at end", sha256(source) == source_hash)

    # ------------------------------------------------- report
    log("")
    log(f"RESULTS: {len(PASS)} passed, {len(FAIL)} failed")
    (RESULTS / "live_verification_log.txt").write_text(
        "\n".join(LOG_LINES) + "\n", encoding="utf-8")
    if FAIL:
        for line in FAIL:
            log(f"FAILED: {line}")
    config.settings.data_dir = real_data_dir
    return 1 if FAIL else 0


if __name__ == "__main__":
    sys.exit(main())
