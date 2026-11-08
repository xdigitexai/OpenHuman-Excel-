"""GUI smoke test: launch the real Tkinter app and drive a full task through it.

Uses a stub provider so no API key or network is needed. Screenshots are saved
to test_results/. Run: ./.venv/Scripts/python.exe scripts/gui_smoke_test.py
"""
from __future__ import annotations

import json
import shutil
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from live_excel_verification import focus_window

RESULTS = ROOT / "test_results"
RESULTS.mkdir(exist_ok=True)
WORK = RESULTS / "gui"
if WORK.exists():
    shutil.rmtree(WORK)
WORK.mkdir(parents=True)

PASS, FAIL = [], []


def check(label: str, ok: bool, detail: str = ""):
    (PASS if ok else FAIL).append(label)
    print(f"[{'PASS' if ok else 'FAIL'}] {label}" + (f" :: {detail}" if detail else ""),
          flush=True)


def screenshot(name: str) -> None:
    focus_window("OpenHuman Excel")
    target = RESULTS / name
    ps = (
        "Add-Type -AssemblyName System.Windows.Forms,System.Drawing;"
        "$b = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds;"
        "$bmp = New-Object System.Drawing.Bitmap $b.Width, $b.Height;"
        "$g = [System.Drawing.Graphics]::FromImage($bmp);"
        "$g.CopyFromScreen($b.Location, [System.Drawing.Point]::Empty, $b.Size);"
        f"$bmp.Save('{target.as_posix()}');"
        "$g.Dispose(); $bmp.Dispose()"
    )
    subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                   check=True, timeout=30)
    print(f"Screenshot saved: {target}", flush=True)


def main() -> int:
    import tkinter as tk
    import tkinter.messagebox as mb

    from openpyxl import Workbook, load_workbook

    from openhuman_excel import config
    from openhuman_excel.app import OpenHumanExcelApp

    source = WORK / "inventory.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Stock"
    ws.append(["item", "qty"])
    ws.append(["widget", 5])
    ws.append(["gadget", 0])
    wb.save(source)

    config.settings.data_dir = WORK
    plan = {"summary": "Flag zero-stock items", "steps": [
        {"type": "set_cell", "sheet": "Stock", "cell": "C1", "value": "restock"},
        {"type": "write_formula", "sheet": "Stock", "cell": "C3",
         "value": "=IF(B3=0,\"YES\",\"no\")"},
    ], "risks": []}

    root = tk.Tk()
    app = OpenHumanExcelApp(root)
    # Mirror the GUI's live activity log to stdout for diagnosis.
    original_append = app._append_log
    app._append_log = lambda message: (original_append(message),
                                       print(f"  GUI| {message}", flush=True))
    app.service.provider = type("Stub", (), {
        "complete": lambda self, s, u: json.dumps(plan)})()
    mb.askyesno = lambda *a, **k: True  # auto-approve approval dialogs

    app.workbook_path = source
    app.instruction.delete("1.0", "end")
    app.instruction.insert("1.0", "Add a restock flag column.")

    state = {"phase": "initial", "deadline": 0.0}
    phases: dict[str, object] = {}

    def wait_until(phase, condition, next_phase, timeout=30, then=None):
        state["phase"] = phase
        state["deadline"] = time.time() + timeout
        state["next"] = next_phase
        state["then"] = then
        state["condition"] = condition
        poll()

    def poll():
        if state["phase"] == "done":
            return
        if time.time() > state["deadline"]:
            check(f"timeout waiting for {state['next']}", False)
            root.destroy()
            return
        if state["condition"]():
            if state["then"]:
                state["then"]()
            state["phase"] = state["next"]
            root.after(200, phases[state["next"]])
        else:
            root.after(150, poll)

    def phase_plan_ready():
        check("plan arrived and Apply plan enabled",
              str(app.apply_button["state"]) == "normal")
        check("pause/stop were enabled during the run",
              state.get("controls_seen_active", False))
        screenshot("03_gui_plan_ready.png")
        app.apply_plan()
        wait_until("applying", lambda: not app._busy, "phase_applied", timeout=90)

    def phase_applied():
        check("apply finished and controls re-enabled", not app._busy)
        screenshot("04_gui_after_apply.png")
        copies = list((WORK / "checkpoints").glob("inventory.openh-*.xlsx"))
        check("working copy exists", len(copies) == 1)
        if copies:
            check_wb = load_workbook(copies[0])
            check("working copy edited via GUI",
                  check_wb["Stock"]["C1"].value == "restock",
                  f"C1={check_wb['Stock']['C1'].value}")
            check_wb.close()
        check("original workbook preserved",
              load_workbook(source)["Stock"]["C1"].value is None)
        check("task completed in ledger",
              app.service.ledger.get_task_state(app.service.current_task_id)
              == "completed")
        # Stop pressed now must interrupt the NEXT run.
        app.service.stop()
        try:
            app.service.inspect_and_plan(source, "another task")
            check("stop interrupts a fresh run", False, "no InterruptedError")
        except InterruptedError:
            check("stop interrupts a fresh run", True)
        finish()

    def finish():
        state["phase"] = "done"
        print(f"\nGUI RESULTS: {len(PASS)} passed, {len(FAIL)} failed", flush=True)
        root.after(100, root.destroy)

    phases["phase_plan_ready"] = phase_plan_ready
    phases["phase_applied"] = phase_applied

    def phase_initial():
        check("GUI launched with controls", app.start_button["text"] == "Start"
              and str(app.pause_button["state"]) == "disabled")
        screenshot("02_gui_started.png")
        app.start()
        wait_until("planning",
                   lambda: (state.update(controls_seen_active=(
                       state.get("controls_seen_active", False)
                       or str(app.stop_button["state"]) == "normal"))
                       or str(app.apply_button["state"]) == "normal"),
                   "phase_plan_ready")

    root.after(300, phase_initial)
    root.mainloop()
    return 1 if FAIL else 0


if __name__ == "__main__":
    sys.exit(main())
