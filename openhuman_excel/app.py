from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from openhuman_excel.agent import AgentService


class OpenHumanExcelApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("OpenHuman Excel")
        self.root.geometry("900x650")
        self.root.minsize(760, 540)
        self.service = AgentService(self.log)
        self.workbook_path: Path | None = None
        self._last_plan: dict | None = None
        self._busy = False
        self._build()

    def _build(self):
        outer = ttk.Frame(self.root, padding=18)
        outer.pack(fill="both", expand=True)
        ttk.Label(outer, text="OpenHuman Excel", font=("Segoe UI", 22, "bold")).pack(anchor="w")
        ttk.Label(outer, text="A visible desktop agent for real Excel work",
                  font=("Segoe UI", 10)).pack(anchor="w", pady=(0, 14))

        file_row = ttk.Frame(outer)
        file_row.pack(fill="x", pady=5)
        self.file_label = ttk.Label(file_row, text="No workbook selected")
        self.file_label.pack(side="left", fill="x", expand=True)
        ttk.Button(file_row, text="Choose workbook", command=self.choose_workbook).pack(side="right")

        ttk.Label(outer, text="Describe exactly what you want done").pack(anchor="w", pady=(14, 4))
        self.instruction = tk.Text(outer, height=5, wrap="word", font=("Segoe UI", 10))
        self.instruction.pack(fill="x")
        self.instruction.insert("1.0", "Inspect this workbook and propose a safe, step-by-step plan. Do not edit it yet.")

        controls = ttk.Frame(outer)
        controls.pack(fill="x", pady=12)
        self.start_button = ttk.Button(controls, text="Start", command=self.start)
        self.start_button.pack(side="left")
        self.pause_button = ttk.Button(controls, text="Pause", command=self.service.pause,
                                       state="disabled")
        self.pause_button.pack(side="left", padx=6)
        self.resume_button = ttk.Button(controls, text="Resume", command=self.service.resume,
                                        state="disabled")
        self.resume_button.pack(side="left")
        self.stop_button = ttk.Button(controls, text="Stop", command=self.service.stop,
                                      state="disabled")
        self.stop_button.pack(side="left", padx=6)
        self.apply_button = ttk.Button(controls, text="Apply plan", command=self.apply_plan,
                                       state="disabled")
        self.apply_button.pack(side="left", padx=(18, 0))
        ttk.Label(outer, text="Live activity").pack(anchor="w")
        self.output = tk.Text(outer, height=18, state="disabled", wrap="word",
                              bg="#101820", fg="#d9e4ee", font=("Consolas", 9))
        self.output.pack(fill="both", expand=True)
        ttk.Label(outer, text="Safe mode: edits happen in a working copy. Destructive actions ask for approval first.",
                  foreground="#555").pack(anchor="w", pady=(8, 0))

    def log(self, message: str):
        self.root.after(0, self._append_log, message)

    def _append_log(self, message: str):
        self.output.configure(state="normal")
        self.output.insert("end", message + "\n")
        self.output.see("end")
        self.output.configure(state="disabled")

    def choose_workbook(self):
        filename = filedialog.askopenfilename(filetypes=[
            ("Excel workbooks", "*.xlsx *.xlsm"), ("All files", "*.*")])
        if filename:
            self.workbook_path = Path(filename)
            self.file_label.configure(text=str(self.workbook_path))

    def _set_running(self, running: bool):
        self._busy = running
        state = "disabled" if running else "normal"
        self.start_button.configure(state=state)
        self.apply_button.configure(state=state)
        self.pause_button.configure(state="normal" if running else "disabled")
        self.stop_button.configure(state="normal" if running else "disabled")
        self.resume_button.configure(state="disabled")

    def _on_pause(self):
        # Tk button commands run on the UI thread; the service flags are read
        # by the worker thread, so no additional locking is needed here.
        self.service.pause()
        self.pause_button.configure(state="disabled")
        self.resume_button.configure(state="normal")
        self.log("Paused. Press Resume to continue.")

    def _on_resume(self):
        self.service.resume()
        self.resume_button.configure(state="disabled")
        self.pause_button.configure(state="normal")

    def start(self):
        if self._busy:
            return
        if not self.workbook_path:
            messagebox.showwarning("Select workbook", "Choose an Excel workbook first.")
            return
        instruction = self.instruction.get("1.0", "end").strip()
        if not instruction:
            messagebox.showwarning("Task required", "Describe the work you want done.")
            return
        self._set_running(True)
        self.service.reset_controls()
        self.log("Starting safe inspection and planning...")

        def worker():
            try:
                result = self.service.inspect_and_plan(self.workbook_path, instruction)
                self._last_plan = result
                plan = result["plan"]
                self.log(f"Plan for task {result['task_id']}:")
                self.log(str(plan))
                self.log(f"Working copy: {result['working_copy']}")
                self.root.after(0, self._on_plan_ready)
            except InterruptedError:
                self.log("Stopped during planning; the working copy was kept for review.")
            except Exception as exc:
                self.log(f"ERROR: {type(exc).__name__}: {exc}")
            finally:
                self.root.after(0, lambda: self._set_running(False))

        threading.Thread(target=worker, daemon=True).start()

    def _on_plan_ready(self):
        self.apply_button.configure(
            state="normal" if self._last_plan else "disabled")

    def apply_plan(self):
        if self._busy or not self._last_plan:
            return
        plan = self._last_plan["plan"]
        steps = plan.get("steps") if isinstance(plan, dict) else None
        if not steps and not messagebox.askyesno(
                "No steps", "This plan has no executable steps. Apply anyway?"):
            return
        self._set_running(True)
        self.service.reset_controls()
        self.log("Applying plan to the working copy...")

        def worker():
            try:
                result = self.service.run_plan(self._last_plan, approve=self._ask_approval)
                self.log(f"Executed {result['executed']} action(s).")
                if result.get("issues"):
                    self.log(f"Rejected steps: {len(result['issues'])}")
                if result.get("verification"):
                    self.log("Excel verification result: "
                             f"ok={result['verification'].get('ok')}")
            except InterruptedError:
                self.log("Stopped. Progress so far is saved in the working copy "
                         "and can be resumed by applying the plan again.")
            except Exception as exc:
                self.log(f"ERROR: {type(exc).__name__}: {exc}")
            finally:
                self.root.after(0, lambda: self._set_running(False))

        threading.Thread(target=worker, daemon=True).start()

    def _ask_approval(self, action: dict) -> bool:
        """Ask on the UI thread; block the worker until the user answers."""
        done = threading.Event()
        answer = {"approved": False}

        def ask():
            answer["approved"] = messagebox.askyesno(
                "Approval required",
                f"The agent requests approval for:\n\n{action.get('type')}: "
                f"{action.get('sheet', '')} {action.get('range', '') or action.get('new_name', '') or action.get('cell', '')}\n\n"
                "Allow this action?")
            done.set()

        self.root.after(0, ask)
        done.wait()
        return answer["approved"]


def main():
    root = tk.Tk()
    OpenHumanExcelApp(root)
    root.mainloop()
