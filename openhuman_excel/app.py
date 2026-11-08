from __future__ import annotations

import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

from openhuman_excel.agent import AgentService


class OpenHumanExcelApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("OpenHuman Excel")
        self.root.geometry("900x650")
        self.root.minsize(760, 540)
        self.service = AgentService(self.log)
        self.workbook_path: Path | None = None
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
        ttk.Button(controls, text="Pause", command=self.service.pause).pack(side="left", padx=6)
        ttk.Button(controls, text="Resume", command=self.service.resume).pack(side="left")
        ttk.Button(controls, text="Stop", command=self.service.stop).pack(side="left", padx=6)
        ttk.Label(outer, text="Live activity").pack(anchor="w")
        self.output = tk.Text(outer, height=18, state="disabled", wrap="word",
                              bg="#101820", fg="#d9e4ee", font=("Consolas", 9))
        self.output.pack(fill="both", expand=True)
        ttk.Label(outer, text="Safe mode: works on a copy. This first milestone inspects and plans; it does not apply edits.",
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
            ("Excel workbooks", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")])
        if filename:
            self.workbook_path = Path(filename)
            self.file_label.configure(text=str(self.workbook_path))

    def start(self):
        if not self.workbook_path:
            messagebox.showwarning("Select workbook", "Choose an Excel workbook first.")
            return
        instruction = self.instruction.get("1.0", "end").strip()
        if not instruction:
            messagebox.showwarning("Task required", "Describe the work you want done.")
            return
        self.start_button.configure(state="disabled")
        self.log("Starting safe inspection and planning...")
        def worker():
            try:
                result = self.service.inspect_and_plan(self.workbook_path, instruction)
                self.log(f"Plan for task {result['task_id']}:")
                self.log(str(result["plan"]))
                self.log(f"Working copy: {result['working_copy']}")
            except Exception as exc:
                self.log(f"ERROR: {type(exc).__name__}: {exc}")
            finally:
                self.root.after(0, lambda: self.start_button.configure(state="normal"))
        threading.Thread(target=worker, daemon=True).start()


def main():
    root = tk.Tk()
    OpenHumanExcelApp(root)
    root.mainloop()
