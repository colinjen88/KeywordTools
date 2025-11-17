#!/usr/bin/env python3
"""
簡易 GUI 執行器：讓使用者透過視覺介面輸入 Search Console property、起訖日與輸出格式，並執行 gsc_keyword_report.py

功能：
- 輸入欄位：property、start-date、end-date
- 選項：mock 模式開關、輸出為 CSV / XLSX（可同時勾選）
- 執行後會在下方顯示執行 log

用法：
  python run_gui.py

注意：若選 XLSX 輸出，需要安裝 `pandas` 與 `openpyxl`（已列在 `requirements.txt`）。
"""
import subprocess
import sys
import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


SCRIPT = "gsc_keyword_report.py"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("GSC Keyword Reporter - GUI")
        self.geometry("760x480")

        frm = ttk.Frame(self, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frm, text="Property (Search Console URL):").grid(row=0, column=0, sticky=tk.W)
        self.property_var = tk.StringVar(value="https://pm.shiny.com.tw/")
        ttk.Entry(frm, textvariable=self.property_var, width=60).grid(row=0, column=1, columnspan=3, sticky=tk.W)

        ttk.Label(frm, text="Start date (YYYY-MM-DD):").grid(row=1, column=0, sticky=tk.W)
        self.start_var = tk.StringVar(value="2025-10-01")
        ttk.Entry(frm, textvariable=self.start_var, width=20).grid(row=1, column=1, sticky=tk.W)

        ttk.Label(frm, text="End date (YYYY-MM-DD):").grid(row=1, column=2, sticky=tk.W)
        self.end_var = tk.StringVar(value="2025-10-31")
        ttk.Entry(frm, textvariable=self.end_var, width=20).grid(row=1, column=3, sticky=tk.W)

        self.mock_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(frm, text="Use mock data (no GSC auth)", variable=self.mock_var).grid(row=2, column=0, columnspan=2, sticky=tk.W)

        ttk.Label(frm, text="Keywords file:").grid(row=3, column=0, sticky=tk.W)
        self.kws_var = tk.StringVar(value="allKeyWord_normalized.csv")
        ttk.Entry(frm, textvariable=self.kws_var, width=40).grid(row=3, column=1, sticky=tk.W)
        ttk.Button(frm, text="Browse", command=self.browse_kws).grid(row=3, column=2, sticky=tk.W)

        ttk.Label(frm, text="Output base name:").grid(row=4, column=0, sticky=tk.W)
        self.outbase_var = tk.StringVar(value="gsc_keyword_report")
        ttk.Entry(frm, textvariable=self.outbase_var, width=30).grid(row=4, column=1, sticky=tk.W)

        self.csv_var = tk.BooleanVar(value=True)
        self.xlsx_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(frm, text="CSV", variable=self.csv_var).grid(row=5, column=1, sticky=tk.W)
        ttk.Checkbutton(frm, text="Excel (.xlsx)", variable=self.xlsx_var).grid(row=5, column=2, sticky=tk.W)

        self.run_btn = ttk.Button(frm, text="Run Report", command=self.on_run)
        self.run_btn.grid(row=6, column=1, pady=8, sticky=tk.W)

        self.log = tk.Text(frm, height=18)
        self.log.grid(row=7, column=0, columnspan=4, pady=6, sticky=tk.NSEW)
        frm.rowconfigure(7, weight=1)
        frm.columnconfigure(3, weight=1)

    def browse_kws(self):
        p = filedialog.askopenfilename(initialdir='.', filetypes=[('CSV files','*.csv'),('All files','*.*')])
        if p:
            self.kws_var.set(p)

    def append_log(self, text):
        self.log.insert(tk.END, text + "\n")
        self.log.see(tk.END)

    def on_run(self):
        prop = self.property_var.get().strip()
        start = self.start_var.get().strip()
        end = self.end_var.get().strip()
        kws = self.kws_var.get().strip() or 'allKeyWord_normalized.csv'
        base = self.outbase_var.get().strip() or 'gsc_keyword_report'
        use_mock = self.mock_var.get()
        do_csv = self.csv_var.get()
        do_xlsx = self.xlsx_var.get()

        if not prop or not start or not end:
            messagebox.showerror('Missing', 'Please provide property, start date and end date')
            return

        self.run_btn.config(state=tk.DISABLED)
        self.log.delete('1.0', tk.END)

        def worker():
            try:
                outputs = []
                if do_csv:
                    out = base + '.csv'
                    cmd = [sys.executable, SCRIPT, '--property', prop, '--keywords', kws, '--start-date', start, '--end-date', end, '--output', out]
                    if use_mock:
                        cmd.append('--mock')
                    self.append_log('Running: ' + ' '.join(cmd))
                    proc = subprocess.run(cmd, capture_output=True, text=True)
                    self.append_log(proc.stdout)
                    if proc.stderr:
                        self.append_log(proc.stderr)
                    outputs.append(out)

                if do_xlsx:
                    out = base + '.xlsx'
                    cmd = [sys.executable, SCRIPT, '--property', prop, '--keywords', kws, '--start-date', start, '--end-date', end, '--output', out]
                    if use_mock:
                        cmd.append('--mock')
                    self.append_log('Running: ' + ' '.join(cmd))
                    proc = subprocess.run(cmd, capture_output=True, text=True)
                    self.append_log(proc.stdout)
                    if proc.stderr:
                        self.append_log(proc.stderr)
                    outputs.append(out)

                for f in outputs:
                    if os.path.exists(f):
                        self.append_log(f'Generated: {f}')
                    else:
                        self.append_log(f'Failed to generate: {f}')
            except Exception as e:
                self.append_log('Error: ' + str(e))
            finally:
                self.run_btn.config(state=tk.NORMAL)

        threading.Thread(target=worker, daemon=True).start()


if __name__ == '__main__':
    app = App()
    app.mainloop()
