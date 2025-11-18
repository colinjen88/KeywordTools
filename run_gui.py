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
import csv
from datetime import date, timedelta


SCRIPT = "gsc_keyword_report.py"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("GSC Keyword Reporter - GUI")
        self.geometry("760x480")

        frm = ttk.Frame(self, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frm, text="Search Console 屬性 (URL)：").grid(row=0, column=0, sticky=tk.W)
        self.property_var = tk.StringVar(value="https://pm.shiny.com.tw/")
        ttk.Entry(frm, textvariable=self.property_var, width=60).grid(row=0, column=1, columnspan=3, sticky=tk.W)

        ttk.Label(frm, text="起始日期（YYYY-MM-DD）：").grid(row=1, column=0, sticky=tk.W)
        self.start_var = tk.StringVar(value=(date.today() - timedelta(days=30)).isoformat())
        ttk.Entry(frm, textvariable=self.start_var, width=20).grid(row=1, column=1, sticky=tk.W)

        ttk.Label(frm, text="結束日期（YYYY-MM-DD）：").grid(row=1, column=2, sticky=tk.W)
        self.end_var = tk.StringVar(value=date.today().isoformat())
        ttk.Entry(frm, textvariable=self.end_var, width=20).grid(row=1, column=3, sticky=tk.W)

        self.mock_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(frm, text="使用 mock 範例資料（不呼叫 GSC API）", variable=self.mock_var).grid(row=2, column=0, columnspan=2, sticky=tk.W)

        # preset ranges
        preset_frame = ttk.Frame(frm)
        preset_frame.grid(row=2, column=2, columnspan=2, sticky=tk.E)
        ttk.Label(preset_frame, text="快速區間：").grid(row=0, column=0, sticky=tk.W)
        ttk.Button(preset_frame, text="近7天", command=lambda: self.set_preset(7)).grid(row=0, column=1, padx=4)
        ttk.Button(preset_frame, text="近30天", command=lambda: self.set_preset(30)).grid(row=0, column=2, padx=4)
        ttk.Button(preset_frame, text="近1季", command=lambda: self.set_preset(90)).grid(row=0, column=3, padx=4)
        ttk.Button(preset_frame, text="近1年", command=lambda: self.set_preset(365)).grid(row=0, column=4, padx=4)

        ttk.Label(frm, text="關鍵字檔案：").grid(row=3, column=0, sticky=tk.W)
        self.kws_var = tk.StringVar(value="allKeyWord_normalized.csv")
        ttk.Entry(frm, textvariable=self.kws_var, width=40).grid(row=3, column=1, sticky=tk.W)
        ttk.Button(frm, text="Browse", command=self.browse_kws).grid(row=3, column=2, sticky=tk.W)

        ttk.Label(frm, text="Service account JSON（選填）：").grid(row=4, column=0, sticky=tk.W)
        self.sa_var = tk.StringVar(value="")
        ttk.Entry(frm, textvariable=self.sa_var, width=40).grid(row=4, column=1, sticky=tk.W)
        ttk.Button(frm, text="瀏覽", command=self.browse_sa).grid(row=4, column=2, sticky=tk.W)

        ttk.Label(frm, text="輸出檔案基底名稱：").grid(row=5, column=0, sticky=tk.W)
        self.outbase_var = tk.StringVar(value="gsc_keyword_report")
        ttk.Entry(frm, textvariable=self.outbase_var, width=30).grid(row=5, column=1, sticky=tk.W)

        ttk.Label(frm, text="輸出格式：").grid(row=6, column=0, sticky=tk.W)
        self.format_var = tk.StringVar(value='CSV')
        fmt_combo = ttk.Combobox(frm, textvariable=self.format_var, values=['CSV', 'Excel (.xlsx)'], state='readonly', width=18)
        fmt_combo.grid(row=6, column=1, sticky=tk.W)

        # keep legacy run_btn for compatibility (hidden)
        self.run_btn = ttk.Button(frm, text="執行報表", command=self.on_run)
        # hide original small run_btn (we use larger one in button frame)
        self.run_btn.grid_forget()

        self.log = tk.Text(frm, height=18)
        self.log.grid(row=8, column=0, columnspan=4, pady=6, sticky=tk.NSEW)
        frm.rowconfigure(8, weight=1)
        frm.columnconfigure(3, weight=1)

        # results frame
        ttk.Label(frm, text="結果：").grid(row=9, column=0, sticky=tk.W, pady=(8,0))
        # status label (left of results)
        self.status_var = tk.StringVar(value='待命')
        self.status_label = tk.Label(frm, text='狀態：待命', bg='#808080', fg='white', padx=8, pady=2)
        self.status_label.grid(row=9, column=1, sticky=tk.W, padx=(8,0))

        # statistics frame (right side)
        stats_frame = ttk.Frame(frm, relief=tk.RIDGE)
        stats_frame.grid(row=9, column=2, columnspan=2, sticky=tk.EW, padx=(8,0))
        stats_frame.columnconfigure(0, weight=1)
        ttk.Label(stats_frame, text='統計').grid(row=0, column=0, sticky=tk.W)
        self.stat_kw_var = tk.StringVar(value='關鍵字數: 0')
        self.stat_clicks_var = tk.StringVar(value='總點擊: 0')
        self.stat_impr_var = tk.StringVar(value='總曝光: 0')
        self.stat_pos_var = tk.StringVar(value='平均排名: -')
        ttk.Label(stats_frame, textvariable=self.stat_kw_var).grid(row=1, column=0, sticky=tk.W)
        ttk.Label(stats_frame, textvariable=self.stat_clicks_var).grid(row=2, column=0, sticky=tk.W)
        ttk.Label(stats_frame, textvariable=self.stat_impr_var).grid(row=3, column=0, sticky=tk.W)
        ttk.Label(stats_frame, textvariable=self.stat_pos_var).grid(row=4, column=0, sticky=tk.W)

        self.table_frame = ttk.Frame(frm)
        self.table_frame.grid(row=10, column=0, columnspan=4, sticky=tk.NSEW)
        frm.rowconfigure(10, weight=1)

        self.tree = None
        self.current_rows = []
        self.current_columns = []
        btn_frame = ttk.Frame(frm)
        btn_frame.grid(row=11, column=0, columnspan=4, sticky=tk.W, pady=6)
        # enlarge Run button style
        try:
            style = ttk.Style()
            style.configure('Big.TButton', font=('Segoe UI', 10, 'bold'), padding=(12,8))
        except Exception:
            pass

        self.save_btn = ttk.Button(btn_frame, text="輸出檔案", command=self.export_csv)
        self.save_btn.grid(row=0, column=0, padx=(0,8))
        self.clear_btn = ttk.Button(btn_frame, text="清除表格", command=self.clear_table)
        self.clear_btn.grid(row=0, column=1, padx=(0,8))
        # autoload toggle
        self.autoload_var = tk.BooleanVar(value=True)
        self.autoload_cb = ttk.Checkbutton(btn_frame, text='自動載入 CSV', variable=self.autoload_var)
        self.autoload_cb.grid(row=0, column=3, padx=(8,8))
        # Run button bigger and styled
        self.run_btn_big = ttk.Button(btn_frame, text="執行報表", command=self.on_run, style='Big.TButton')
        self.run_btn_big.grid(row=0, column=2, padx=(12,8))

        # start file watcher to auto-load CSV created externally
        try:
            self.start_file_watcher()
        except Exception:
            pass

    def browse_kws(self):
        p = filedialog.askopenfilename(initialdir='.', filetypes=[('CSV files','*.csv'),('All files','*.*')])
        if p:
            self.kws_var.set(p)

    def browse_sa(self):
        p = filedialog.askopenfilename(initialdir='.', filetypes=[('JSON files','*.json'),('All files','*.*')])
        if p:
            self.sa_var.set(p)

    def set_preset(self, days:int):
        end = date.today()
        start = end - timedelta(days=days-1)
        self.start_var.set(start.isoformat())
        self.end_var.set(end.isoformat())

    def clear_table(self):
        if self.tree:
            for item in self.tree.get_children():
                self.tree.delete(item)
            self.tree.destroy()
            self.tree = None
            self.current_rows = []
            self.current_columns = []

    def load_csv_into_table(self, path, max_rows=10000):
        # read CSV and populate Treeview
        with open(path, newline='', encoding='utf-8-sig') as fh:
            reader = csv.reader(fh)
            try:
                header = next(reader)
            except StopIteration:
                header = []

            rows = []
            for i, r in enumerate(reader):
                rows.append(r)
                if i+1 >= max_rows:
                    break

        # clear existing
        self.clear_table()

        # map headers to Chinese columns if possible
        src_cols = [c.strip().lower() for c in header]
        # find indices
        def idx(names):
            for n in names:
                if n in src_cols:
                    return src_cols.index(n)
            return None

        idx_keyword = idx(['keyword', 'query'])
        idx_clicks = idx(['clicks', 'click'])
        idx_impr = idx(['impressions', 'impression'])
        idx_pos = idx(['position', 'avg_position', 'pos'])

        # '搜尋' 改為更常用的名稱 '曝光'
        display_cols = ['關鍵字', '點擊', '曝光', '排名']
        mapped_rows = []
        for r in rows:
            mapped = []
            mapped.append(r[idx_keyword] if idx_keyword is not None and idx_keyword < len(r) else '')
            mapped.append(r[idx_clicks] if idx_clicks is not None and idx_clicks < len(r) else '')
            mapped.append(r[idx_impr] if idx_impr is not None and idx_impr < len(r) else '')
            mapped.append(r[idx_pos] if idx_pos is not None and idx_pos < len(r) else '')
            mapped_rows.append(mapped)

        self.current_columns = display_cols
        self.current_rows = mapped_rows

        # create tree (height shows 20 rows)
        tree = ttk.Treeview(self.table_frame, columns=display_cols, show='headings', height=20)
        vsb = ttk.Scrollbar(self.table_frame, orient='vertical', command=tree.yview)
        hsb = ttk.Scrollbar(self.table_frame, orient='horizontal', command=tree.xview)
        tree.configure(yscroll=vsb.set, xscroll=hsb.set)
        tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        self.table_frame.rowconfigure(0, weight=1)
        self.table_frame.columnconfigure(0, weight=1)

        for c in display_cols:
            tree.heading(c, text=c)
            tree.column(c, width=160, anchor='w')

        for r in mapped_rows:
            tree.insert('', tk.END, values=r)

        self.tree = tree

    def export_csv(self):
        # unified export: use selected format
        if not self.current_columns:
            messagebox.showinfo('無資料', '目前表格沒有資料可匯出')
            return
        fmt = self.format_var.get() if hasattr(self, 'format_var') else 'CSV'
        if fmt == 'CSV':
            p = filedialog.asksaveasfilename(defaultextension='.csv', filetypes=[('CSV','*.csv')], initialfile=self.outbase_var.get() + '.csv')
            if not p:
                return
            try:
                with open(p, 'w', newline='', encoding='utf-8-sig') as fh:
                    writer = csv.writer(fh)
                    writer.writerow(self.current_columns)
                    for r in self.current_rows:
                        writer.writerow(r)
                messagebox.showinfo('已儲存', f'已儲存 CSV 到 {p}')
            except Exception as e:
                messagebox.showerror('錯誤', str(e))
        else:
            # Excel
            p = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel','*.xlsx')], initialfile=self.outbase_var.get() + '.xlsx')
            if not p:
                return
            try:
                try:
                    import pandas as pd
                except Exception:
                    messagebox.showerror('缺少套件', '匯出 XLSX 需要安裝 pandas 和 openpyxl')
                    return
                df = pd.DataFrame(self.current_rows, columns=self.current_columns)
                df.to_excel(p, index=False)
                messagebox.showinfo('已儲存', f'已儲存 Excel 到 {p}')
            except Exception as e:
                messagebox.showerror('錯誤', str(e))

    def append_log(self, text):
        self.log.insert(tk.END, text + "\n")
        self.log.see(tk.END)

    def set_status(self, text: str, color: str):
        # thread-safe status update
        def _update():
            # map basic color names to hex for better visibility
            color_map = {
                'green': '#2e7d32',
                'blue': '#1565c0',
                'red': '#c62828'
            }
            bg = color_map.get(color, color if color and color.startswith('#') else '#808080')
            self.status_label.config(text=f'狀態：{text}', bg=bg)
        try:
            self.after(0, _update)
        except Exception:
            pass

    def start_file_watcher(self):
        # start background thread to watch for new/updated CSV files and auto-load
        self._watch_stop = False
        self._watch_last_mtime = 0
        def watcher():
            import time, glob
            while not self._watch_stop:
                try:
                    csvs = glob.glob(os.path.join('.', '*.csv'))
                    if not csvs:
                        time.sleep(2)
                        continue
                    latest = max(csvs, key=os.path.getmtime)
                    try:
                        m = os.path.getmtime(latest)
                    except OSError:
                        m = 0
                    if m and m > self._watch_last_mtime:
                        self._watch_last_mtime = m
                        # schedule load on main thread
                        self.after(0, lambda p=latest: self._auto_load_if_needed(p))
                except Exception:
                    pass
                time.sleep(2)

        t = threading.Thread(target=watcher, daemon=True)
        t.start()

    def _auto_load_if_needed(self, path):
        # Only auto-load if table is empty or the file is different from current loaded
        try:
            # check autoload setting on main thread
            try:
                if hasattr(self, 'autoload_var') and not self.autoload_var.get():
                    return
            except Exception:
                pass
            if not os.path.exists(path):
                return
            if self.current_rows and os.path.abspath(path) == getattr(self, '_last_loaded_path', None):
                # already loaded
                return
            # load
            self.append_log(f'偵測到新 CSV：{path}，自動載入表格')
            self.load_csv_into_table(path)
            self._last_loaded_path = os.path.abspath(path)
            # update status
            try:
                self.set_status('查詢完成', 'blue')
            except Exception:
                pass
        except Exception as e:
            self.append_log('自動載入失敗: ' + str(e))


    def on_run(self):
        prop = self.property_var.get().strip()
        start = self.start_var.get().strip()
        end = self.end_var.get().strip()
        kws = self.kws_var.get().strip() or 'allKeyWord_normalized.csv'
        base = self.outbase_var.get().strip() or 'gsc_keyword_report'
        use_mock = self.mock_var.get()
        fmt = self.format_var.get() if hasattr(self, 'format_var') else 'CSV'

        if not prop or not start or not end:
            messagebox.showerror('缺少參數', '請提供 property、開始日期與結束日期')
            return

        # disable both run buttons (big and legacy) while running
        try:
            self.run_btn_big.config(state=tk.DISABLED)
        except Exception:
            pass
        try:
            self.run_btn.config(state=tk.DISABLED)
        except Exception:
            pass
        # set status to querying
        try:
            self.set_status('查詢中', 'green')
        except Exception:
            pass
        self.log.delete('1.0', tk.END)

        def worker():
            try:
                outputs = []
                sa_path = self.sa_var.get().strip() if hasattr(self, 'sa_var') else ''
                # Decide output based on selected format
                if fmt == 'CSV':
                    out = base + '.csv'
                    cmd = [sys.executable, SCRIPT, '--property', prop, '--keywords', kws, '--start-date', start, '--end-date', end, '--output', out]
                    if use_mock:
                        cmd.append('--mock')
                    else:
                        if sa_path:
                            cmd.extend(['--service-account', sa_path])
                    self.append_log('執行: ' + ' '.join(cmd))
                    proc = subprocess.run(cmd, capture_output=True, text=True)
                    self.append_log(proc.stdout)
                    if proc.stderr:
                        self.append_log(proc.stderr)
                    outputs.append(out)
                else:
                    out = base + '.xlsx'
                    cmd = [sys.executable, SCRIPT, '--property', prop, '--keywords', kws, '--start-date', start, '--end-date', end, '--output', out]
                    if use_mock:
                        cmd.append('--mock')
                    else:
                        if sa_path:
                            cmd.extend(['--service-account', sa_path])
                    self.append_log('執行: ' + ' '.join(cmd))
                    proc = subprocess.run(cmd, capture_output=True, text=True)
                    self.append_log(proc.stdout)
                    if proc.stderr:
                        self.append_log(proc.stderr)
                    outputs.append(out)

                for f in outputs:
                    if os.path.exists(f):
                        self.append_log(f'Generated: {f}')
                        # if CSV, load into table
                        if f.lower().endswith('.csv'):
                            try:
                                self.load_csv_into_table(f)
                            except Exception as e:
                                self.append_log('Failed to load CSV into table: ' + str(e))
                    else:
                        self.append_log(f'Failed to generate: {f}')
            except Exception as e:
                self.append_log('Error: ' + str(e))
                try:
                    self.set_status('錯誤', 'red')
                except Exception:
                    pass
            finally:
                try:
                    self.run_btn_big.config(state=tk.NORMAL)
                except Exception:
                    pass
                try:
                    self.run_btn.config(state=tk.NORMAL)
                except Exception:
                    pass
                # if no exception, set completed (if not already set to error)
                try:
                    self.set_status('查詢完成', 'blue')
                except Exception:
                    pass

        threading.Thread(target=worker, daemon=True).start()


if __name__ == '__main__':
    app = App()
    app.mainloop()

