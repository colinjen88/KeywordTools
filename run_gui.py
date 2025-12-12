#!/usr/bin/env python3
"""
KeywordTools V1.2
簡易 GUI 執行器：讓使用者透過視覺介面輸入 Search Console property、起訖日與輸出格式，並執行 gsc_keyword_report.py

功能：
- 輸入欄位：property、start-date、end-date
- 選項：mock 模式開關、輸出為 CSV / XLSX（可同時勾選）
- 執行後會在下方顯示執行 log

用法：
  python run_gui.py

注意：若選 XLSX 輸出，需要安裝 `pandas` 與 `openpyxl`（已列在 `requirements.txt`）。

Author: Colinjen (colinjen88@gmail.com)
Version: 1.2
"""
import subprocess
import sys
import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import csv
import tkinter.font as tkfont
import re
import json
from datetime import date, timedelta
from datetime import datetime

# Try to import ttkbootstrap for modern theming. Style will be created
# in the App __init__ (bound to the existing Tk root) to avoid creating
# a second hidden root window.
USE_TTB = True
try:
    import ttkbootstrap as tb
    from ttkbootstrap.constants import *
except Exception:
    USE_TTB = False
    tb = None



SCRIPT = "gsc_keyword_report.py"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("GSC Keyword Reporter - GUI")
        # increase window width & height to show all elements
        self.geometry("900x900")

        frm = ttk.Frame(self, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        # initialize ttkbootstrap Style bound to this root (avoid extra root)
        if USE_TTB and tb is not None:
            try:
                self.tb_style = tb.Style(master=self, theme='superhero')
                self.tb_style.configure('TLabel', font=('Segoe UI', 10))
                self.tb_style.configure('TEntry', font=('Segoe UI', 10))
                self.tb_style.configure('TButton', font=('Segoe UI', 10))
                self.tb_style.configure('Big.TButton', font=('Segoe UI', 11, 'bold'), padding=(16,10))
            except Exception:
                pass
        else:
            self.tb_style = None
        
        # create a consistent ttk style for inputs and buttons (AFTER theme init)
        try:
            style = ttk.Style()
            style.configure('Uniform.TEntry', font=('Segoe UI', 10), padding=(6, 4))
            style.configure('Uniform.TButton', font=('Segoe UI', 10), padding=(8, 4))
            style.configure('Uniform.TLabel', font=('Segoe UI', 10))
            style.configure('Uniform.TCombobox', font=('Segoe UI', 10), padding=(6, 4))
            # Custom styles for preset buttons
            style.configure('Preset.TButton', font=('Segoe UI', 10), padding=(6, 4), background='#efefef', foreground='#1565c0')
            style.map('Preset.TButton', background=[('active', '#e0e0e0')], foreground=[('active', '#1565c0')])
            style.configure('Selected.Preset.TButton', font=('Segoe UI', 10), padding=(6, 4), background='#1565c0', foreground='white')
            style.map('Selected.Preset.TButton', background=[('active', '#0d47a1')], foreground=[('active', 'white')])
            # Wide button style for 開始查詢
            style.configure('Wide.TButton', font=('Segoe UI', 11, 'bold'), padding=(20, 8))
        except Exception:
            style = None
        
        # last used preset label (e.g., '近7天', '上個月')
        self.last_preset = None
        # sort state per column: True = descending, False = ascending
        self.sort_state = {}
        # favorites set (stores keywords)
        self.load_favorites()
        # link id counter for log file links
        self._link_count = 0
        # pinned SA path
        self.pinned_sa_path = self.load_pinned_sa()

        current_row = 0

        # === Row 0: 日期區間 ===
        ttk.Label(frm, text="日期區間（YYYY-MM-DD)：", style='Uniform.TLabel').grid(row=current_row, column=0, sticky=tk.W, padx=(8,8), pady=(8,8))
        date_range_frame = ttk.Frame(frm)
        date_range_frame.grid(row=current_row, column=1, columnspan=3, sticky=tk.W, padx=(4,8), pady=(8,8))

        self.start_var = tk.StringVar(value=(date.today() - timedelta(days=30)).isoformat())
        if USE_TTB:
            self.start_entry = tb.DateEntry(date_range_frame, bootstyle="primary", startdate=date.today() - timedelta(days=30), firstweekday=0, dateformat='%Y-%m-%d')
            self.start_entry.pack(side=tk.LEFT, padx=(0, 4))
        else:
            ttk.Entry(date_range_frame, textvariable=self.start_var, width=20, style='Uniform.TEntry').pack(side=tk.LEFT, padx=(0, 4))

        ttk.Label(date_range_frame, text="～", style='Uniform.TLabel').pack(side=tk.LEFT, padx=(0, 4))
        
        self.end_var = tk.StringVar(value=date.today().isoformat())
        if USE_TTB:
            self.end_entry = tb.DateEntry(date_range_frame, bootstyle="primary", startdate=date.today(), firstweekday=0, dateformat='%Y-%m-%d')
            self.end_entry.pack(side=tk.LEFT, padx=(0, 0))
        else:
            ttk.Entry(date_range_frame, textvariable=self.end_var, width=20, style='Uniform.TEntry').pack(side=tk.LEFT, padx=(0, 0))

        current_row += 1

        # === Row 1: 日期快選按鈕 ===
        preset_frame = ttk.Frame(frm)
        preset_frame.grid(row=current_row, column=1, columnspan=3, sticky=tk.W, padx=(8,8), pady=(0,8))
        
        self.preset_btns = {}
        presets = [
            ("日期區間", None),
            ("近7天", 7),
            ("近30天", 30),
            ("近1季", 90),
            ("近1年", 365),
            ("上個月", -1)
        ]
        
        for idx, (label, val) in enumerate(presets):
            cmd = lambda l=label, v=val: self.on_preset_click(l, v)
            btn = ttk.Button(preset_frame, text=label, command=cmd, style='Preset.TButton')
            btn.grid(row=0, column=idx, padx=(2,2))
            self.preset_btns[label] = btn
            
        self.update_preset_visuals("日期區間")

        current_row += 1

        # === Row 2: 可摺疊設定區塊 ===
        self.settings_expanded = tk.BooleanVar(value=False)  # 預設摺疊
        
        settings_header = ttk.Frame(frm)
        settings_header.grid(row=current_row, column=0, columnspan=4, sticky=tk.W, padx=(8,8), pady=(8,4))
        
        self.toggle_settings_btn = ttk.Button(settings_header, text="▶ 打開設定", command=self.toggle_settings, width=12)
        self.toggle_settings_btn.pack(side=tk.LEFT)
        
        current_row += 1

        # === Row 3: 設定內容 (可摺疊) ===
        self.settings_frame = ttk.LabelFrame(frm, text="設定", padding=8)
        self.settings_frame.grid(row=current_row, column=0, columnspan=4, sticky=tk.EW, padx=(8,8), pady=(0,8))
        
        # Search Console 屬性 (URL)
        ttk.Label(self.settings_frame, text="Search Console 屬性 (URL)：", style='Uniform.TLabel').grid(row=0, column=0, sticky=tk.W, padx=(4,8), pady=(4,4))
        self.property_var = tk.StringVar(value="https://pm.shiny.com.tw/")
        ttk.Entry(self.settings_frame, textvariable=self.property_var, width=50, style='Uniform.TEntry').grid(row=0, column=1, columnspan=2, sticky=tk.W, padx=(4,8), pady=(4,4))
        
        # 關鍵字檔案
        ttk.Label(self.settings_frame, text="關鍵字檔案：", style='Uniform.TLabel').grid(row=1, column=0, sticky=tk.W, padx=(4,8), pady=(4,4))
        self.kws_var = tk.StringVar(value="data/keywords/allKeyWord.csv")
        ttk.Entry(self.settings_frame, textvariable=self.kws_var, width=40, style='Uniform.TEntry').grid(row=1, column=1, sticky=tk.W, padx=(4,8), pady=(4,4))
        ttk.Button(self.settings_frame, text="瀏覽", command=self.browse_kws, style='Uniform.TButton').grid(row=1, column=2, sticky=tk.W, padx=(4,8), pady=(4,4))
        
        # Service account JSON (with 釘選 button)
        ttk.Label(self.settings_frame, text="Service account JSON：", style='Uniform.TLabel').grid(row=2, column=0, sticky=tk.W, padx=(4,8), pady=(4,4))
        self.sa_var = tk.StringVar(value=self.pinned_sa_path if self.pinned_sa_path else "")
        ttk.Entry(self.settings_frame, textvariable=self.sa_var, width=40, style='Uniform.TEntry').grid(row=2, column=1, sticky=tk.W, padx=(4,8), pady=(4,4))
        
        sa_btn_frame = ttk.Frame(self.settings_frame)
        sa_btn_frame.grid(row=2, column=2, sticky=tk.W, padx=(4,8), pady=(4,4))
        ttk.Button(sa_btn_frame, text="瀏覽", command=self.browse_sa, style='Uniform.TButton').pack(side=tk.LEFT, padx=(0,4))
        ttk.Button(sa_btn_frame, text="📌 釘選", command=self.pin_sa, style='Uniform.TButton').pack(side=tk.LEFT)
        
        # 輸出檔名前綴
        ttk.Label(self.settings_frame, text="輸出檔名前綴：", style='Uniform.TLabel').grid(row=3, column=0, sticky=tk.W, padx=(4,8), pady=(4,4))
        self.outbase_var = tk.StringVar(value="gsc_keyword_report")
        ttk.Entry(self.settings_frame, textvariable=self.outbase_var, width=30, style='Uniform.TEntry').grid(row=3, column=1, sticky=tk.W, padx=(4,8), pady=(4,4))
        
        # 輸出格式
        self.format_var = tk.StringVar(value='CSV')

        # Log 區域放在設定區塊內
        ttk.Label(self.settings_frame, text="執行日誌：", style='Uniform.TLabel').grid(row=4, column=0, sticky=tk.NW, padx=(4,8), pady=(8,4))
        self.log = tk.Text(self.settings_frame, height=10, width=70)  # 高度縮小為 1/3
        self.log.grid(row=4, column=1, columnspan=2, sticky=tk.EW, padx=(4,8), pady=(4,4))

        # 預設摺疊設定區塊
        self.settings_frame.grid_remove()

        current_row += 1

        # keep legacy run_btn for compatibility (hidden)
        self.run_btn = ttk.Button(frm, text="執行報表", command=self.on_run)
        self.run_btn.grid_forget()

        current_row += 1

        # === Row 5: 結果狀態 ===
        results_frame = ttk.Frame(frm)
        results_frame.grid(row=current_row, column=0, columnspan=4, sticky=tk.W, padx=(8,8), pady=(8,0))
        self.status_var = tk.StringVar(value='待命')
        ttk.Label(results_frame, text="結果：", style='Uniform.TLabel').pack(side='left')
        if USE_TTB:
            self.status_label = tb.Label(results_frame, text='待命', bootstyle='secondary', padding=(6,2))
        else:
            self.status_label = tk.Label(results_frame, text='待命', bg='#808080', fg='white', padx=8, pady=2)
        self.status_label.pack(side='left', padx=(4,0))
        self.progress = ttk.Progressbar(results_frame, mode='indeterminate', length=100)

        current_row += 1

        # === Row 6: 統計資訊 ===
        self.stats_line_var = tk.StringVar(value='關鍵字數: 0  |  總點擊: 0  |  總曝光: 0  |  平均排名: -')

        current_row += 1

        # === Row 7: 表格區塊 (高度增加60px) ===
        self.table_frame = ttk.Frame(frm)
        self.table_frame.grid(row=current_row, column=0, columnspan=4, sticky=tk.NSEW, padx=(8,8), pady=(8,8))
        frm.rowconfigure(current_row, weight=1)  # 表格區塊

        # create a persistent stats label at the top of the table_frame
        try:
            self.stats_label = ttk.Label(self.table_frame, textvariable=self.stats_line_var, style='Uniform.TLabel')
            self.stats_label.grid(row=0, column=0, columnspan=2, sticky=tk.W, padx=(4,4), pady=(4,8))
        except Exception:
            self.stats_label = None

        self.tree = None
        self.current_rows = []
        self.current_columns = []
        
        # 「加入關鍵字總表」按鈕狀態
        self.add_kw_btn = None
        self.last_search_keyword = None

        current_row += 1

        # === Row 8: 底部按鈕列 ===
        btn_frame = ttk.Frame(frm)
        btn_frame.grid(row=current_row, column=0, columnspan=4, sticky=tk.EW, padx=(8,8), pady=(8,8))
        
        # 格式選擇
        try:
            self.fmt_combo_btn = ttk.Combobox(btn_frame, textvariable=self.format_var, values=['CSV', 'Excel (.xlsx)'], state='readonly', width=14, style='Uniform.TCombobox')
            self.fmt_combo_btn.grid(row=0, column=0, padx=(0,8), pady=(0,0))
        except Exception:
            pass

        if USE_TTB:
            self.save_btn = tb.Button(btn_frame, text="輸出檔案", command=self.export_csv, bootstyle='primary-outline')
        else:
            self.save_btn = ttk.Button(btn_frame, text="輸出檔案", command=self.export_csv)
        self.save_btn.grid(row=0, column=1, padx=(0,8), pady=(0,0))
        
        if USE_TTB:
            self.clear_btn = tb.Button(btn_frame, text="清除表格", command=self.clear_table, bootstyle='secondary')
        else:
            self.clear_btn = ttk.Button(btn_frame, text="清除表格", command=self.clear_table)
        self.clear_btn.grid(row=0, column=2, padx=(0,8), pady=(0,0))
        
        # View Favorites Button
        if USE_TTB:
            self.view_fav_btn = tb.Button(btn_frame, text="查看收藏清單", command=self.view_favorites, bootstyle='info-outline')
        else:
            self.view_fav_btn = ttk.Button(btn_frame, text="查看收藏清單", command=self.view_favorites)
        self.view_fav_btn.grid(row=0, column=3, padx=(0,8), pady=(0,0))
        
        # Export Favorites Button
        if USE_TTB:
            self.export_fav_btn = tb.Button(btn_frame, text="匯出收藏關鍵字", command=self.export_favorites, bootstyle='success-outline')
        else:
            self.export_fav_btn = ttk.Button(btn_frame, text="匯出收藏關鍵字", command=self.export_favorites)
        self.export_fav_btn.grid(row=0, column=4, padx=(0,8), pady=(0,0))

        # autoload toggle
        self.autoload_var = tk.BooleanVar(value=True)
        self.autoload_cb = ttk.Checkbutton(btn_frame, text='自動載入 CSV', variable=self.autoload_var)
        self.autoload_cb.grid(row=0, column=5, padx=(8,8), pady=(0,0))

        # 「開始查詢」按鈕 - 在底部按鈕列
        if USE_TTB:
            self.run_btn_big = tb.Button(btn_frame, text='開始查詢', command=self.on_run, bootstyle='success', width=12)
        else:
            self.run_btn_big = ttk.Button(btn_frame, text='開始查詢', command=self.on_run, style='Wide.TButton', width=12)
        self.run_btn_big.grid(row=0, column=6, padx=(16,0), pady=(0,0))

        current_row += 1

        # === Row 9: Version info ===
        version_frame = ttk.Frame(frm)
        version_frame.grid(row=current_row, column=0, columnspan=4, sticky=tk.E, padx=(8,8), pady=(4,8))
        
        version_label = tk.Label(version_frame, text='KeywordTools V1.5 Product by ', 
                                fg='#808080', font=('Segoe UI', 8))
        version_label.pack(side=tk.LEFT)
        
        author_label = tk.Label(version_frame, text='Colinjen', 
                               fg='#808080', font=('Segoe UI', 8), cursor='hand2')
        author_label.pack(side=tk.LEFT)
        author_label.bind('<Button-1>', lambda e: self.open_email())

        # start file watcher to auto-load CSV created externally
        try:
            self.start_file_watcher()
        except Exception:
            pass
        # clear last_preset if user edits start/end manually
        try:
            self.ignore_trace = False
            def _clear_preset(*args):
                if getattr(self, 'ignore_trace', False):
                    return
                self.last_preset = None
                self.update_preset_visuals("日期區間")
            self.start_var.trace_add('write', _clear_preset)
            self.end_var.trace_add('write', _clear_preset)
        except Exception:
            pass

    def toggle_settings(self):
        """展開/摺疊設定區塊"""
        if self.settings_expanded.get():
            self.settings_frame.grid_remove()
            self.toggle_settings_btn.configure(text="▶ 打開設定")
            self.settings_expanded.set(False)
        else:
            self.settings_frame.grid()
            self.toggle_settings_btn.configure(text="▼ 收起設定")
            self.settings_expanded.set(True)

    def load_pinned_sa(self):
        """載入釘選的 Service Account JSON 路徑"""
        try:
            config_path = 'config/pinned_sa.json'
            if os.path.exists(config_path):
                with open(config_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    return data.get('sa_path', '')
        except Exception:
            pass
        return ''

    def pin_sa(self):
        """釘選當前的 Service Account JSON 路徑"""
        sa_path = self.sa_var.get()
        if not sa_path:
            messagebox.showwarning('無路徑', '請先選擇 Service Account JSON 檔案')
            return
        try:
            config_path = 'config/pinned_sa.json'
            os.makedirs('config', exist_ok=True)
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump({'sa_path': sa_path}, f, ensure_ascii=False, indent=2)
            self.pinned_sa_path = sa_path
            messagebox.showinfo('已釘選', f'已釘選 Service Account 路徑：\n{sa_path}')
        except Exception as e:
            messagebox.showerror('錯誤', f'釘選失敗：{e}')

    def on_preset_click(self, label, days):
        self.update_preset_visuals(label)
        if days is None:
            return
        self.ignore_trace = True
        try:
            if days == -1:
                self.set_preset_last_month()
            else:
                self.set_preset(days)
        finally:
            self.ignore_trace = False

    def update_preset_visuals(self, selected_label):
        self.current_preset_label = selected_label
        for label, btn in self.preset_btns.items():
            if label == selected_label:
                btn.configure(style='Selected.Preset.TButton')
            else:
                btn.configure(style='Preset.TButton')

    def browse_kws(self):
        p = filedialog.askopenfilename(initialdir='.', filetypes=[('CSV files','*.csv'),('All files','*.*')])
        if p:
            self.kws_var.set(p)

    def browse_sa(self):
        p = filedialog.askopenfilename(initialdir='.', filetypes=[('JSON files','*.json'),('All files','*.*')])
        if p:
            self.sa_var.set(p)

    def open_email(self):
        import webbrowser
        webbrowser.open('mailto:colinjen88@gmail.com')

    def set_preset(self, days:int):
        end = date.today()
        start = end - timedelta(days=days-1)
        self.start_var.set(start.isoformat())
        self.end_var.set(end.isoformat())
        if USE_TTB:
            try:
                self.start_entry.entry.delete(0, tk.END)
                self.start_entry.entry.insert(0, start.isoformat())
                self.end_entry.entry.delete(0, tk.END)
                self.end_entry.entry.insert(0, end.isoformat())
            except Exception:
                pass
        # record last_preset for status label (e.g., '近7天')
        if days in (7, 30, 90, 365):
            self.last_preset = f'近{days}天'
        else:
            self.last_preset = None

    def set_preset_last_month(self):
        # set start and end to last calendar month
        today = date.today()
        first_of_this_month = today.replace(day=1)
        last_day_last_month = first_of_this_month - timedelta(days=1)
        start = last_day_last_month.replace(day=1)
        end = last_day_last_month
        self.start_var.set(start.isoformat())
        self.end_var.set(end.isoformat())
        if USE_TTB:
            try:
                self.start_entry.entry.delete(0, tk.END)
                self.start_entry.entry.insert(0, start.isoformat())
                self.end_entry.entry.delete(0, tk.END)
                self.end_entry.entry.insert(0, end.isoformat())
            except Exception:
                pass
        self.last_preset = '上個月'

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
        # resolve path for PyInstaller onefile bundles (sys._MEIPASS) if needed
        try:
            if not os.path.exists(path) and getattr(sys, 'frozen', False):
                base = getattr(sys, '_MEIPASS', None) or os.path.dirname(sys.executable)
                alt = os.path.join(base, os.path.basename(path))
                if os.path.exists(alt):
                    path = alt
        except Exception:
            pass
        # If running from PyInstaller bundle, and default keywords file exists inside the bundle, use that path
        try:
            if getattr(sys, 'frozen', False):
                base = getattr(sys, '_MEIPASS', None) or os.path.dirname(sys.executable)
                kws_name = os.path.basename(self.kws_var.get())
                bundle_kws = os.path.join(base, kws_name)
                if os.path.exists(bundle_kws):
                    self.kws_var.set(bundle_kws)
        except Exception:
            pass
        rows = []
        header = []
        used_encoding = None
        encodings_to_try = ['utf-8-sig', 'utf-8', 'utf-16', 'cp950', 'cp936', 'latin1']
        for enc in encodings_to_try:
            try:
                with open(path, newline='', encoding=enc) as fh:
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
                used_encoding = enc
                break
            except UnicodeDecodeError:
                # try next encoding
                continue
            except Exception as e:
                # for other errors, log and try next
                self.append_log(f'嘗試使用編碼 {enc} 讀取失敗: {e}')
                continue
        if not used_encoding:
            raise ValueError('無法開啟 CSV：不支援的編碼或檔案已損毀')

        # Validate CSV format - prevent UI freeze from too many columns
        MAX_COLUMNS = 50
        if len(header) > MAX_COLUMNS:
            # This likely means the CSV is malformed (e.g., all keywords on one comma-separated line)
            self.append_log(f'⚠ CSV 格式異常：偵測到 {len(header)} 個欄位，超過 {MAX_COLUMNS} 個上限。')
            self.append_log(f'  這可能表示 CSV 檔案格式錯誤（所有關鍵字在同一行以逗號分隔）。')
            self.append_log(f'  請確認 CSV 格式正確，每行一個關鍵字，並包含標題行。')
            raise ValueError(f'CSV 欄位數異常 ({len(header)} 欄)，可能是格式錯誤。請檢查 CSV 檔案。')

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
        idx_prev_pos = idx(['prev_month_position', 'prev_pos'])

        # Desired columns: Mark, Keyword, Trend, Change, Position, Prev Position, Clicks, Impressions, CTR
        display_cols = ['標記', '關鍵字', '趨勢', '變化', '排名', '前月排名', '點擊', '曝光', '點擊率']
        mapped_rows = []
        for r in rows:
            mapped = []
            # mark (checkbox)
            kw = r[idx_keyword] if idx_keyword is not None and idx_keyword < len(r) else ''
            if kw in self.favorites:
                mapped.append('☑')
            else:
                mapped.append('☐')
            # keyword
            mapped.append(kw)
            
            # Prepare numeric values for trend calculation
            try:
                pv = float(str(r[idx_pos]).replace(',',''))
            except: pv = 0.0
            
            try:
                ppv = float(str(r[idx_prev_pos]).replace(',',''))
            except: ppv = 0.0
            
            # Trend arrow
            trend_text = ''
            change_text = ''
            if pv > 0 and ppv > 0:
                diff = ppv - pv # Prev - Curr. Positive means improved (rank got smaller)
                if diff >= 0.5: # Improved
                    trend_text = '▲'
                    change_text = f'+{round(diff, 1)}'
                elif diff <= -0.5: # Worsened
                    trend_text = 'X'
                    change_text = f'{round(diff, 1)}'
            mapped.append(trend_text)
            mapped.append(change_text)

            # position
            try:
                mapped.append(str(round(pv, 1)) if pv > 0 else '')
            except Exception:
                mapped.append(r[idx_pos] if idx_pos is not None and idx_pos < len(r) else '')
            
            # prev position
            if ppv == 0:
                mapped.append('-')
            else:
                mapped.append(str(round(ppv, 1)))

            # clicks
            mapped.append(r[idx_clicks] if idx_clicks is not None and idx_clicks < len(r) else '')
            # impressions
            mapped.append(r[idx_impr] if idx_impr is not None and idx_impr < len(r) else '')
            # ctr (clicks / impressions)
            try:
                c = float(str(r[idx_clicks]).replace(',', '')) if idx_clicks is not None and idx_clicks < len(r) and str(r[idx_clicks]) != '' else 0.0
            except Exception:
                c = 0.0
            try:
                im = float(str(r[idx_impr]).replace(',', '')) if idx_impr is not None and idx_impr < len(r) and str(r[idx_impr]) != '' else 0.0
            except Exception:
                im = 0.0
            if im:
                ctr = f"{round((c / im) * 100, 2)}%"
            else:
                ctr = ''
            mapped.append(ctr)
            mapped_rows.append(mapped)

        self.current_columns = display_cols
        self.current_rows = mapped_rows

        # log detected encoding for debugging
        try:
            self.append_log(f'已偵測 CSV 編碼：{used_encoding}')
        except Exception:
            pass

        # create tree (height shows ~40 rows; 加大一倍以顯示更多結果)
        # place the tree below the stats label (row=1)
        tree = ttk.Treeview(self.table_frame, columns=display_cols, show='headings', height=40)
        vsb = ttk.Scrollbar(self.table_frame, orient='vertical', command=tree.yview)
        hsb = ttk.Scrollbar(self.table_frame, orient='horizontal', command=tree.xview)
        tree.configure(yscroll=vsb.set, xscroll=hsb.set)
        tree.grid(row=2, column=0, sticky='nsew')
        vsb.grid(row=2, column=1, sticky='ns')
        hsb.grid(row=3, column=0, sticky='ew')
        self.table_frame.rowconfigure(2, weight=1)
        self.table_frame.columnconfigure(0, weight=1)

        # style headings (dark background + white text)
        try:
            style = ttk.Style()
            style.configure('Treeview.Heading', background='#2f2f2f', foreground='white', font=('Segoe UI', 10, 'bold'))
        except Exception:
            pass

        for i, c in enumerate(display_cols):
            if i == 1: # Keyword column
                tree.heading(c, text=c, anchor='w')
            elif i == 0: # Mark column
                tree.heading(c, text=c, anchor='center')
            else:
                tree.heading(c, text=c, anchor='e')
            
            # column widths
            if i == 0: # Mark
                tree.column(c, width=40, anchor='center', stretch=False)
            elif i == 1: # Keyword
                tree.column(c, width=80, anchor='w')
            else:
                tree.column(c, width=160, anchor='e')

        # insert rows with alternating background (visual separator)
        try:
            # Configure tags for background colors
            tree.tag_configure('even', background='#fefefe')
            tree.tag_configure('odd', background='#ededed')
            
            # Configure tags for text colors (favorites based on rank)
            tree.tag_configure('fav_top3', foreground='#cf79a6', font=('Segoe UI', 10, 'bold'))
            tree.tag_configure('fav_21_30', foreground='#e06914', font=('Segoe UI', 10, 'bold'))
            tree.tag_configure('fav_gt31', foreground='#eb4e4e', font=('Segoe UI', 10, 'bold'))

            for idx, r in enumerate(mapped_rows):
                row_tags = []
                
                # Background tag
                bg_tag = 'even' if idx % 2 == 0 else 'odd'
                row_tags.append(bg_tag)
                
                # Foreground color logic for favorites
                if r[0] == '☑':
                    try:
                        rank_val = float(r[4])
                        if rank_val <= 3:
                            row_tags.append('fav_top3')
                        elif 21 <= rank_val <= 30:
                            row_tags.append('fav_21_30')
                        elif rank_val >= 31:
                            row_tags.append('fav_gt31')
                    except (ValueError, TypeError):
                        pass
                
                tree.insert('', tk.END, values=r, tags=tuple(row_tags))
        except Exception:
            for r in mapped_rows:
                tree.insert('', tk.END, values=r)

        self.tree = tree

        # update statistics line (single row, separated by |)
        try:
            kw_count = len(mapped_rows)
            total_clicks = 0
            total_impr = 0
            pos_vals = []
            
            # variables for weighted rank
            weighted_sum = 0.0
            total_impr_for_weight = 0.0

            for r in mapped_rows:
                # clicks (col 6), impressions (col 7), position (col 4) - shifted due to Mark, Trend, Change
                c_val = 0.0
                im_val = 0.0
                p_val = 0.0
                
                try:
                    c = str(r[6]).replace(',', '')
                    c_val = float(c) if c != '' else 0.0
                    total_clicks += c_val
                except Exception:
                    pass
                try:
                    im = str(r[7]).replace(',', '')
                    im_val = float(im) if im != '' else 0.0
                    total_impr += im_val
                except Exception:
                    pass
                try:
                    p = float(str(r[4]).replace(',', ''))
                    p_val = p
                    pos_vals.append(p)
                except Exception:
                    pass
                
                # Weighted Rank Calculation: Sum(Position * Impressions) / Sum(Impressions)
                # Only consider if Impressions > 0
                if im_val > 0:
                    weighted_sum += p_val * im_val
                    total_impr_for_weight += im_val

            avg_pos = round(sum(pos_vals) / len(pos_vals), 1) if pos_vals else '-'
            
            # Calculate weighted avg
            if total_impr_for_weight > 0:
                weighted_avg_pos = round(weighted_sum / total_impr_for_weight, 2)
            else:
                weighted_avg_pos = '-'

            stats_text = f'關鍵字數: {kw_count}  |  總點擊: {int(total_clicks)}  |  總曝光: {int(total_impr)}  |  平均排名: {avg_pos}  |  加權平均排名: {weighted_avg_pos}'
            self.stats_line_var.set(stats_text)
        except Exception:
            pass
        # after populating, enable table interactions (sorting, right-click, auto-width)
        try:
            self.setup_table_features()
        except Exception:
            pass

        # stats label is persistent (created in __init__); just ensure value updated and lifted
        try:
            if getattr(self, 'stats_label', None):
                try:
                    # bring stats label to front in case other widgets overlap
                    self.stats_label.lift()
                except Exception:
                    pass
        except Exception:
            pass

    def export_csv(self):
        # unified export: use selected format
        if not self.current_columns:
            messagebox.showinfo('無資料', '目前表格沒有資料可匯出')
            return
        fmt = self.format_var.get() if hasattr(self, 'format_var') else 'CSV'
        # helper for default filename uses class method
        if fmt == 'CSV':
            p = filedialog.asksaveasfilename(defaultextension='.csv', filetypes=[('CSV','*.csv')], initialfile=self.get_export_filename('.csv'))
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
            # Excel export
            p = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel','*.xlsx')], initialfile=self.get_export_filename('.xlsx'))
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
    def export_favorites(self):
        if not self.favorites:
            messagebox.showinfo('無收藏', '目前沒有收藏任何關鍵字')
            return
        
        # Filter rows where keyword (index 1) is in favorites
        fav_rows = [r for r in self.current_rows if len(r) > 1 and r[1] in self.favorites]
        
        if not fav_rows:
             messagebox.showinfo('無收藏', '目前沒有收藏任何關鍵字')
             return

        # Generate filename
        today_str = date.today().strftime('%Y%m%d')
        default_name = f"收藏關鍵字_{today_str}.csv"
        
        p = filedialog.asksaveasfilename(defaultextension='.csv', filetypes=[('CSV','*.csv')], initialfile=default_name)
        if not p:
            return
            
        try:
            with open(p, 'w', newline='', encoding='utf-8-sig') as fh:
                writer = csv.writer(fh)
                writer.writerow(self.current_columns)
                for r in fav_rows:
                    writer.writerow(r)
            messagebox.showinfo('已儲存', f'已匯出 {len(fav_rows)} 筆收藏關鍵字到 {p}')
        except Exception as e:
            messagebox.showerror('錯誤', str(e))

    def load_favorites(self):
        self.favorites = set()
        try:
            if os.path.exists('config/favorites.json'):
                with open('config/favorites.json', 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    if isinstance(data, list):
                        self.favorites = set(data)
        except Exception:
            pass

    def save_favorites(self):
        try:
            with open('config/favorites.json', 'w', encoding='utf-8') as f:
                json.dump(list(self.favorites), f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def view_favorites(self):
        if not self.favorites:
            messagebox.showinfo('無收藏', '目前沒有收藏任何關鍵字')
            return
        
        # Create a new dialog window
        dialog = tk.Toplevel(self)
        dialog.title('收藏清單')
        dialog.geometry('400x500')
        
        # Title label
        title_frame = ttk.Frame(dialog, padding=10)
        title_frame.pack(fill=tk.X)
        ttk.Label(title_frame, text=f'收藏的關鍵字 (共 {len(self.favorites)} 個)', font=('Segoe UI', 12, 'bold')).pack()
        
        # Listbox with scrollbar
        list_frame = ttk.Frame(dialog, padding=10)
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, font=('Segoe UI', 10))
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=listbox.yview)
        
        # Add favorites to listbox (sorted)
        for kw in sorted(self.favorites):
            listbox.insert(tk.END, kw)
        
        # Button frame
        btn_frame = ttk.Frame(dialog, padding=10)
        btn_frame.pack(fill=tk.X)
        
        ttk.Button(btn_frame, text='關閉', command=dialog.destroy).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_frame, text='複製全部', command=lambda: self.copy_to_clipboard('\n'.join(sorted(self.favorites)))).pack(side=tk.RIGHT, padx=5)

    # ----- Table interactions: sorting, auto-width, filter, right-click -----
    def setup_table_features(self):
        # add column-sorting handlers (toggle sort on header click)
        for col in self.current_columns:
            try:
                # numeric is True for position/clicks/impr/ctr except keyword
                numeric = col in ('排名', '前月排名', '點擊', '曝光', '點擊率')
                self.tree.heading(col, text=col, command=lambda c=col, n=numeric: self.sort_by_column(c, n))
            except Exception:
                pass

        # enable right-click menu
        # enable right-click menu
        self.tree.bind('<Button-3>', self.on_tree_right_click)
        # enable left-click for checkbox toggle
        self.tree.bind('<Button-1>', self.on_tree_click)

        # auto adjust column widths
        self.adjust_column_widths()

        # add simple filter UI above table
        try:
            if getattr(self, 'filter_frame', None):
                self.filter_frame.destroy()
            self.filter_frame = ttk.Frame(self.table_frame)
            self.filter_frame.grid(row=1, column=0, sticky='ew', pady=(0,4))
            self.filter_frame.columnconfigure(10, weight=1)  # 讓右側有彈性空間
            
            ttk.Label(self.filter_frame, text='欄位篩選：').grid(row=0, column=0, sticky=tk.W)
            
            # Default to '關鍵字' column instead of '標記' to avoid auto-filtering on load
            default_col = '關鍵字' if '關鍵字' in self.current_columns else (self.current_columns[1] if len(self.current_columns) > 1 else self.current_columns[0] if self.current_columns else '')
            self.filter_col_var = tk.StringVar(value=default_col)
            col_combo = ttk.Combobox(self.filter_frame, textvariable=self.filter_col_var, values=self.current_columns, state='readonly', width=12)
            col_combo.grid(row=0, column=1, padx=4)
            col_combo.bind('<<ComboboxSelected>>', self.on_filter_col_change)
            
            self.filter_op_var = tk.StringVar(value='>')
            self.op_combo = ttk.Combobox(self.filter_frame, textvariable=self.filter_op_var, values=['>', '=', '<'], state='readonly', width=3)
            self.op_combo.grid(row=0, column=2, padx=2)
            
            self.filter_val_var = tk.StringVar()
            self.val_entry = ttk.Entry(self.filter_frame, textvariable=self.filter_val_var, width=24)
            self.val_entry.grid(row=0, column=3, padx=4)
            # 綁定 Enter 鍵進行篩選
            self.val_entry.bind('<Return>', lambda e: self.apply_filter())
            
            ttk.Button(self.filter_frame, text='套用', command=self.apply_filter).grid(row=0, column=4, padx=4)
            ttk.Button(self.filter_frame, text='清除', command=self.clear_filter).grid(row=0, column=5, padx=4)
            
            # 「加入關鍵字總表」按鈕 - 預設不可按
            if USE_TTB:
                self.add_kw_btn = tb.Button(self.filter_frame, text='加入關鍵字總表', command=self.add_to_keyword_list, bootstyle='warning-outline', state='disabled')
            else:
                self.add_kw_btn = ttk.Button(self.filter_frame, text='加入關鍵字總表', command=self.add_to_keyword_list, state='disabled')
            self.add_kw_btn.grid(row=0, column=6, padx=(8,4))
            
            # 空白區填充
            ttk.Frame(self.filter_frame).grid(row=0, column=10, sticky='ew')
            
            # 「開始查詢」按鈕 - 置右對齊，加寬
            if USE_TTB:
                self.query_btn = tb.Button(self.filter_frame, text='開始查詢', command=self.on_run, bootstyle='success', width=14)
            else:
                self.query_btn = ttk.Button(self.filter_frame, text='開始查詢', command=self.on_run, style='Wide.TButton', width=14)
            self.query_btn.grid(row=0, column=11, padx=(4,0), sticky='e')
            
            self.on_filter_col_change()
        except Exception:
            pass

    def sort_by_column(self, col, numeric=False):
        # sort tree items by given column; toggles ascending/descending and update heading indicator
        try:
            children = list(self.tree.get_children(''))
            data = [(self.tree.set(k, col), k) for k in children]
            # try numeric
            try:
                # remove percentage / commas
                def to_num(v):
                    if isinstance(v, str):
                        v2 = v.replace('%', '').replace(',', '')
                        return float(v2) if v2 != '' else 0.0
                    return float(v)
                data = [(to_num(v), k) for v, k in data]
            except Exception:
                pass
            # toggle state
            cur = self.sort_state.get(col, False)
            # current False means ascending next; set reverse accordingly
            rev = not cur
            data.sort(reverse=rev)
            # save toggled state
            self.sort_state[col] = not cur
            for index, (val, k) in enumerate(data):
                self.tree.move(k, '', index)
            # after reorder, restore alternating row colors by reassigning tags
            for i, k in enumerate(self.tree.get_children('')):
                tags = ['even' if i % 2 == 0 else 'odd']
                
                # Restore favorite styling
                if self.tree.set(k, '標記') == '☑':
                    try:
                        rank_val = float(self.tree.set(k, '排名'))
                        if rank_val <= 3:
                            tags.append('fav_top3')
                        elif 21 <= rank_val <= 30:
                            tags.append('fav_21_30')
                        elif rank_val >= 31:
                            tags.append('fav_gt31')
                    except (ValueError, TypeError):
                        pass
                self.tree.item(k, tags=tuple(tags))
            # update heading indicator
            try:
                # remove arrows from all headings
                for heading in self.current_columns:
                    text = heading
                    self.tree.heading(heading, text=text)
                # set indicator for current column
                indicator = '▲' if self.sort_state.get(col, False) else '▼'
                self.tree.heading(col, text=f"{col} {indicator}")
            except Exception:
                pass
        except Exception as e:
            self.append_log('排序失敗: ' + str(e))

    def on_tree_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region == "cell" or region == "tree":
            col = self.tree.identify_column(event.x)
            # col is like '#1', '#2'. Mark is column #1 (index 0)
            # Expand clickable area: allow clicking anywhere in first 60 pixels
            if col == '#1' or (col and event.x < 60):
                row_id = self.tree.identify_row(event.y)
                if row_id:
                    current_val = self.tree.set(row_id, '標記')
                    kw = self.tree.set(row_id, '關鍵字')
                    if current_val == '☐':
                        new_val = '☑'
                        self.favorites.add(kw)
                    else:
                        new_val = '☐'
                        if kw in self.favorites:
                            self.favorites.remove(kw)
                    self.save_favorites()
                    self.tree.set(row_id, '標記', new_val)
                    
                    # Update tags immediately
                    current_tags = list(self.tree.item(row_id, 'tags'))
                    
                    # Remove old favorite tags
                    for tag in ['fav_top3', 'fav_21_30', 'fav_gt31']:
                        if tag in current_tags:
                            current_tags.remove(tag)

                    # Add favorite tags if marked
                    if new_val == '☑':
                        try:
                            rank_val = float(self.tree.set(row_id, '排名'))
                            if rank_val <= 3:
                                current_tags.append('fav_top3')
                            elif 21 <= rank_val <= 30:
                                current_tags.append('fav_21_30')
                            elif rank_val >= 31:
                                current_tags.append('fav_gt31')
                        except (ValueError, TypeError):
                            pass

                    self.tree.item(row_id, tags=tuple(current_tags))
                    
                    # Auto sort removed as per user request
                    # self.sort_favorites()

    def sort_favorites(self):
        # Sort by Mark (descending: ☑ > ☐) then by Keyword (ascending)
        try:
            children = list(self.tree.get_children(''))
            # key: (is_marked (bool), keyword)
            # ☑ is greater than ☐ in unicode? U+2611 vs U+2610. 2611 > 2610. So descending works.
            def sort_key(k):
                mark = self.tree.set(k, '標記')
                kw = self.tree.set(k, '關鍵字')
                return (mark, kw)
            
            # We want marked first (descending mark), but keyword ascending.
            # So we can't just use simple sort.
            # Let's sort by keyword asc first, then stable sort by mark desc.
            children.sort(key=lambda k: self.tree.set(k, '關鍵字'))
            children.sort(key=lambda k: self.tree.set(k, '標記'), reverse=True)
            
            for index, k in enumerate(children):
                self.tree.move(k, '', index)
            
            # restore row colors
            for i, k in enumerate(self.tree.get_children('')):
                tag = 'even' if i % 2 == 0 else 'odd'
                self.tree.item(k, tags=(tag,))
        except Exception:
            pass

    def adjust_column_widths(self, padding=12):
        # measure content width and set column widths
        try:
            f = tkfont.Font()
            for i, col in enumerate(self.current_columns):
                maxw = f.measure(col)
                for r in self.current_rows:
                    text = str(r[i]) if i < len(r) else ''
                    w = f.measure(text)
                    if w > maxw:
                        maxw = w
                # Reduce keyword column width
                if i == 1: # Keyword is now index 1
                    w_out = max(60, int((maxw + padding) / 4)) # Shrink factor increased
                    self.tree.column(col, width=w_out, stretch=True)
                elif i == 0: # Mark column fixed width
                    w_out = 40
                    self.tree.column(col, width=w_out, stretch=False)
                elif i == 2: # Trend column
                    w_out = 70 # Increased for text
                    self.tree.column(col, width=w_out, anchor='e', stretch=False)
                else:
                    # add an extra right padding for numeric columns
                    w_out = maxw + padding + 16
                    self.tree.column(col, width=w_out, stretch=False)
        except Exception:
            pass

    def on_filter_col_change(self, event=None):
        col = self.filter_col_var.get()
        
        if col == '標記':
            try:
                # Hide operator and entry for cleaner UI
                self.op_combo.grid_remove()
                self.val_entry.grid_remove()
                if hasattr(self, 'trend_combo'):
                    self.trend_combo.grid_remove()
                if hasattr(self, 'sign_combo'):
                    self.sign_combo.grid_remove()
                
                self.filter_val_var.set('☑') # Auto-fill for clarity
                # Auto-apply filter immediately
                self.apply_filter()
            except: pass
            return
        
        # Handle Trend column - show dropdown with ▲/X options
        if col == '趨勢':
            try:
                self.op_combo.grid_remove()
                self.val_entry.grid_remove()
                if hasattr(self, 'sign_combo'):
                    self.sign_combo.grid_remove()
                
                # Create trend dropdown if not exists
                if not hasattr(self, 'trend_combo'):
                    self.trend_var = tk.StringVar(value='▲')
                    self.trend_combo = ttk.Combobox(self.filter_frame, textvariable=self.trend_var, 
                                                    values=['▲', 'X'], state='readonly', width=10)
                
                self.trend_combo.grid(row=0, column=3, padx=4)
            except: pass
            return
        
        # Handle Change column - show +/- dropdown before input
        if col == '變化':
            try:
                self.op_combo.grid_remove()
                if hasattr(self, 'trend_combo'):
                    self.trend_combo.grid_remove()
                
                # Create sign dropdown if not exists
                if not hasattr(self, 'sign_combo'):
                    self.sign_var = tk.StringVar(value='+')
                    self.sign_combo = ttk.Combobox(self.filter_frame, textvariable=self.sign_var,
                                                   values=['+', '-'], state='readonly', width=5)
                
                self.sign_combo.grid(row=0, column=2, padx=2)
                self.val_entry.grid(row=0, column=3, padx=4)
            except: pass
            return

        # Restore visibility for other columns
        try:
            if hasattr(self, 'trend_combo'):
                self.trend_combo.grid_remove()
            if hasattr(self, 'sign_combo'):
                self.sign_combo.grid_remove()
            self.op_combo.grid()
            self.val_entry.grid()
        except: pass

        # Show operator only for numeric columns (not Keyword)
        # Numeric columns: 排名, 前月排名, 點擊, 曝光, 點擊率
        if col in ('排名', '前月排名', '點擊', '曝光', '點擊率'):
            try:
                self.op_combo.config(state='readonly')
            except: pass
        else:
            try:
                self.op_combo.config(state='disabled')
            except: pass

    def apply_filter(self):
        col = self.filter_col_var.get()
        op = self.filter_op_var.get()
        val_str = self.filter_val_var.get().strip().lower()
        if not col:
            return
            
        # Special handling for Mark column (auto-filter)
        if col == '標記':
             # Allow empty val_str because we auto-set it, but if user cleared it, assume they want ☑
             val_str = '☑'
        elif col == '趨勢':
            # For Trend column, get value from dropdown
            if hasattr(self, 'trend_var'):
                val_str = self.trend_var.get()
            else:
                return
        elif col == '變化':
            # For Change column, combine sign + value
            if hasattr(self, 'sign_var') and val_str:
                sign = self.sign_var.get()
                val_str = sign + val_str
            elif not val_str:
                return
        elif val_str == '':
            return

        try:
            filtered = []
            idx = self.current_columns.index(col)
            is_numeric = col in ('排名', '前月排名', '點擊', '曝光', '點擊率')
            
            for r in self.current_rows:
                if idx >= len(r): continue
                cell_val = str(r[idx])
                
                if col == '標記':
                    if cell_val == '☑':
                        filtered.append(r)
                elif col == '趨勢' or col == '變化':
                    # Exact match for Trend and Change
                    if val_str in cell_val:
                        filtered.append(r)
                elif col == '關鍵字' or not is_numeric:
                    if val_str in cell_val.lower():
                        filtered.append(r)
                else:
                    # Numeric comparison
                    try:
                        c_num = float(cell_val.replace(',', '').replace('%', ''))
                        v_num = float(val_str)
                        match = False
                        if op == '>': match = c_num > v_num
                        elif op == '=': match = abs(c_num - v_num) < 1e-9
                        elif op == '<': match = c_num < v_num
                        if match:
                            filtered.append(r)
                    except:
                        pass
            
            # clear tree
            for it in self.tree.get_children():
                self.tree.delete(it)
            for idx, r in enumerate(filtered):
                tags = ['even' if idx % 2 == 0 else 'odd']
                
                # Foreground color logic for favorites
                if r[0] == '☑':
                    try:
                        rank_val = float(r[4])
                        if rank_val <= 3:
                            tags.append('fav_top3')
                        elif 21 <= rank_val <= 30:
                            tags.append('fav_21_30')
                        elif rank_val >= 31:
                            tags.append('fav_gt31')
                    except: pass
                
                self.tree.insert('', tk.END, values=r, tags=tuple(tags))
            self.append_log(f'已套用篩選：{col} {op if is_numeric else "包含"} "{val_str}"（{len(filtered)} 筆）')
            
            # 控制「加入關鍵字總表」按鈕狀態
            # 當篩選關鍵字欄位且查無結果時，啟用按鈕
            if col == '關鍵字' and len(filtered) == 0 and val_str:
                self.last_search_keyword = self.filter_val_var.get().strip()
                if self.add_kw_btn:
                    try:
                        self.add_kw_btn.configure(state='normal')
                    except:
                        self.add_kw_btn['state'] = 'normal'
            else:
                self.last_search_keyword = None
                if self.add_kw_btn:
                    try:
                        self.add_kw_btn.configure(state='disabled')
                    except:
                        self.add_kw_btn['state'] = 'disabled'
                        
        except Exception as e:
            self.append_log('篩選失敗: ' + str(e))

    def clear_filter(self):
        try:
            for it in self.tree.get_children():
                self.tree.delete(it)
            for idx, r in enumerate(self.current_rows):
                tags = ['even' if idx % 2 == 0 else 'odd']
                
                # Foreground color logic for favorites
                if r[0] == '☑':
                    try:
                        rank_val = float(r[4])
                        if rank_val <= 3:
                            tags.append('fav_top3')
                        elif 21 <= rank_val <= 30:
                            tags.append('fav_21_30')
                        elif rank_val >= 31:
                            tags.append('fav_gt31')
                    except: pass
                
                self.tree.insert('', tk.END, values=r, tags=tuple(tags))
            self.filter_val_var.set('')
            self.append_log('已清除篩選')
            
            # 清除篩選時也禁用加入按鈕
            self.last_search_keyword = None
            if self.add_kw_btn:
                try:
                    self.add_kw_btn.configure(state='disabled')
                except:
                    self.add_kw_btn['state'] = 'disabled'
                    
        except Exception as e:
            self.append_log('清除篩選失敗: ' + str(e))

    def add_to_keyword_list(self):
        """將搜尋的關鍵字加入關鍵字總表 CSV"""
        if not self.last_search_keyword:
            messagebox.showwarning('無關鍵字', '沒有待加入的關鍵字')
            return
        
        kw_file = self.kws_var.get()
        if not kw_file:
            messagebox.showerror('錯誤', '請先設定關鍵字檔案路徑')
            return
        
        try:
            # 讀取現有關鍵字
            existing_keywords = set()
            if os.path.exists(kw_file):
                with open(kw_file, 'r', encoding='utf-8') as f:
                    reader = csv.reader(f)
                    for row in reader:
                        if row:
                            existing_keywords.add(row[0].strip())
            
            # 檢查是否已存在
            if self.last_search_keyword in existing_keywords:
                messagebox.showinfo('已存在', f'關鍵字「{self.last_search_keyword}」已在總表中')
                return
            
            # 加入關鍵字
            with open(kw_file, 'a', encoding='utf-8', newline='') as f:
                writer = csv.writer(f)
                writer.writerow([self.last_search_keyword])
            
            self.append_log(f'✅ 已將關鍵字「{self.last_search_keyword}」加入 {kw_file}')
            messagebox.showinfo('成功', f'已將關鍵字「{self.last_search_keyword}」加入關鍵字總表')
            
            # 清除狀態
            self.last_search_keyword = None
            if self.add_kw_btn:
                try:
                    self.add_kw_btn.configure(state='disabled')
                except:
                    self.add_kw_btn['state'] = 'disabled'
                    
        except Exception as e:
            self.append_log(f'加入關鍵字失敗: {e}')
            messagebox.showerror('錯誤', f'加入關鍵字失敗：{e}')

    def on_tree_right_click(self, event):
        # show context menu for copy cell / export row
        try:
            iid = self.tree.identify_row(event.y)
            col = self.tree.identify_column(event.x)
            if not iid:
                return
            # translate col '#1' -> index
            col_index = int(col.replace('#','')) - 1
            values = self.tree.item(iid, 'values')
            cell_value = values[col_index] if col_index < len(values) else ''

            menu = tk.Menu(self, tearoff=0)
            menu.add_command(label='複製儲存格', command=lambda v=cell_value: self.copy_to_clipboard(v))
            menu.add_command(label='匯出此列為 CSV', command=lambda v=values: self.export_row(v))
            menu.tk_popup(event.x_root, event.y_root)
        except Exception as e:
            self.append_log('右鍵選單錯誤: ' + str(e))

    def copy_to_clipboard(self, text):
        try:
            self.clipboard_clear()
            self.clipboard_append(str(text))
            self.append_log('已複製到剪貼簿')
        except Exception as e:
            self.append_log('複製失敗: ' + str(e))

    def export_row(self, values):
        try:
            # default filename include base + today + range
            def _row_default():
                base = self.outbase_var.get().strip() or 'gsc_keyword_report'
                now = datetime.now().strftime('%Y%m%d')
                start = (self.start_var.get().strip() if hasattr(self, 'start_var') else '')
                end = (self.end_var.get().strip() if hasattr(self, 'end_var') else '')
                start_clean = start.replace('-', '') if start else ''
                end_clean = end.replace('-', '') if end else ''
                if start_clean and end_clean:
                    return f"{base}_{now}查詢({start_clean}-{end_clean})_row.csv"
                else:
                    return f"{base}_{now}查詢_row.csv"
            p = filedialog.asksaveasfilename(defaultextension='.csv', filetypes=[('CSV','*.csv')], initialfile=_row_default())
            if not p:
                return
            with open(p, 'w', newline='', encoding='utf-8-sig') as fh:
                writer = csv.writer(fh)
                writer.writerow(self.current_columns)
                writer.writerow(values)
            self.append_log(f'已匯出列到 {p}')
        except Exception as e:
            self.append_log('匯出列失敗: ' + str(e))
        else:
            # Excel
            p = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel','*.xlsx')], initialfile=self.get_export_filename('.xlsx'))
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
        # Detect file paths in text and make them clickable links in the log
        # Strategy: break text into tokens and if a token refers to an existing file path, insert it as a clickable tag
        try:
            if not getattr(self, 'log', None) or not self.log.winfo_exists():
                # fallback to stdout
                print(text)
                return
        except Exception:
            try:
                print(text)
            except Exception:
                pass
            return
        tokens = text.split()
        inserted_any = False
        for i, token in enumerate(tokens):
            if os.path.exists(token):
                # prefix before the path
                prefix = ' '.join(tokens[:i])
                suffix = ' '.join(tokens[i+1:])
                if prefix:
                    self.log.insert(tk.END, prefix + ' ')
                # insert the file path as a clickable tag
                tag_name = f'filelink_{self._link_count}'
                self._link_count += 1
                self.log.insert(tk.END, token, tag_name)
                # style the tag
                try:
                    self.log.tag_config(tag_name, foreground='#1565c0', underline=True)
                    # bind click event
                    self.log.tag_bind(tag_name, '<Button-1>', lambda e, p=token: self.open_file(p))
                except Exception:
                    pass
                if suffix:
                    self.log.insert(tk.END, ' ' + suffix)
                self.log.insert(tk.END, '\n')
                inserted_any = True
                break
        if not inserted_any:
            self.log.insert(tk.END, text + "\n")
        self.log.see(tk.END)

    def open_file(self, path: str):
        try:
            # On Windows, os.startfile is appropriate. Fall back to subprocess on other platforms
            if os.name == 'nt':
                os.startfile(path)
            else:
                import subprocess
                subprocess.run(['xdg-open', path], check=False)
        except Exception as e:
            messagebox.showerror('無法開啟檔案', str(e))

    def set_status(self, text: str, color: str):
        # thread-safe status update
        def _update():
            # map basic color names to bootstyle if ttkbootstrap, else hex
            if USE_TTB:
                # tb supports bootstyle names like 'success', 'info', 'danger'
                bs = 'secondary'
                if color == 'green':
                    bs = 'success'
                elif color == 'blue':
                    bs = 'info'
                elif color == 'red':
                    bs = 'danger'
                try:
                    # status_label text should be just the status (e.g., '查詢完成')
                    self.status_label.configure(text=text, bootstyle=bs)
                except Exception:
                    self.status_label.config(text=text)
            else:
                color_map = {
                    'green': '#2e7d32',
                    'blue': '#1565c0',
                    'red': '#c62828'
                }
                bg = color_map.get(color, color if color and color.startswith('#') else '#808080')
                self.status_label.config(text=text, bg=bg)
        try:
            self.after(0, _update)
        except Exception:
            pass

    def _animate_running_man(self, idx=0):
        if getattr(self, '_anim_stop', True):
            return
        frames = ['🏃', '🏃‍♂️']
        frame = frames[idx % len(frames)]
        # moving dots
        dots = '.' * ((idx // 2) % 4)
        text = f"查詢中 {frame} {dots}"
        # use set_status to update label safely
        self.set_status(text, 'green')
        self.after(200, lambda: self._animate_running_man(idx+1))

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

    def format_range_label(self, start: str, end: str) -> str:
        """Return a human readable range description like '近7天' or '2025-10-01~2025-10-31'."""
        try:
            sdt = datetime.fromisoformat(start).date()
            edt = datetime.fromisoformat(end).date()
            # If end is today and start is N-1 days back, show '近N天'
            today = date.today()
            if edt == today:
                delta = (today - sdt).days + 1
                # common presets: 7, 30, 90, 365
                if delta in (7, 30, 90, 365):
                    return f'近{delta}天'
            # otherwise show start~end
            return f'{sdt.isoformat()}~{edt.isoformat()}'
        except Exception:
            # fallback to raw start-end
            return f'{start}~{end}'

    def get_export_filename(self, ext='.csv') -> str:
        base = self.outbase_var.get().strip() if hasattr(self, 'outbase_var') else 'gsc_keyword_report'
        if not base:
            base = 'gsc_keyword_report'
        now = datetime.now().strftime('%Y%m%d')
        start = (self.start_var.get().strip() if hasattr(self, 'start_var') else '')
        end = (self.end_var.get().strip() if hasattr(self, 'end_var') else '')
        start_clean = start.replace('-', '') if start else ''
        end_clean = end.replace('-', '') if end else ''
        if start_clean and end_clean:
            return f"{base}_{now}查詢({start_clean}-{end_clean}){ext}"
        else:
            return f"{base}_{now}查詢{ext}"


    def on_run(self):
        prop = self.property_var.get().strip()
        if USE_TTB:
            try:
                start = self.start_entry.entry.get().strip()
                end = self.end_entry.entry.get().strip()
            except Exception:
                start = self.start_var.get().strip()
                end = self.end_var.get().strip()
        else:
            start = self.start_var.get().strip()
            end = self.end_var.get().strip()
        kws = self.kws_var.get().strip() or 'allKeyWord_normalized.csv'
        base = self.outbase_var.get().strip() or 'gsc_keyword_report'
        # mock removed: always use service-account if provided
        fmt = self.format_var.get() if hasattr(self, 'format_var') else 'CSV'

        if not prop or not start or not end:
            messagebox.showerror('缺少參數', '請提供 property、開始日期與結束日期')
            return
        # Security policy: require Service Account explicitly (do not use env var fallback)
        sa_path = self.sa_var.get().strip() if hasattr(self, 'sa_var') else ''
        if not sa_path:
            # no service account provided -> show error and refuse to run
            messagebox.showerror('缺少 Service Account', '請選擇您的 Google Service Account 憑證檔案。程式需要此檔案才能向 Google 查詢資料。\n\n請點擊「Service account JSON」旁邊的「瀏覽」按鈕來選擇您的 .json 檔案。')
            return
        # if the file is inside repo, warn the user (avoid committing credentials)
        try:
            repo_root = os.path.abspath('.')
            abs_sa_path = os.path.abspath(sa_path)
            if abs_sa_path.startswith(repo_root):
                res = messagebox.askyesno('警告', '你所選的 Service Account JSON 位於專案目錄下（可能會被 Commit）。是否確定要使用它？')
                if not res:
                    return
        except Exception:
            pass

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
        # set status to querying
        try:
            self._anim_stop = False
            self._animate_running_man(0)
            # self.set_status('查詢中', 'green')
            # try:
            #     self.progress.pack(side='left', padx=(8,0))
            #     self.progress.start(10)
            # except Exception:
            #     pass
        except Exception:
            pass
        self.log.delete('1.0', tk.END)

        def worker():
            try:
                outputs = []
                sa_path = self.sa_var.get().strip() if hasattr(self, 'sa_var') else ''
                out_ext = '.csv' if fmt == 'CSV' else '.xlsx'
                out = self.get_export_filename(out_ext)
                cli_args = ['--property', prop, '--keywords', kws, '--start-date', start, '--end-date', end, '--output', out]
                if sa_path:
                    cli_args.extend(['--service-account', sa_path])

                # log 查詢參數
                self.append_log(f'查詢參數: property={prop}, keywords={kws}, start={start}, end={end}, output={out}, service-account={sa_path}')
                # 檢查關鍵檔案是否存在
                if not os.path.exists(kws):
                    self.append_log(f'關鍵字檔案不存在: {kws}')
                    self.set_status('錯誤', 'red')
                    return
                if sa_path and not os.path.exists(sa_path):
                    self.append_log(f'Service Account 檔案不存在: {sa_path}')
                    self.set_status('錯誤', 'red')
                    return

                script_exit_code = 1  # 預設為失敗

                try:
                    import importlib, io, traceback
                    module = importlib.import_module('gsc_keyword_report')
                    imported_cli = True
                except Exception as ex:
                    module = None
                    imported_cli = False
                    err_tb = traceback.format_exc()
                    self.append_log('無法 import gsc_keyword_report，將 fallback 到 subprocess；錯誤詳情:\n' + err_tb)
                
                if imported_cli:
                    try:
                        old_argv = sys.argv
                        sys.argv = [old_argv[0]] + cli_args
                        buf_out = io.StringIO()
                        buf_err = io.StringIO()
                        old_stdout, old_stderr = sys.stdout, sys.stderr
                        try:
                            sys.stdout, sys.stderr = buf_out, buf_err
                            try:
                                module.main()
                                script_exit_code = 0  # 執行成功
                            except SystemExit as e:
                                code = getattr(e, "code", 1)
                                self.append_log(f'gsc_keyword_report exited with code: {code}')
                                script_exit_code = code if code is not None else 1
                        finally:
                            sys.stdout, sys.stderr = old_stdout, old_stderr
                            sys.argv = old_argv
                        out_text = buf_out.getvalue()
                        err_text = buf_err.getvalue()
                        if out_text:
                            self.append_log(out_text)
                        if err_text:
                            self.append_log(err_text)
                        outputs.append(out)
                    except Exception as e:
                        self.append_log('無法以模組方式執行 CLI: ' + str(e))
                        script_exit_code = 1 # 執行失敗
                
                if not imported_cli:
                    interpreter = sys.executable
                    script_path = SCRIPT
                    if not os.path.exists(script_path):
                        candidate = os.path.join(getattr(sys, '_MEIPASS', os.path.dirname(sys.executable)), SCRIPT) if getattr(sys, 'frozen', False) else None
                        if candidate and os.path.exists(candidate):
                            script_path = candidate
                    cmd = [interpreter, script_path] + cli_args
                    kwargs = {'capture_output': True, 'text': True, 'encoding': 'utf-8'}
                    if os.name == 'nt':
                        kwargs['creationflags'] = subprocess.CREATE_NO_WINDOW
                    self.append_log('執行: ' + ' '.join(cmd))
                    proc = subprocess.run(cmd, **kwargs)
                    self.append_log(proc.stdout)
                    if proc.stderr:
                        self.append_log(proc.stderr)
                    outputs.append(out)
                    script_exit_code = proc.returncode

                if script_exit_code == 0:
                    any_success = False
                    for f in outputs:
                        if os.path.exists(f):
                            self.append_log(f'Generated: {f}')
                            any_success = True
                            if f.lower().endswith('.csv'):
                                try:
                                    # Use after() to update UI on main thread (Tkinter is not thread-safe)
                                    self.after(0, lambda path=f: self.load_csv_into_table(path))
                                except Exception as e:
                                    self.append_log('Failed to load CSV into table: ' + str(e))
                        else:
                            self.append_log(f'Failed to generate: {f}')
                    
                    if any_success:
                        try:
                            if getattr(self, 'last_preset', None):
                                desc = self.last_preset
                            else:
                                desc = self.format_range_label(start, end)
                        except Exception:
                            desc = ''
                        status_text = f'查詢完成_{desc}' if desc else '查詢完成'
                        self.set_status(status_text, 'blue')
                    else:
                        self.set_status('錯誤', 'red')
                else:
                    self.set_status('錯誤', 'red')

            except Exception as e:
                import traceback
                self.append_log('Error: ' + str(e))
                self.append_log(traceback.format_exc())
                self.set_status('錯誤', 'red')
            finally:
                try:
                    self._anim_stop = True
                    try:
                        self.progress.stop()
                        self.progress.pack_forget()
                    except: pass
                except Exception:
                    pass
                try:
                    self.run_btn_big.config(state=tk.NORMAL)
                except Exception:
                    pass
                try:
                    self.run_btn.config(state=tk.NORMAL)
                except Exception:
                    pass

        threading.Thread(target=worker, daemon=True).start()



if __name__ == '__main__':
    app = App()
    app.mainloop()

