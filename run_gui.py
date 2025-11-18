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
import tkinter.font as tkfont
from datetime import date, timedelta
from datetime import datetime

# Try to import ttkbootstrap for modern theming. Style will be created
# in the App __init__ (bound to the existing Tk root) to avoid creating
# a second hidden root window.
USE_TTB = False
try:
    import ttkbootstrap as tb
    USE_TTB = True
except Exception:
    tb = None



SCRIPT = "gsc_keyword_report.py"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("GSC Keyword Reporter - GUI")
        # increase window width & height slightly to show all elements
        self.geometry("780x960")

        frm = ttk.Frame(self, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        # layout constants
        ELEMENT_GAP = 16  # pixels between elements
        LINE_HEIGHT = 16   # fixed row height (改為 16px)

        # create a consistent ttk style for inputs and buttons
        try:
            style = ttk.Style()
            style.configure('Uniform.TEntry', font=('Segoe UI', 10), padding=(6, 4))
            style.configure('Uniform.TButton', font=('Segoe UI', 10), padding=(8, 4))
            style.configure('Uniform.TLabel', font=('Segoe UI', 10))
            style.configure('Uniform.TCombobox', font=('Segoe UI', 10), padding=(6, 4))
        except Exception:
            style = None

        # enforce fixed row height for rows we will use (0..13)
        for r in range(0, 14):
            try:
                frm.rowconfigure(r, minsize=LINE_HEIGHT)
            except Exception:
                pass
        # make column 2 (between start and end inputs) have a min width of ~40px
        try:
            frm.columnconfigure(2, minsize=40)
        except Exception:
            pass

        # initialize ttkbootstrap Style bound to this root (avoid extra root)
        try:
            if USE_TTB and tb is not None:
                self.tb_style = tb.Style(master=self, theme='superhero')
                # tweak some defaults
                try:
                    self.tb_style.configure('TLabel', font=('Segoe UI', 10))
                    self.tb_style.configure('TEntry', font=('Segoe UI', 10))
                    self.tb_style.configure('TButton', font=('Segoe UI', 10))
                    self.tb_style.configure('Big.TButton', font=('Segoe UI', 11, 'bold'), padding=(16,10))
                except Exception:
                    pass
        except Exception:
            self.tb_style = None
            # last used preset label (e.g., '近7天', '上個月')
            self.last_preset = None
            # sort state per column: True = descending, False = ascending
            self.sort_state = {}

        ttk.Label(frm, text="Search Console 屬性 (URL)：", style='Uniform.TLabel').grid(row=0, column=0, sticky=tk.W, padx=(8,8), pady=(8,8))
        self.property_var = tk.StringVar(value="https://pm.shiny.com.tw/")
        ttk.Entry(frm, textvariable=self.property_var, width=60, style='Uniform.TEntry').grid(row=0, column=1, columnspan=3, sticky=tk.W, padx=(8,8), pady=(8,8))

        ttk.Label(frm, text="起始日期（YYYY-MM-DD)：", style='Uniform.TLabel').grid(row=1, column=0, sticky=tk.W, padx=(8,8), pady=(8,8))
        self.start_var = tk.StringVar(value=(date.today() - timedelta(days=30)).isoformat())
        ttk.Entry(frm, textvariable=self.start_var, width=20, style='Uniform.TEntry').grid(row=1, column=1, sticky=tk.W, padx=(8,8), pady=(8,8))

        ttk.Label(frm, text="結束日期（YYYY-MM-DD)：", style='Uniform.TLabel').grid(row=1, column=2, sticky=tk.W, padx=(8,8), pady=(8,8))
        self.end_var = tk.StringVar(value=date.today().isoformat())
        ttk.Entry(frm, textvariable=self.end_var, width=20, style='Uniform.TEntry').grid(row=1, column=3, sticky=tk.W, padx=(8,8), pady=(8,8))

        # preset ranges
        # place preset buttons aligned to the start-date entry's left side
        preset_frame = ttk.Frame(frm)
        preset_frame.grid(row=2, column=1, columnspan=3, sticky=tk.W, padx=(8,8), pady=(0,0))
        # remove the "快速區間：" label; align buttons under start-date
        ttk.Button(preset_frame, text="近7天", command=lambda: self.set_preset(7), style='Uniform.TButton').grid(row=0, column=0, padx=(4,4))
        ttk.Button(preset_frame, text="近30天", command=lambda: self.set_preset(30), style='Uniform.TButton').grid(row=0, column=1, padx=(4,4))
        ttk.Button(preset_frame, text="近1季", command=lambda: self.set_preset(90), style='Uniform.TButton').grid(row=0, column=2, padx=(4,4))
        ttk.Button(preset_frame, text="近1年", command=lambda: self.set_preset(365), style='Uniform.TButton').grid(row=0, column=3, padx=(4,4))
        # add last month preset
        ttk.Button(preset_frame, text="上個月", command=self.set_preset_last_month, style='Uniform.TButton').grid(row=0, column=4, padx=(4,4))

        ttk.Label(frm, text="關鍵字檔案：", style='Uniform.TLabel').grid(row=3, column=0, sticky=tk.W, padx=(8,8), pady=(8,8))
        self.kws_var = tk.StringVar(value="allKeyWord_normalized.csv")
        ttk.Entry(frm, textvariable=self.kws_var, width=40, style='Uniform.TEntry').grid(row=3, column=1, sticky=tk.W, padx=(8,8), pady=(8,8))
        ttk.Button(frm, text="Browse", command=self.browse_kws, style='Uniform.TButton').grid(row=3, column=2, sticky=tk.W, padx=(8,8), pady=(8,8))

        ttk.Label(frm, text="Service account JSON（選填）：", style='Uniform.TLabel').grid(row=4, column=0, sticky=tk.W, padx=(8,8), pady=(8,8))
        self.sa_var = tk.StringVar(value="")
        ttk.Entry(frm, textvariable=self.sa_var, width=40, style='Uniform.TEntry').grid(row=4, column=1, sticky=tk.W, padx=(8,8), pady=(8,8))
        ttk.Button(frm, text="瀏覽", command=self.browse_sa, style='Uniform.TButton').grid(row=4, column=2, sticky=tk.W, padx=(8,8), pady=(8,8))

        ttk.Label(frm, text="輸出檔案基底名稱：", style='Uniform.TLabel').grid(row=5, column=0, sticky=tk.W, padx=(8,8), pady=(8,8))
        self.outbase_var = tk.StringVar(value="gsc_keyword_report")
        ttk.Entry(frm, textvariable=self.outbase_var, width=30, style='Uniform.TEntry').grid(row=5, column=1, sticky=tk.W, padx=(8,8), pady=(8,8))

        # 輸出格式已移至下方按鈕列，預設值保留
        self.format_var = tk.StringVar(value='CSV')

        # keep legacy run_btn for compatibility (hidden)
        self.run_btn = ttk.Button(frm, text="執行報表", command=self.on_run)
        # hide original small run_btn (we use larger one in button frame)
        self.run_btn.grid_forget()

        self.log = tk.Text(frm, height=18)
        self.log.grid(row=8, column=0, columnspan=4, padx=(8,8), pady=(8,8), sticky=tk.NSEW)
        frm.rowconfigure(8, weight=1)
        frm.columnconfigure(3, weight=1)

        # results frame
        # create a results frame to hold status and Results label on one line
        results_frame = ttk.Frame(frm)
        results_frame.grid(row=9, column=0, columnspan=4, sticky=tk.W, padx=(8,8), pady=(8,0))
        self.status_var = tk.StringVar(value='待命')
        # Results label first
        try:
            ttk.Label(results_frame, text="結果：", style='Uniform.TLabel').pack(side='left')
        except Exception:
            ttk.Label(results_frame, text="結果：", style='Uniform.TLabel').grid(row=0, column=0, sticky=tk.W)
        # status text after results (no '狀態：' prefix)
        if USE_TTB:
            self.status_label = tb.Label(results_frame, text='待命', bootstyle='secondary', padding=(6,2))
        else:
            self.status_label = tk.Label(results_frame, text='待命', bg='#808080', fg='white', padx=8, pady=2)
        try:
            # keep Results label and status close together
            self.status_label.pack(side='left', padx=(4,0))
        except Exception:
            self.status_label.grid(row=0, column=1, sticky=tk.W, padx=(8,0))

        # statistics placeholder (will be placed inside the table_frame at its top)
        self.stats_line_var = tk.StringVar(value='關鍵字數: 0  |  總點擊: 0  |  總曝光: 0  |  平均排名: -')

        # results table frame (table below results and stats)
        self.table_frame = ttk.Frame(frm)
        self.table_frame.grid(row=11, column=0, columnspan=4, sticky=tk.NSEW, padx=(8,8), pady=(8,8))
        frm.rowconfigure(10, weight=1)
        frm.rowconfigure(11, weight=1)

        # create a persistent stats label at the top of the table_frame
        try:
            self.stats_label = ttk.Label(self.table_frame, textvariable=self.stats_line_var, style='Uniform.TLabel')
            self.stats_label.grid(row=0, column=0, columnspan=2, sticky=tk.W, padx=(4,4), pady=(4,8))
        except Exception:
            self.stats_label = None

        self.tree = None
        self.current_rows = []
        self.current_columns = []
        btn_frame = ttk.Frame(frm)
        btn_frame.grid(row=12, column=0, columnspan=4, sticky=tk.W, padx=(8,8), pady=(8,8))
        # enlarge Run button style
        try:
            style = ttk.Style()
            style.configure('Big.TButton', font=('Segoe UI', 10, 'bold'), padding=(12,8))
        except Exception:
            pass

        # if ttkbootstrap is present, use its Button for nicer style
        # format combobox placed to the left of the save button for clarity
        try:
            self.fmt_combo_btn = ttk.Combobox(btn_frame, textvariable=self.format_var, values=['CSV', 'Excel (.xlsx)'], state='readonly', width=18, style='Uniform.TCombobox')
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
        # autoload toggle
        self.autoload_var = tk.BooleanVar(value=True)
        # clearer description for auto-load behavior
        self.autoload_cb = ttk.Checkbutton(btn_frame, text='自動載入 CSV（偵測目錄中新產生的 CSV 並自動載入）', variable=self.autoload_var)
        self.autoload_cb.grid(row=0, column=4, padx=(8,8), pady=(0,0))
        # Run button bigger and styled
        if USE_TTB:
            self.run_btn_big = tb.Button(btn_frame, text="執行報表", command=self.on_run, bootstyle='success')
        else:
            self.run_btn_big = ttk.Button(btn_frame, text="執行報表", command=self.on_run, style='Big.TButton')
        self.run_btn_big.grid(row=0, column=3, padx=(12,8), pady=(0,0))

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

        # Desired columns: Keyword, Position, Clicks, Impressions, CTR
        display_cols = ['關鍵字', '排名', '點擊', '曝光', '點擊率']
        mapped_rows = []
        for r in rows:
            mapped = []
            # keyword
            mapped.append(r[idx_keyword] if idx_keyword is not None and idx_keyword < len(r) else '')
            # position
            mapped.append(r[idx_pos] if idx_pos is not None and idx_pos < len(r) else '')
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
        tree.grid(row=1, column=0, sticky='nsew')
        vsb.grid(row=1, column=1, sticky='ns')
        hsb.grid(row=2, column=0, sticky='ew')
        self.table_frame.rowconfigure(1, weight=1)
        self.table_frame.columnconfigure(0, weight=1)

        # style headings (dark background + white text)
        try:
            style = ttk.Style()
            style.configure('Treeview.Heading', background='#2f2f2f', foreground='white', font=('Segoe UI', 10, 'bold'))
        except Exception:
            pass

        for i, c in enumerate(display_cols):
            if i == 0:
                tree.heading(c, text=c, anchor='w')
            else:
                tree.heading(c, text=c, anchor='e')
            # make keyword column half width and align others to right
            if i == 0:
                tree.column(c, width=80, anchor='w')
            else:
                tree.column(c, width=160, anchor='e')

        # insert rows with alternating background (visual separator)
        try:
            for idx, r in enumerate(mapped_rows):
                tag = 'even' if idx % 2 == 0 else 'odd'
                # ensure numeric columns are right aligned; set tag for entire row
                tree.insert('', tk.END, values=r, tags=(tag,))
            # configure tag backgrounds
            try:
                tree.tag_configure('even', background='#ffffff')
                tree.tag_configure('odd', background='#f6f6f6')
            except Exception:
                pass
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
            for r in mapped_rows:
                # clicks (col 1), impressions (col 2), position (col 3)
                try:
                    c = str(r[1]).replace(',', '')
                    total_clicks += float(c) if c != '' else 0.0
                except Exception:
                    pass
                try:
                    im = str(r[2]).replace(',', '')
                    total_impr += float(im) if im != '' else 0.0
                except Exception:
                    pass
                try:
                    p = float(str(r[3]).replace(',', ''))
                    pos_vals.append(p)
                except Exception:
                    pass
            avg_pos = round(sum(pos_vals) / len(pos_vals), 2) if pos_vals else '-'
            stats_text = f'關鍵字數: {kw_count}  |  總點擊: {int(total_clicks)}  |  總曝光: {int(total_impr)}  |  平均排名: {avg_pos}'
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

    # ----- Table interactions: sorting, auto-width, filter, right-click -----
    def setup_table_features(self):
        # add column-sorting handlers (toggle sort on header click)
        for col in self.current_columns:
            try:
                # numeric is True for position/clicks/impr/ctr except keyword
                numeric = col in ('排名', '點擊', '曝光', '點擊率')
                self.tree.heading(col, text=col, command=lambda c=col, n=numeric: self.sort_by_column(c, n))
            except Exception:
                pass

        # enable right-click menu
        self.tree.bind('<Button-3>', self.on_tree_right_click)

        # auto adjust column widths
        self.adjust_column_widths()

        # add simple filter UI above table
        try:
            if getattr(self, 'filter_frame', None):
                self.filter_frame.destroy()
            self.filter_frame = ttk.Frame(self.table_frame)
            self.filter_frame.grid(row= -1, column=0, sticky='ew', pady=(0,4))
            ttk.Label(self.filter_frame, text='欄位篩選：').grid(row=0, column=0, sticky=tk.W)
            self.filter_col_var = tk.StringVar(value=self.current_columns[0] if self.current_columns else '')
            col_combo = ttk.Combobox(self.filter_frame, textvariable=self.filter_col_var, values=self.current_columns, state='readonly', width=12)
            col_combo.grid(row=0, column=1, padx=4)
            self.filter_val_var = tk.StringVar()
            ttk.Entry(self.filter_frame, textvariable=self.filter_val_var, width=24).grid(row=0, column=2, padx=4)
            ttk.Button(self.filter_frame, text='套用', command=self.apply_filter).grid(row=0, column=3, padx=4)
            ttk.Button(self.filter_frame, text='清除', command=self.clear_filter).grid(row=0, column=4, padx=4)
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
                tag = 'even' if i % 2 == 0 else 'odd'
                self.tree.item(k, tags=(tag,))
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
                # Reduce keyword column width to roughly half
                if i == 0:
                    w_out = max(60, int((maxw + padding) / 2))
                else:
                    # add an extra right padding for numeric columns
                    w_out = maxw + padding + 16
                self.tree.column(col, width=w_out)
        except Exception:
            pass

    def apply_filter(self):
        col = self.filter_col_var.get()
        val = self.filter_val_var.get().strip().lower()
        if not col or val == '':
            return
        # filter current_rows and reload tree
        try:
            filtered = []
            idx = self.current_columns.index(col)
            for r in self.current_rows:
                if idx < len(r) and val in str(r[idx]).lower():
                    filtered.append(r)
            # clear tree
            for it in self.tree.get_children():
                self.tree.delete(it)
            for idx, r in enumerate(filtered):
                tag = 'even' if idx % 2 == 0 else 'odd'
                self.tree.insert('', tk.END, values=r, tags=(tag,))
            self.append_log(f'已套用篩選：{col} 包含 "{val}"（{len(filtered)} 筆）')
        except Exception as e:
            self.append_log('篩選失敗: ' + str(e))

    def clear_filter(self):
        try:
            for it in self.tree.get_children():
                self.tree.delete(it)
            for idx, r in enumerate(self.current_rows):
                tag = 'even' if idx % 2 == 0 else 'odd'
                self.tree.insert('', tk.END, values=r, tags=(tag,))
            self.filter_val_var.set('')
            self.append_log('已清除篩選')
        except Exception as e:
            self.append_log('清除篩選失敗: ' + str(e))

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
            p = filedialog.asksaveasfilename(defaultextension='.csv', filetypes=[('CSV','*.csv')], initialfile='row_export.csv')
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


    def on_run(self):
        prop = self.property_var.get().strip()
        start = self.start_var.get().strip()
        end = self.end_var.get().strip()
        kws = self.kws_var.get().strip() or 'allKeyWord_normalized.csv'
        base = self.outbase_var.get().strip() or 'gsc_keyword_report'
        # mock removed: always use service-account if provided
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
                    # include range/preset description in status
                    try:
                        # prefer last_preset if user clicked preset
                        if getattr(self, 'last_preset', None):
                            desc = self.last_preset
                        else:
                            desc = self.format_range_label(start, end)
                    except Exception:
                        desc = ''
                    status_text = f'查詢完成_{desc}' if desc else '查詢完成'
                    self.set_status(status_text, 'blue')
                except Exception:
                    pass

        threading.Thread(target=worker, daemon=True).start()


if __name__ == '__main__':
    app = App()
    app.mainloop()

