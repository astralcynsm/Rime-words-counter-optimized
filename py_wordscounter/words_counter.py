# -*- coding: utf-8 -*-
"""
Rime-words-counter 优化版
原作者：hyuan42
依赖：
pip install portalocker pillow watchdog schedule pywin32 ttkbootstrap
(已移除 pystray)
"""
import os
import io
import sys
import json
import time
import math
import queue
import random
import threading
import traceback
import winreg 
from datetime import datetime, date, timedelta
from collections import defaultdict
from tkinter import font as tkFont

import portalocker
import schedule
from PIL import Image, ImageDraw, ImageFont
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# Tk & UI 美化
import tkinter as tk
from tkinter import messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *

# ================== 路径与配置 (按需修改) ==================
CSV_FILE = r'C:\Users\你的用户名\AppData\Roaming\Rime\py_wordscounter\words_input.csv'
JSON_FILE = r'C:\Users\你的用户名\AppData\Roaming\Rime\py_wordscounter\words_count_history.json'
DEVICE_KEY = 'win_main'

# ================== UI 个性化配置 ==================
FLOAT_FONT_NAME = "Ubuntu Nerd Font" 
FLOAT_FONT_SIZE = 12
FLOAT_FONT_SPEED_SIZE = 10
LIGHT_THEME = "litera"
DARK_THEME = "superhero"

# ================== 常量与工具 ==================
CSV_HEADER = '"timestamp","chinese_count"\n'
# <editor-fold desc="辅助函数和类">
def ensure_dirs(): os.makedirs(os.path.dirname(CSV_FILE), exist_ok=True); os.makedirs(os.path.dirname(JSON_FILE), exist_ok=True)
def safe_file_access(file_path, mode, lock_type="exclusive", retries=6, delay=0.12):
    for attempt in range(retries):
        f = None
        try:
            f = open(file_path, mode, encoding='utf-8');
            if lock_type: portalocker.lock(f, (portalocker.LOCK_SH if 'r' in mode else portalocker.LOCK_EX) | portalocker.LOCK_NB)
            return f
        except Exception:
            if f:
                try: f.close()
                except Exception: pass
            time.sleep(delay * (1.8 ** attempt))
    f = open(file_path, mode, encoding='utf-8')
    if lock_type: portalocker.lock(f, portalocker.LOCK_SH if 'r' in mode else portalocker.LOCK_EX)
    return f
def init_csv_if_missing():
    if not os.path.exists(CSV_FILE):
        with safe_file_access(CSV_FILE, "w", "exclusive") as f: f.write(CSV_HEADER)
def init_json_if_missing(settings_manager):
    if not os.path.exists(JSON_FILE):
        payload = {"per_day": {}, "last_offsets": {}, "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "settings": settings_manager.get_default_settings()}
        with safe_file_access(JSON_FILE, "w", "exclusive") as f: json.dump(payload, f, ensure_ascii=False, indent=2)
def parse_csv_lines(raw_bytes):
    if not raw_bytes: return []
    text = raw_bytes.decode("utf-8", errors="ignore"); lines = [ln.strip() for ln in text.splitlines() if ln.strip()]; res = []
    for ln in lines:
        if ln.lower().startswith('"timestamp","chinese_count"'): continue
        try:
            if ln.count('","') == 1 and ln.startswith('"') and ln.endswith('"'): ts, cnt = ln[1:-1].split('","', 1); res.append((ts, int(cnt)))
        except Exception: pass
    return res
def aggregate_counts(per_day):
    today_str = date.today().strftime("%Y-%m-%d"); year_str, month_str = today_str[:4], today_str[:7]
    return {"today": per_day.get(today_str, 0), "month": sum(v for k, v in per_day.items() if k.startswith(month_str)), "year": sum(v for k, v in per_day.items() if k.startswith(year_str)), "total": sum(per_day.values())}
class Debouncer:
    def __init__(self, delay_sec, fn): self.delay, self.fn = delay_sec, fn; self._timer, self._lock = None, threading.Lock()
    def trigger(self, *args, **kwargs):
        with self._lock:
            if self._timer: self._timer.cancel()
            self._timer = threading.Timer(self.delay, self._run, args=args, kwargs=kwargs); self._timer.daemon = True; self._timer.start()
    def _run(self, *args, **kwargs):
        try: self.fn(*args, **kwargs)
        except Exception: traceback.print_exc()
class CSVHandler(FileSystemEventHandler):
    def __init__(self, on_change): super().__init__(); self.debouncer = Debouncer(0.05, on_change)
    def on_modified(self, event):
        if os.path.abspath(event.src_path) == os.path.abspath(CSV_FILE): self.debouncer.trigger()
    def on_created(self, event):
        if os.path.abspath(event.src_path) == os.path.abspath(CSV_FILE): self.debouncer.trigger()
# </editor-fold>
# <editor-fold desc="SettingsManager">
class SettingsManager:
    def __init__(self, file_path, data_processor_writer):
        self.file_path = file_path; self._trigger_save = data_processor_writer
        self.settings = self.get_default_settings(); self.load_settings()
    def get_default_settings(self): return {"theme": "auto", "float_show_speed": False}
    def load_settings(self):
        try:
            with safe_file_access(self.file_path, "r", "shared") as f: self.settings.update(json.load(f).get("settings", {}))
        except (FileNotFoundError, json.JSONDecodeError): pass
        return self.settings
    def save_setting(self, key, value): self.settings[key] = value; self._trigger_save()
    def get(self, key): return self.settings.get(key)
# </editor-fold>
# <editor-fold desc="DataProcessor">
class DataProcessor:
    def __init__(self, update_callback=None):
        self.update_callback = update_callback; self.data_lock = threading.Lock()
        self.per_day = {}; self.last_offsets = {}
        self.json_writer = Debouncer(2.0, self._write_json_to_disk); self.settings_manager = None
    def set_settings_manager(self, manager): self.settings_manager = manager
    def _load_initial_data(self):
        try:
            with safe_file_access(JSON_FILE, "r", "shared") as f:
                payload = json.load(f); self.per_day = payload.get("per_day", {}); self.last_offsets = payload.get("last_offsets", {})
        except (FileNotFoundError, json.JSONDecodeError): self.per_day = {}; self.last_offsets = {}
        self.process_incremental()
    def _write_json_to_disk(self):
        with self.data_lock:
            payload = {"per_day": self.per_day, "last_offsets": self.last_offsets, "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "settings": self.settings_manager.settings if self.settings_manager else {}}
        try:
            with safe_file_access(JSON_FILE, "w", "exclusive") as f: json.dump(payload, f, ensure_ascii=False, indent=2)
        except Exception as e: print(f"Error writing to JSON: {e}")
    def process_incremental(self):
        new_data_found = False
        with self.data_lock:
            offset = int(self.last_offsets.get(DEVICE_KEY, 0))
            if not os.path.exists(CSV_FILE): init_csv_if_missing(); return
            try:
                with open(CSV_FILE, "rb") as f:
                    file_size = f.seek(0, io.SEEK_END)
                    if offset > file_size or (offset == 0 and file_size > 0 and file_size <= len(CSV_HEADER.encode())): offset = 0
                    f.seek(offset); chunk = f.read(); new_offset = f.tell()
            except Exception: return
            if chunk:
                new_data_found = True; rows = parse_csv_lines(chunk)
                for ts, cnt in rows: self.per_day[ts[:10]] = self.per_day.get(ts[:10], 0) + cnt
                self.last_offsets[DEVICE_KEY] = new_offset
        if new_data_found:
            totals = self.get_totals()
            if self.update_callback:
                try: self.update_callback(totals)
                except Exception: pass
            self.json_writer.trigger()
    def get_totals(self):
        with self.data_lock: return aggregate_counts(self.per_day)
    def read_history(self):
        with self.data_lock: per_day_copy = self.per_day.copy()
        year_data = defaultdict(lambda: {"total": 0, "months": defaultdict(int)})
        for d, v in per_day_copy.items(): yr, ym = d[:4], d[:7]; year_data[yr]["total"] += int(v); year_data[yr]["months"][ym] += int(v)
        return year_data
    def clear_csv_and_reset_offset(self):
        try:
            with safe_file_access(CSV_FILE, "w", "exclusive") as f: f.write(CSV_HEADER)
            with self.data_lock: self.last_offsets[DEVICE_KEY] = 0
            self.json_writer.trigger(); self.process_incremental(); return True
        except Exception as e: messagebox.showerror("清理失败", str(e)); return False
# </editor-fold>
# <editor-fold desc="SpeedTester">
class SpeedTester:
    def __init__(self, get_today_count_callable):
        self._get_today = get_today_count_callable; self.active, self.start_time, self.start_count = False, None, 0
        self.lock = threading.Lock(); self.history = []; self.current_speed = 0.0
    def start(self):
        with self.lock: self.active = True; self.start_time = time.monotonic(); self.start_count = int(self._get_today()); self.history.clear()
    def stop(self):
        with self.lock:
            self.active = False; secs = time.monotonic() - self.start_time if self.start_time else 0
            total_typed = int(self._get_today()) - self.start_count; return (total_typed / secs * 60) if secs > 0 else 0.0
    def update_and_get_speed(self):
        with self.lock:
            if not self.active: return 0.0
            now = time.monotonic(); current_total = int(self._get_today()); self.history.append((now, current_total))
            self.history = [h for h in self.history if now - h[0] < 60]
            if len(self.history) < 2: return 0.0
            time_span = self.history[-1][0] - self.history[0][0]; count_span = self.history[-1][1] - self.history[0][1]
            if time_span < 1: return self.current_speed
            self.current_speed = (count_span / time_span) * 60; return self.current_speed
# </editor-fold>

# ================== 悬浮窗 =====================
class FloatingWindow(tk.Toplevel):
    def __init__(self, master, get_today_callable, speed_tester):
        super().__init__(master)
        self.master = master
        self.get_today = get_today_callable
        self.speed_tester = speed_tester
        self.settings = master.settings_manager

        self.withdraw(); self.overrideredirect(True); self.attributes("-topmost", True)
        self.attributes("-alpha", 0.95); self.visible = True
        self._drag_data = {"x": 0, "y": 0}; self.snap_threshold = 25
        self.is_animating = False
        self.last_count = -1

        self.font_main = tkFont.Font(family=FLOAT_FONT_NAME, size=FLOAT_FONT_SIZE, weight="bold")
        self.font_speed = tkFont.Font(family=FLOAT_FONT_NAME, size=FLOAT_FONT_SPEED_SIZE)
        
        if self.font_main.actual("family").lower() != FLOAT_FONT_NAME.lower():
            print(f"--- FONT WARNING: '{FLOAT_FONT_NAME}' not found. Falling back to '{self.font_main.actual('family')}'. Check font name in Windows Settings. ---")

        self.transparent_color = "#00FF01"
        self.config(bg=self.transparent_color); self.attributes("-transparentcolor", self.transparent_color)

        self.canvas = tk.Canvas(self, bg=self.transparent_color, highlightthickness=0); self.canvas.pack()

        self._build_context_menu()
        self.canvas.bind("<ButtonPress-1>", self._on_drag_start); self.canvas.bind("<B1-Motion>", self._on_drag_move)
        self.canvas.bind("<ButtonRelease-1>", self._on_drag_end); self.canvas.bind("<Double-Button-1>", self._on_double_click)
        self.canvas.bind("<Button-3>", self._show_context_menu)

        self.deiconify(); self.update_layout(); self.update_speed_display()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight(); self.geometry(f"+{sw - 220}+{sh - 120}")

    def _build_context_menu(self):
        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(label="显示主界面", command=self.master.show_main_window)
        self.context_menu.add_command(label="切换速度显示", command=self.toggle_speed_view)
        
        self.theme_menu = tk.Menu(self.context_menu, tearoff=0)
        self.theme_var = tk.StringVar(value=self.settings.get("theme"))
        self.theme_menu.add_radiobutton(label="自动跟随系统", variable=self.theme_var, value="auto", command=lambda: self.master.on_theme_change("auto"))
        self.theme_menu.add_radiobutton(label="亮色", variable=self.theme_var, value="light", command=lambda: self.master.on_theme_change("light"))
        self.theme_menu.add_radiobutton(label="暗色", variable=self.theme_var, value="dark", command=lambda: self.master.on_theme_change("dark"))
        self.context_menu.add_cascade(label="主题", menu=self.theme_menu)

        self.context_menu.add_separator()
        self.context_menu.add_command(label="退出", command=self.master.quit_app)

    def _show_context_menu(self, event):
        self.theme_var.set(self.settings.get("theme"))
        self.context_menu.post(event.x_root, event.y_root)

    def update_layout(self, text_override=None):
        if not self.winfo_exists() or (self.is_animating and not text_override): return
        style = ttk.Style.get_instance(); bg_color = style.colors.primary; fg_color = style.colors.selectfg
        show_speed = self.settings.get("float_show_speed")
        
        main_text = text_override if text_override is not None else f"Today: {self.get_today():,}"
        main_text_width = self.font_main.measure(main_text)
        
        width = main_text_width + 30; height = 38
        
        if show_speed:
            speed_text = f"{self.speed_tester.current_speed:.1f} 字/分"
            speed_text_width = self.font_speed.measure(speed_text)
            width = max(width, speed_text_width + 30); height = 60

        radius = 12
        self.canvas.config(width=width, height=height); self.canvas.delete("all")
        self.canvas.create_polygon(self._rounded_rect_points(0, 0, width, height, radius), fill=bg_color, smooth=True, tags="bg_shape")
        self.canvas.create_text(width / 2, height / 2 if not show_speed else 22, text=main_text, fill=fg_color, font=self.font_main, tags="main_text")
        if show_speed:
            self.canvas.create_text(width / 2, 45, text=speed_text, fill=fg_color, font=self.font_speed, tags="speed_text")

    def scramble_animation(self, new_count):
        if not self.winfo_exists() or self.is_animating: return
        self.is_animating = True

        old_str = f"{self.last_count:,}" if self.last_count != -1 else ""
        new_str = f"{new_count:,}"
        
        # 使新旧字符串等长，便于比较
        max_len = max(len(old_str), len(new_str))
        old_str = old_str.rjust(max_len)
        new_str = new_str.rjust(max_len)
        
        duration = 300  # ms
        interval = 30 # ms
        steps = duration // interval
        
        def animate(step):
            if not self.winfo_exists(): self.is_animating = False; return
            
            if step > steps:
                self.is_animating = False; self.last_count = new_count
                self.update_layout(); return

            display_str = ""
            for i in range(max_len):
                if old_str[i] == new_str[i]:
                    display_str += new_str[i]
                else:
                    # 动画中
                    if step < steps / 2:
                        display_str += random.choice("0123456789, ")
                    # 动画后半段，逐渐锁定
                    else:
                        if random.random() > (step - steps/2) / (steps/2):
                             display_str += random.choice("0123456789, ")
                        else:
                             display_str += new_str[i]
            
            self.update_layout(text_override=f"Today: {display_str.lstrip()}")
            self.after(interval, lambda: animate(step + 1))
            
        animate(1)

    def update_count(self, count):
        if not self.winfo_exists() or count == self.last_count: return
        self.lift()
        self.scramble_animation(count)

    # <editor-fold desc="Other FloatingWindow methods (unchanged)">
    def toggle(self):
        if not self.winfo_exists(): return
        if self.visible: self.withdraw()
        else: self.deiconify()
        self.visible = not self.visible
    def _rounded_rect_points(self, x1, y1, x2, y2, r): return [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2, x2 - r, y2, x1 + r, y2, x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]
    def _on_drag_start(self, e): self._drag_data = {"x": e.x, "y": e.y}
    def _on_drag_move(self, e): self.geometry(f"+{self.winfo_x() + (e.x - self._drag_data['x'])}+{self.winfo_y() + (e.y - self._drag_data['y'])}")
    def _on_drag_end(self, e):
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight(); win_x, win_y = self.winfo_x(), self.winfo_y(); win_w, win_h = self.winfo_width(), self.winfo_height()
        if win_x < self.snap_threshold: win_x = 0
        if win_y < self.snap_threshold: win_y = 0
        if sw - (win_x + win_w) < self.snap_threshold: win_x = sw - win_w
        if sh - (win_y + win_h) < self.snap_threshold: win_y = sh - win_h
        self.geometry(f"+{win_x}+{win_y}")
    def _on_double_click(self, _): self.master.show_main_window()
    def update_speed_display(self):
        if self.winfo_exists() and self.settings.get("float_show_speed"):
            self.speed_tester.update_and_get_speed(); self.update_layout()
        self.after(1000, self.update_speed_display)
    def toggle_speed_view(self):
        if not self.winfo_exists(): return
        new_state = not self.settings.get("float_show_speed"); self.settings.save_setting("float_show_speed", new_state)
        self.speed_tester.active = new_state
        if new_state: self.speed_tester.start()
        self.update_layout()
    # </editor-fold>

# <editor-fold desc="HeatmapWindow (compatible)">
class HeatmapWindow(tk.Toplevel):
    def __init__(self, parent, data_processor):
        super().__init__(parent); self.data_processor = data_processor
        self.title("贡献热力图"); self.transient(parent)
        
        style = ttk.Style.get_instance(); self.colors = [style.colors.light, style.colors.info, style.colors.warning, style.colors.danger]
        self.bg_color = style.colors.bg; self.fg_color = style.colors.fg; self.config(bg=self.bg_color)
        
        self.view_var = tk.StringVar(value="year")
        button_bar = ttk.Frame(self, bootstyle="secondary", padding=(5, 5, 5, 0)); button_bar.pack(fill=X)
        
        radio_frame = ttk.Frame(button_bar); radio_frame.pack(expand=True)
        ttk.Radiobutton(radio_frame, text="年度", variable=self.view_var, value="year", command=self.redraw_heatmap, bootstyle="outline-toolbutton-primary").pack(side=LEFT, expand=True, fill=X)
        ttk.Radiobutton(radio_frame, text="本月", variable=self.view_var, value="month", command=self.redraw_heatmap, bootstyle="outline-toolbutton-primary").pack(side=LEFT, expand=True, fill=X)
        ttk.Radiobutton(radio_frame, text="本周", variable=self.view_var, value="week", command=self.redraw_heatmap, bootstyle="outline-toolbutton-primary").pack(side=LEFT, expand=True, fill=X)

        self.canvas = tk.Canvas(self, bg=self.bg_color, highlightthickness=0); self.canvas.pack(fill=BOTH, expand=True, padx=15, pady=15)
        self.redraw_heatmap()

    def redraw_heatmap(self, _=None):
        self.canvas.delete("all")
        view = self.view_var.get()
        if view == "year": self.geometry("820x220"); self.draw_year_heatmap()
        elif view == "month": self.geometry("320x280"); self.draw_month_heatmap()
        elif view == "week": self.geometry("400x120"); self.draw_week_heatmap()

    def draw_year_heatmap(self):
        per_day_data = self.data_processor.per_day; today = date.today()
        start_date = today - timedelta(days=365); max_count = max(per_day_data.values(), default=1)
        cell_size, spacing = 13, 3; current_date = start_date; offset_x, offset_y = 15, 25
        for i in range(366):
            if current_date > today: break
            day_key = current_date.strftime("%Y-%m-%d"); count = per_day_data.get(day_key, 0)
            week_of_year = int(current_date.strftime("%U")); start_week_of_year = int(start_date.strftime("%U"))
            col = week_of_year - start_week_of_year if week_of_year >= start_week_of_year else 52 - start_week_of_year + week_of_year + 1
            row = current_date.weekday() + 1; row = 0 if row == 7 else row
            x = offset_x + col * (cell_size + spacing); y = offset_y + row * (cell_size + spacing)
            color_index = 0
            if count > 0: ratio = count / max_count; color_index = 3 if ratio > 0.6 else 2 if ratio > 0.3 else 1
            self.canvas.create_rectangle(x, y, x + cell_size, y + cell_size, fill=self.colors[color_index], outline=self.bg_color, width=1)
            current_date += timedelta(days=1)
            
    def draw_month_heatmap(self):
        per_day_data = self.data_processor.per_day; today = date.today()
        first_day_of_month = today.replace(day=1)
        next_month = (first_day_of_month.replace(day=28) + timedelta(days=4))
        days_in_month = (next_month - timedelta(days=next_month.day)).day
        start_weekday = first_day_of_month.weekday()
        
        max_count = max((v for k, v in per_day_data.items() if k.startswith(today.strftime("%Y-%m"))), default=1)
        cell_size, spacing = 40, 5; day_names = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
        for i, name in enumerate(day_names): self.canvas.create_text(i * (cell_size + spacing) + cell_size/2, 10, text=name, fill=self.fg_color)
        
        current_day = 1
        for row in range(6):
            for col in range(7):
                if (row == 0 and col < start_weekday) or current_day > days_in_month: continue
                day_key = f"{today.strftime('%Y-%m')}-{current_day:02d}"
                count = per_day_data.get(day_key, 0)
                x1, y1 = col * (cell_size+spacing), 25 + row * (cell_size+spacing)
                color_index = 0
                if count > 0: ratio = count / max_count; color_index = 3 if ratio > 0.6 else 2 if ratio > 0.3 else 1
                self.canvas.create_rectangle(x1, y1, x1+cell_size, y1+cell_size, fill=self.colors[color_index], outline=self.bg_color, width=1)
                self.canvas.create_text(x1+cell_size/2, y1+cell_size/2, text=str(current_day), fill=self.fg_color)
                current_day += 1

    def draw_week_heatmap(self):
        per_day_data = self.data_processor.per_day; today = date.today()
        start_of_week = today - timedelta(days=today.weekday())
        week_data = { (start_of_week + timedelta(days=i)).strftime("%Y-%m-%d"): per_day_data.get((start_of_week + timedelta(days=i)).strftime("%Y-%m-%d"), 0) for i in range(7) }
        max_count = max(week_data.values(), default=1); bar_width, spacing = 45, 10
        day_names = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
        for i, (day_key, count) in enumerate(week_data.items()):
            x1 = i * (bar_width+spacing); y1_max = 80
            height = (count / max_count) * y1_max if count > 0 else 0
            color_index = 0
            if count > 0: ratio = count / max_count; color_index = 3 if ratio > 0.6 else 2 if ratio > 0.3 else 1
            self.canvas.create_rectangle(x1, y1_max - height, x1 + bar_width, y1_max, fill=self.colors[color_index], outline="")
            self.canvas.create_text(x1 + bar_width/2, y1_max + 10, text=day_names[i], fill=self.fg_color)
            self.canvas.create_text(x1 + bar_width/2, y1_max - height - 10, text=f"{count:,}", fill=self.fg_color)
# </editor-fold>
# <editor-fold desc="HistoryWindow">
class HistoryWindow(tk.Toplevel):
    def __init__(self, parent, data_processor):
        super().__init__(parent); self.data_processor = data_processor; self.title("历史记录"); self.geometry("600x450"); self.transient(parent)
        year_data = self.data_processor.read_history(); notebook = ttk.Notebook(self, bootstyle="primary"); notebook.pack(fill='both', expand=True, padx=10, pady=10)
        years = sorted(year_data.keys(), reverse=True)
        if not years: ttk.Label(self, text="暂无历史数据").pack(pady=30); return
        for yr in years:
            frame = ttk.Frame(notebook, padding=10); notebook.add(frame, text=f" {yr}年 ")
            tree = ttk.Treeview(frame, columns=("时间段", "字数"), show="headings", bootstyle="primary"); tree.heading("时间段", text="时间段"); tree.heading("字数", text="字数")
            tree.column("时间段", width=200, anchor='w'); tree.column("字数", width=100, anchor='e')
            yt = int(year_data[yr]["total"]); total_id = tree.insert("", "end", values=(f"【{yr}年总计】", f"{yt:,}"), tags=("total",))
            months = sorted(year_data[yr]["months"].keys(), reverse=True)
            for m in months:
                mv = int(year_data[yr]["months"][m]); month_id = tree.insert(total_id, "end", values=(f"    {m}", f"{mv:,}"), tags=("month",))
                tree.insert(month_id, "end", values=("", ""), iid=f"dummy_{m}")
            tree.tag_configure("total", font=("Microsoft YaHei UI", 10, "bold")); tree.pack(fill="both", expand=True)
            tree.bind("<<TreeviewOpen>>", lambda e, t=tree, y=yr: self.on_month_expand(e, t, y))
    def on_month_expand(self, event, tree, year):
        item_id = tree.focus()
        if not tree.item(item_id, "tags"): return
        month_str = tree.item(item_id, "values")[0].strip(); dummy_id = f"dummy_{month_str}"
        if not tree.exists(dummy_id): return
        tree.delete(dummy_id)
        per_day_copy = self.data_processor.per_day.copy()
        days_in_month = sorted([(d, v) for d, v in per_day_copy.items() if d.startswith(month_str)], reverse=True)
        for d, v in days_in_month: tree.insert(item_id, "end", values=(f"        {d}", f"{int(v):,}"))
# </editor-fold>

# ================== 主应用 ==================
class Application(ttk.Window):
    def __init__(self):
        self.data_processor = DataProcessor(update_callback=self._on_data_updated)
        self.settings_manager = SettingsManager(JSON_FILE, self.data_processor.json_writer.trigger)
        self.data_processor.set_settings_manager(self.settings_manager)

        self.theme_name = self.get_initial_theme()
        super().__init__(themename=self.theme_name)
        
        self.title("字数统计工具"); self.geometry("400x280"); self.resizable(False, False)

        self.speed_tester = SpeedTester(self._get_today_count)
        self.data_processor._load_initial_data()

        self._build_ui(); self.counts = self.data_processor.get_totals()
        
        self.floating = FloatingWindow(self, self._get_today_count, self.speed_tester)
        self._start_watch()

        schedule.every().day.at("00:00").do(self._clear_csv_job)
        threading.Thread(target=self._schedule_loop, daemon=True).start()

        self.after(100, self.update_ui_labels)
        self.after(5000, self.check_windows_theme_periodically)
        self.protocol("WM_DELETE_WINDOW", self.hide_main_window)

    def get_initial_theme(self):
        choice = self.settings_manager.get("theme")
        if choice == "auto": return DARK_THEME if self.is_windows_dark_mode() else LIGHT_THEME
        return DARK_THEME if choice == "dark" else LIGHT_THEME

    def _build_ui(self):
        wrap = ttk.Frame(self, padding=20); wrap.pack(fill="both", expand=True)
        header_font = ("Microsoft YaHei UI", 12, "bold")
        self.lbl_day = ttk.Label(wrap, text="当日字数: 0", font=header_font, bootstyle="primary")
        self.lbl_month = ttk.Label(wrap, text="本月字数: 0")
        self.lbl_year = ttk.Label(wrap, text="本年字数: 0")
        self.lbl_total = ttk.Label(wrap, text="累计字数: 0")
        self.lbl_day.pack(anchor="w", pady=(0, 10)); self.lbl_month.pack(anchor="w", pady=2)
        self.lbl_year.pack(anchor="w", pady=2); self.lbl_total.pack(anchor="w", pady=(2, 0))
        ttk.Separator(wrap, orient=HORIZONTAL).pack(fill=X, pady=20)
        btn_frame = ttk.Frame(wrap); btn_frame.pack(fill=X, side=BOTTOM, pady=(10,0))
        ttk.Button(btn_frame, text="历史记录", command=lambda: HistoryWindow(self, self.data_processor)).pack(side=LEFT, expand=True, padx=5)
        ttk.Button(btn_frame, text="热力图", command=lambda: HeatmapWindow(self, self.data_processor)).pack(side=RIGHT, expand=True, padx=5)

    def _on_data_updated(self, totals):
        self.counts = totals; self.update_ui_labels()
        if self.floating and self.floating.winfo_exists(): self.floating.update_count(totals['today'])

    def _start_watch(self):
        self.observer = Observer(); self.observer.schedule(CSVHandler(self.data_processor.process_incremental), os.path.dirname(CSV_FILE) or ".", recursive=False); self.observer.start()

    def quit_app(self):
        self.data_processor.json_writer.trigger(); time.sleep(0.1)
        self.observer.stop()
        if self.observer.is_alive(): self.observer.join(timeout=1.0)
        self.destroy(); os._exit(0)
    
    def on_theme_change(self, choice):
        self.settings_manager.save_setting("theme", choice)
        self.apply_theme(self.get_initial_theme())

    def apply_theme(self, theme_name):
        if self.style.theme.name == theme_name: return
        self.style.theme_use(theme_name)
        if self.floating and self.floating.winfo_exists(): self.floating.update_layout()
        for win in self.winfo_children():
            if isinstance(win, (HistoryWindow, HeatmapWindow)): win.destroy()

    def check_windows_theme_periodically(self):
        if self.settings_manager.get("theme") == "auto":
            new_theme = DARK_THEME if self.is_windows_dark_mode() else LIGHT_THEME
            self.apply_theme(new_theme)
        self.after(5000, self.check_windows_theme_periodically)

    # <editor-fold desc="Other Application methods (unchanged)">
    def update_ui_labels(self):
        c = self.counts; self.lbl_day.config(text=f"当日字数: {c['today']:,}"); self.lbl_month.config(text=f"本月字数: {c['month']:,}")
        self.lbl_year.config(text=f"本年字数: {c['year']:,}"); self.lbl_total.config(text=f"累计字数: {c['total']:,}")
    def _get_today_count(self): return self.counts.get("today", 0)
    def hide_main_window(self): self.withdraw()
    def show_main_window(self):
        self.deiconify(); self.lift(); self.attributes('-topmost', 1); self.after(100, lambda: self.attributes('-topmost', 0))
    def _schedule_loop(self):
        while True: schedule.run_pending(); time.sleep(1)
    def _clear_csv_job(self): self.data_processor.clear_csv_and_reset_offset()
    def is_windows_dark_mode(self):
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\CurrentVersion\Themes\Personalize')
            value, _ = winreg.QueryValueEx(key, 'AppsUseLightTheme'); winreg.CloseKey(key); return value == 0
        except Exception: return False
    # </editor-fold>

def main():
    data_processor = DataProcessor()
    settings_manager = SettingsManager(JSON_FILE, data_processor.json_writer.trigger)
    data_processor.set_settings_manager(settings_manager)
    
    ensure_dirs(); init_csv_if_missing(); init_json_if_missing(settings_manager)

    app = Application()
    app.hide_main_window()
    app.mainloop()

if __name__ == "__main__":
    main()