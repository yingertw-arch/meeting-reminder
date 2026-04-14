"""
會議提醒管理 - Meeting Reminder
依賴：pystray、Pillow（首次執行時自動安裝）
"""
# ── 自動安裝依賴 ──────────────────────────────────────────
def _ensure_deps():
    import importlib, subprocess, sys
    for pkg, imp in [("pystray", "pystray"), ("Pillow", "PIL")]:
        try:
            importlib.import_module(imp)
        except ImportError:
            print(f"正在安裝 {pkg}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "-q"])

_ensure_deps()

import tkinter as tk
from tkinter import messagebox
import threading
import time
from datetime import datetime, timedelta
import math
import json
import os
import sys
import ctypes
import subprocess
import pystray
from PIL import Image, ImageDraw

# ── 顏色主題 ──────────────────────────────────────────────
BG_DARK  = "#0a0a1a"
BG_CARD  = "#12122a"
BG_ROW   = "#1a1a35"
ACCENT   = "#ff3366"
ACCENT2  = "#ff6600"
GOLD     = "#ffd700"
WHITE    = "#ffffff"
GRAY     = "#8888aa"
GREEN    = "#00ff88"
BLUE     = "#4488ff"
PURPLE   = "#5533cc"

SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
SAVE_FILE   = os.path.join(SCRIPT_DIR, "reminders.json")
APP_NAME    = "MeetingReminder"
WIN_TITLE   = "🔔 會議提醒管理"

STARTUP_FOLDER = os.path.join(
    os.environ.get("APPDATA", os.path.expanduser("~")),
    "Microsoft", "Windows", "Start Menu", "Programs", "Startup"
)
STARTUP_VBS = os.path.join(STARTUP_FOLDER, f"{APP_NAME}.vbs")

WEEKDAY_NAMES = ["一", "二", "三", "四", "五", "六", "日"]
REC_LABELS = {
    "none":     "單次",
    "daily":    "每天",
    "weekdays": "週一～五",
    "weekly":   "每週",
    "monthly":  "每月",
}


# ══════════════════════════════════════════════════════════
#  單一執行個體（防止重複開啟）
# ══════════════════════════════════════════════════════════
_MUTEX_NAME   = "Global\\MeetingReminder_SingleInstance_v2"
_mutex_handle = None

def acquire_instance_lock():
    """回傳 True 表示成功取得鎖（第一個執行個體）"""
    global _mutex_handle
    _mutex_handle = ctypes.windll.kernel32.CreateMutexW(None, False, _MUTEX_NAME)
    return ctypes.windll.kernel32.GetLastError() != 183  # 183 = ERROR_ALREADY_EXISTS

def show_existing_window():
    """若程式已在執行中，把它的視窗帶到最上層"""
    hwnd = ctypes.windll.user32.FindWindowW(None, WIN_TITLE)
    if hwnd:
        ctypes.windll.user32.ShowWindow(hwnd, 9)       # SW_RESTORE
        ctypes.windll.user32.SetForegroundWindow(hwnd)


# ══════════════════════════════════════════════════════════
#  pythonw 路徑
# ══════════════════════════════════════════════════════════
def _pythonw():
    w = sys.executable.replace("python.exe", "pythonw.exe")
    return w if os.path.exists(w) else sys.executable


# ══════════════════════════════════════════════════════════
#  開機自動啟動（登錄機碼 HKCU，不需要管理員權限）
# ══════════════════════════════════════════════════════════
import winreg
REG_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"

def is_startup_enabled():
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_KEY, 0, winreg.KEY_READ)
        winreg.QueryValueEx(key, APP_NAME)
        winreg.CloseKey(key)
        return True
    except (FileNotFoundError, OSError):
        return False

def set_startup(enabled: bool):
    py     = _pythonw()
    script = os.path.abspath(__file__)
    key    = winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_KEY, 0, winreg.KEY_SET_VALUE)
    if enabled:
        cmd = f'"{py}" "{script}" --minimized'
        winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, cmd)
        winreg.CloseKey(key)
        # 同時清除舊的 VBS / 工作排程器（若存在）
        try:
            os.remove(STARTUP_VBS)
        except FileNotFoundError:
            pass
        subprocess.run(
            ["schtasks", "/delete", "/tn", "MeetingReminderAutoStart", "/f"],
            capture_output=True
        )
        return cmd
    else:
        try:
            winreg.DeleteValue(key, APP_NAME)
        except (FileNotFoundError, OSError):
            pass
        winreg.CloseKey(key)
        return None

def diagnose_startup():
    py     = _pythonw()
    script = os.path.abspath(__file__)
    lines  = [
        f"pythonw.exe：{'✅ 存在' if os.path.exists(py) else '❌ 找不到'}",
        f"  {py}",
        f"腳本：{'✅ 存在' if os.path.exists(script) else '❌ 找不到'}",
        f"  {script}",
    ]
    # 讀取登錄機碼
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_KEY, 0, winreg.KEY_READ)
        val, _ = winreg.QueryValueEx(key, APP_NAME)
        winreg.CloseKey(key)
        lines.append(f"\n✅ 登錄機碼已設定：\n  {val}")
    except (FileNotFoundError, OSError):
        lines.append("\n❌ 登錄機碼：未設定")
    return "\n".join(lines)


# ══════════════════════════════════════════════════════════
#  桌面捷徑
# ══════════════════════════════════════════════════════════
def create_desktop_shortcut():
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    lnk = os.path.join(desktop, "會議提醒.lnk")
    py  = _pythonw()
    scr = os.path.abspath(__file__)
    ps  = f"""
$ws = New-Object -ComObject WScript.Shell
$sc = $ws.CreateShortcut('{lnk}')
$sc.TargetPath = '{py}'
$sc.Arguments = '"{scr}"'
$sc.WorkingDirectory = '{SCRIPT_DIR}'
$sc.IconLocation = '{py},0'
$sc.Description = '會議提醒小幫手'
$sc.Save()
"""
    r = subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                       capture_output=True, text=True)
    if r.returncode != 0:
        raise RuntimeError(r.stderr.strip())
    return lnk


# ══════════════════════════════════════════════════════════
#  資料
# ══════════════════════════════════════════════════════════
def load_reminders():
    if os.path.exists(SAVE_FILE):
        try:
            with open(SAVE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return []

def save_reminders(data):
    with open(SAVE_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# ══════════════════════════════════════════════════════════
#  週期描述 & 下次時間
# ══════════════════════════════════════════════════════════
def rec_desc(rem):
    rec = rem.get("recurrence", "none")
    t   = rem.get("time", "")
    if rec == "none":
        return rem.get("datetime", "")
    if rec == "weekly":
        days = "".join(WEEKDAY_NAMES[d] for d in sorted(rem.get("weekdays", [])))
        return f"每週 {days}  {t}"
    if rec == "monthly":
        return f"每月 {rem.get('monthday', 1)} 日  {t}"
    return f"{REC_LABELS.get(rec, rec)}  {t}"

def next_alarm_dt(rem):
    rec = rem.get("recurrence", "none")
    pre = rem.get("pre_min", 0)
    now = datetime.now()
    if rec == "none":
        try:
            dt = datetime.strptime(rem["datetime"], "%Y-%m-%d %H:%M")
            return dt - timedelta(minutes=pre)
        except Exception:
            return None
    h, m  = map(int, rem.get("time", "0:0").split(":"))
    total = h * 60 + m - pre
    if total < 0:
        total = 0
    ah, am = divmod(total, 60)
    base = now.replace(hour=ah, minute=am, second=0, microsecond=0)
    if rec == "daily":
        return base if base > now else base + timedelta(days=1)
    if rec == "weekdays":
        for i in range(8):
            d = base + timedelta(days=i)
            if d > now and d.weekday() < 5:
                return d
    if rec == "weekly":
        wdays = rem.get("weekdays", [])
        for i in range(8):
            d = base + timedelta(days=i)
            if d > now and d.weekday() in wdays:
                return d
    if rec == "monthly":
        mday = rem.get("monthday", 1)
        for add_m in range(3):
            y, mo = now.year, now.month + add_m
            if mo > 12:
                y, mo = y + 1, mo - 12
            try:
                d = datetime(y, mo, mday, ah, am)
                if d > now:
                    return d
            except ValueError:
                pass
    return None

def sort_key(rem):
    nxt = next_alarm_dt(rem)
    return nxt if nxt else datetime.max

def find_expired(reminders):
    """找出單次且已過期的提醒（過了警報時間超過 60 秒）"""
    now = datetime.now()
    expired = []
    for i, rem in enumerate(reminders):
        if rem.get("recurrence", "none") != "none":
            continue
        try:
            dt = datetime.strptime(rem["datetime"], "%Y-%m-%d %H:%M")
            alarm_dt = dt - timedelta(minutes=rem.get("pre_min", 0))
            if now >= alarm_dt + timedelta(seconds=60):
                expired.append(i)
        except Exception:
            pass
    return expired


# ══════════════════════════════════════════════════════════
#  新增/編輯 對話框
# ══════════════════════════════════════════════════════════
class ReminderDialog:
    def __init__(self, parent, on_save, existing=None):
        self.on_save = on_save
        self.win = tk.Toplevel(parent)
        self.win.title("新增提醒" if existing is None else "編輯提醒")
        self.win.geometry("480x500")
        self.win.configure(bg=BG_DARK)
        self.win.resizable(False, False)
        self.win.grab_set()
        self.win.attributes("-topmost", True)
        self._build(existing or {})

    def _spin(self, p, var, f, t, w=5):
        tk.Spinbox(p, textvariable=var, from_=f, to=t, width=w,
                   font=("微軟正黑體", 12), bg=BG_CARD, fg=WHITE,
                   buttonbackground=BG_CARD, relief="flat",
                   insertbackground=WHITE).pack(side="left", padx=2, ipady=4)

    def _lbl(self, p, text):
        return tk.Label(p, text=text, font=("微軟正黑體", 11),
                        bg=BG_DARK, fg=GRAY, anchor="w")

    def _build(self, ex):
        w   = self.win
        now = datetime.now()
        tk.Label(w, text="📋  設定提醒", font=("微軟正黑體", 15, "bold"),
                 bg=ACCENT, fg=WHITE).pack(fill="x", ipady=8)

        main = tk.Frame(w, bg=BG_DARK, padx=28, pady=14)
        main.pack(fill="both", expand=True)
        ekw = dict(font=("微軟正黑體", 12), bg=BG_CARD, fg=WHITE,
                   insertbackground=WHITE, relief="flat", bd=0)

        self._lbl(main, "📋 會議名稱").grid(row=0, column=0, sticky="w", pady=6)
        self.name_var = tk.StringVar(value=ex.get("name", "週會"))
        tk.Entry(main, textvariable=self.name_var, width=22, **ekw).grid(
            row=0, column=1, columnspan=3, sticky="w", padx=8, ipady=6, pady=6)

        self._lbl(main, "🔁 週期").grid(row=1, column=0, sticky="w", pady=6)
        self.rec_var = tk.StringVar(value=ex.get("recurrence", "none"))
        rf = tk.Frame(main, bg=BG_DARK)
        rf.grid(row=1, column=1, columnspan=3, sticky="w", padx=8, pady=6)
        for val, lbl in REC_LABELS.items():
            tk.Radiobutton(rf, text=lbl, variable=self.rec_var, value=val,
                           command=self._on_rec,
                           font=("微軟正黑體", 11), bg=BG_DARK, fg=WHITE,
                           selectcolor=BG_CARD, activebackground=BG_DARK,
                           activeforeground=GOLD, relief="flat"
                           ).pack(side="left", padx=4)

        self.dyn = tk.Frame(main, bg=BG_DARK)
        self.dyn.grid(row=2, column=0, columnspan=4, sticky="w", pady=2)
        self._ex, self._now = ex, now
        self._build_dyn()

        self._lbl(main, "🕐 時間").grid(row=3, column=0, sticky="w", pady=6)
        tf = tk.Frame(main, bg=BG_DARK)
        tf.grid(row=3, column=1, columnspan=3, sticky="w", padx=8, pady=6)
        if ex.get("time"):
            dh, dm = map(int, ex["time"].split(":"))
        elif ex.get("datetime"):
            _dt = datetime.strptime(ex["datetime"], "%Y-%m-%d %H:%M")
            dh, dm = _dt.hour, _dt.minute
        else:
            dh, dm = now.hour, now.minute
        self.h_var  = tk.StringVar(value=str(dh))
        self.mi_var = tk.StringVar(value=str(dm))
        self._spin(tf, self.h_var, 0, 23, 4)
        tk.Label(tf, text="時", bg=BG_DARK, fg=WHITE,
                 font=("微軟正黑體", 12)).pack(side="left", padx=2)
        self._spin(tf, self.mi_var, 0, 59, 4)
        tk.Label(tf, text="分", bg=BG_DARK, fg=WHITE,
                 font=("微軟正黑體", 12)).pack(side="left", padx=2)

        self._lbl(main, "⚡ 提前提醒").grid(row=4, column=0, sticky="w", pady=6)
        pf = tk.Frame(main, bg=BG_DARK)
        pf.grid(row=4, column=1, columnspan=3, sticky="w", padx=8, pady=6)
        self.pre_var = tk.StringVar(value=str(ex.get("pre_min", 5)))
        self._spin(pf, self.pre_var, 0, 120, 5)
        tk.Label(pf, text="分鐘前", bg=BG_DARK, fg=WHITE,
                 font=("微軟正黑體", 12)).pack(side="left", padx=4)

        bf = tk.Frame(w, bg=BG_DARK, pady=12)
        bf.pack()
        tk.Button(bf, text="✅  儲存", font=("微軟正黑體", 12, "bold"),
                  bg=GREEN, fg="#000", relief="flat", padx=20, pady=8,
                  cursor="hand2", command=self._save).pack(side="left", padx=8)
        tk.Button(bf, text="✖  取消", font=("微軟正黑體", 12),
                  bg=GRAY, fg=WHITE, relief="flat", padx=20, pady=8,
                  cursor="hand2", command=self.win.destroy).pack(side="left", padx=8)

    def _build_dyn(self):
        for w in self.dyn.winfo_children():
            w.destroy()
        ex, now, rec = self._ex, self._now, self.rec_var.get()
        if rec == "none":
            self._lbl(self.dyn, "📅 日期").pack(side="left", padx=(0, 8))
            df = tk.Frame(self.dyn, bg=BG_DARK)
            df.pack(side="left")
            _dt = datetime.strptime(ex["datetime"], "%Y-%m-%d %H:%M") if ex.get("datetime") else now
            self.y_var  = tk.StringVar(value=str(_dt.year))
            self.mo_var = tk.StringVar(value=str(_dt.month))
            self.d_var  = tk.StringVar(value=str(_dt.day))
            self._spin(df, self.y_var, 2024, 2099, 6)
            tk.Label(df, text="/", bg=BG_DARK, fg=WHITE,
                     font=("微軟正黑體", 13)).pack(side="left")
            self._spin(df, self.mo_var, 1, 12, 4)
            tk.Label(df, text="/", bg=BG_DARK, fg=WHITE,
                     font=("微軟正黑體", 13)).pack(side="left")
            self._spin(df, self.d_var, 1, 31, 4)
        elif rec == "weekly":
            self._lbl(self.dyn, "📆 星期").pack(side="left", padx=(0, 8))
            self.wd_vars = []
            sel = ex.get("weekdays", [])
            for i, name in enumerate(WEEKDAY_NAMES):
                v = tk.BooleanVar(value=(i in sel))
                self.wd_vars.append(v)
                tk.Checkbutton(self.dyn, text=name, variable=v,
                               font=("微軟正黑體", 12), bg=BG_DARK, fg=WHITE,
                               selectcolor=BG_CARD, activebackground=BG_DARK,
                               activeforeground=GOLD, relief="flat"
                               ).pack(side="left", padx=2)
        elif rec == "monthly":
            self._lbl(self.dyn, "📆 每月幾日").pack(side="left", padx=(0, 8))
            self.mday_var = tk.StringVar(value=str(ex.get("monthday", self._now.day)))
            self._spin(self.dyn, self.mday_var, 1, 31, 4)
            tk.Label(self.dyn, text="日", bg=BG_DARK, fg=WHITE,
                     font=("微軟正黑體", 12)).pack(side="left", padx=4)

    def _on_rec(self):
        self._build_dyn()

    def _save(self):
        name = self.name_var.get().strip() or "會議"
        rec  = self.rec_var.get()
        pre  = int(self.pre_var.get())
        h, mi = int(self.h_var.get()), int(self.mi_var.get())
        data = {"name": name, "recurrence": rec,
                "time": f"{h:02d}:{mi:02d}", "pre_min": pre,
                "enabled": True, "last_fired": ""}
        if rec == "none":
            try:
                dt = datetime(int(self.y_var.get()), int(self.mo_var.get()),
                              int(self.d_var.get()), h, mi)
            except ValueError as e:
                messagebox.showerror("錯誤", f"日期格式錯誤：{e}", parent=self.win)
                return
            data["datetime"] = dt.strftime("%Y-%m-%d %H:%M")
        elif rec == "weekly":
            days = [i for i, v in enumerate(self.wd_vars) if v.get()]
            if not days:
                messagebox.showwarning("提示", "請至少選一個星期幾", parent=self.win)
                return
            data["weekdays"] = days
        elif rec == "monthly":
            data["monthday"] = int(self.mday_var.get())
        self.on_save(data)
        self.win.destroy()


# ══════════════════════════════════════════════════════════
#  系統匣圖示
# ══════════════════════════════════════════════════════════
def _make_tray_image():
    img  = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    draw.ellipse([2, 2, 62, 62],   fill=(255, 51, 102))
    draw.ellipse([10, 10, 54, 54], fill=(10, 10, 26))
    cx, cy = 32, 32
    draw.line([cx, cy, cx - 6, cy - 10], fill="white", width=3)
    draw.line([cx, cy, cx + 9,  cy - 6], fill="white", width=2)
    draw.ellipse([29, 29, 35, 35], fill="white")
    draw.rectangle([20, 50, 44, 55], fill=(255, 51, 102))
    return img

class TrayIcon:
    def __init__(self, app):
        self.app   = app
        self._icon = None

    def show(self):
        if self._icon:
            return
        menu = pystray.Menu(
            pystray.MenuItem("📋 開啟視窗", self._open),
            pystray.MenuItem("❌ 結束程式", self._quit),
        )
        self._icon = pystray.Icon(APP_NAME, _make_tray_image(), "會議提醒", menu)
        threading.Thread(target=self._icon.run, daemon=True).start()

    def hide(self):
        if self._icon:
            self._icon.stop()
            self._icon = None

    def _open(self, *_):
        self.app.root.after(0, self.app.show_window)

    def _quit(self, *_):
        self.app.root.after(0, self.app._force_quit)


# ══════════════════════════════════════════════════════════
#  主視窗
# ══════════════════════════════════════════════════════════
class MainApp:
    def __init__(self, root):
        self.root = root
        root.title(WIN_TITLE)
        root.geometry("780x520")
        root.configure(bg=BG_DARK)
        root.minsize(680, 400)

        self.reminders       = load_reminders()
        self.monitor_running = True
        self.selected_idx    = None
        self.row_vars        = []
        self.row_frames      = []
        self.tray            = TrayIcon(self)

        self._sort()
        self._build_ui()
        self._refresh_list()
        self._start_monitor()
        root.protocol("WM_DELETE_WINDOW", self._on_close)

        # 啟動後延遲 500ms 檢查過期會議
        root.after(500, self._check_expired_on_startup)

    # ── 排序 ─────────────────────────────────────────────
    def _sort(self):
        self.reminders.sort(key=sort_key)

    # ── UI ───────────────────────────────────────────────
    def _build_ui(self):
        hdr = tk.Frame(self.root, bg=ACCENT)
        hdr.pack(fill="x")
        tk.Label(hdr, text="⏰  會議提醒管理",
                 font=("微軟正黑體", 17, "bold"),
                 bg=ACCENT, fg=WHITE).pack(side="left", padx=18, pady=10)
        tk.Label(hdr, text="支援單次 / 每天 / 每週 / 每月 週期提醒",
                 font=("微軟正黑體", 10),
                 bg=ACCENT, fg="#ffcccc").pack(side="right", padx=18)

        # ── 第一列：主要操作 ──
        tb1 = tk.Frame(self.root, bg=BG_CARD, pady=4)
        tb1.pack(fill="x")

        def btn(parent, text, cmd, color=BLUE):
            tk.Button(parent, text=text, font=("微軟正黑體", 11, "bold"),
                      bg=color, fg=WHITE, relief="flat",
                      padx=12, pady=5, cursor="hand2", command=cmd,
                      activebackground=color, activeforeground=WHITE
                      ).pack(side="left", padx=4, pady=3)

        btn(tb1, "➕  新增",    self._add,     GREEN)
        btn(tb1, "✏️  編輯",    self._edit,    BLUE)
        btn(tb1, "🗑️  刪除",    self._delete,  ACCENT)
        btn(tb1, "🧪  測試警報", self._test,    ACCENT2)
        btn(tb1, "🖱️  桌面捷徑", self._shortcut, PURPLE)
        btn(tb1, "❌  結束程式", self._quit,    "#444455")

        # ── 第二列：設定 ──
        tb2 = tk.Frame(self.root, bg="#0d0d22", pady=3)
        tb2.pack(fill="x")

        self.startup_var = tk.BooleanVar(value=is_startup_enabled())
        tk.Checkbutton(tb2, text="🖥️  開機自動啟動",
                       variable=self.startup_var, command=self._toggle_startup,
                       font=("微軟正黑體", 11), bg="#0d0d22", fg=GOLD,
                       selectcolor=BG_DARK, activebackground="#0d0d22",
                       activeforeground=GOLD, relief="flat", cursor="hand2"
                       ).pack(side="left", padx=10, pady=2)


        hrow = tk.Frame(self.root, bg="#222244")
        hrow.pack(fill="x")
        for text, w in [("", 3), ("會議名稱", 14), ("週期 / 日期時間", 22),
                        ("提前", 6), ("狀態", 9), ("下次提醒倒數", 15)]:
            tk.Label(hrow, text=text, font=("微軟正黑體", 10, "bold"),
                     bg="#222244", fg=GRAY, width=w, anchor="w",
                     padx=6).pack(side="left")

        lf = tk.Frame(self.root, bg=BG_DARK)
        lf.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(lf, bg=BG_DARK, highlightthickness=0)
        sb = tk.Scrollbar(lf, orient="vertical", command=self.canvas.yview)
        self.sf = tk.Frame(self.canvas, bg=BG_DARK)
        self.sf.bind("<Configure>",
                     lambda e: self.canvas.configure(
                         scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.sf, anchor="nw")
        self.canvas.configure(yscrollcommand=sb.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self.status_var = tk.StringVar(value="就緒")
        sf2 = tk.Frame(self.root, bg="#050510", pady=3)
        sf2.pack(fill="x", side="bottom")
        tk.Label(sf2, textvariable=self.status_var,
                 font=("微軟正黑體", 10), bg="#050510", fg=GRAY,
                 anchor="w", padx=10).pack(fill="x")

        self._tick()

    # ── 過期會議清理（啟動時） ────────────────────────────
    def _check_expired_on_startup(self):
        expired_idx = find_expired(self.reminders)
        if not expired_idx:
            return

        names = "\n".join(
            f"  • {self.reminders[i]['name']}  ({self.reminders[i].get('datetime','')})"
            for i in expired_idx
        )
        msg = f"以下 {len(expired_idx)} 筆單次提醒已過期：\n\n{names}\n\n是否刪除這些過期提醒？"

        if messagebox.askyesno("過期提醒清理", msg, icon="warning"):
            for i in sorted(expired_idx, reverse=True):
                self.reminders.pop(i)
            save_reminders(self.reminders)
            self._refresh_list()
            self.status_var.set(f"已清除 {len(expired_idx)} 筆過期提醒")

    # ── 清單渲染 ──────────────────────────────────────────
    def _refresh_list(self):
        for w in self.sf.winfo_children():
            w.destroy()
        self.row_vars   = []
        self.row_frames = []
        self.selected_idx = None

        if not self.reminders:
            tk.Label(self.sf, text="尚無提醒，請點「➕ 新增」",
                     font=("微軟正黑體", 14), bg=BG_DARK, fg=GRAY
                     ).pack(pady=60)
            return

        for i, rem in enumerate(self.reminders):
            bg  = BG_ROW if i % 2 == 0 else BG_CARD
            row = tk.Frame(self.sf, bg=bg, pady=3)
            row.pack(fill="x")
            self.row_frames.append((row, bg))

            is_on   = rem.get("enabled", True)
            tog_btn = tk.Label(row, text="🟢" if is_on else "⭕",
                               font=("Segoe UI Emoji", 14),
                               bg=bg, cursor="hand2")
            tog_btn.pack(side="left", padx=6)

            def make_toggle(btn, idx):
                def toggle(e):
                    self.reminders[idx]["enabled"] = not self.reminders[idx].get("enabled", True)
                    btn.configure(text="🟢" if self.reminders[idx]["enabled"] else "⭕")
                    save_reminders(self.reminders)
                btn.bind("<Button-1>", toggle)
            make_toggle(tog_btn, i)

            tk.Label(row, text=rem["name"], font=("微軟正黑體", 12, "bold"),
                     bg=bg, fg=WHITE, width=14, anchor="w").pack(side="left")
            tk.Label(row, text=rec_desc(rem), font=("微軟正黑體", 11),
                     bg=bg, fg=GOLD, width=22, anchor="w").pack(side="left")
            tk.Label(row, text=f"-{rem['pre_min']}分",
                     font=("微軟正黑體", 11), bg=bg, fg=GRAY,
                     width=6, anchor="w").pack(side="left")
            stxt, sclr = self._status(rem)
            tk.Label(row, text=stxt, font=("微軟正黑體", 11),
                     bg=bg, fg=sclr, width=9, anchor="w").pack(side="left")
            cd = tk.StringVar(value=self._countdown(rem))
            self.row_vars.append((i, cd))
            tk.Label(row, textvariable=cd, font=("微軟正黑體", 11),
                     bg=bg, fg=GREEN, width=15, anchor="w").pack(side="left")

            def bind_click(widget, idx=i):
                if widget is not tog_btn:
                    widget.bind("<Button-1>", lambda e, ix=idx: self._select(ix))
                for child in widget.winfo_children():
                    bind_click(child, idx)
            bind_click(row)

    def _status(self, rem):
        if not rem.get("enabled", True):
            return "⏸ 停用", GRAY
        rec = rem.get("recurrence", "none")
        if rec != "none":
            return "🔁 週期中", BLUE
        try:
            dt = datetime.strptime(rem["datetime"], "%Y-%m-%d %H:%M")
            if datetime.now() >= dt - timedelta(minutes=rem["pre_min"]):
                return "✅ 已過", GRAY
        except Exception:
            pass
        return "🟢 監控中", GREEN

    def _countdown(self, rem):
        if not rem.get("enabled", True):
            return "—"
        nxt = next_alarm_dt(rem)
        if nxt is None:
            return "—"
        diff = (nxt - datetime.now()).total_seconds()
        if diff <= 0:
            return "已過" if rem.get("recurrence", "none") == "none" else "—"
        h = int(diff // 3600)
        m = int((diff % 3600) // 60)
        s = int(diff % 60)
        return f"{h}時{m:02d}分{s:02d}秒" if h > 0 else f"{m:02d}分{s:02d}秒"

    def _tick(self):
        if self.row_vars:
            for idx, cd in self.row_vars:
                if idx < len(self.reminders):
                    cd.set(self._countdown(self.reminders[idx]))
        self.root.after(1000, self._tick)

    def _select(self, idx):
        self.selected_idx = idx
        for j, (frame, orig_bg) in enumerate(self.row_frames):
            new_bg = "#2a2a60" if j == idx else orig_bg
            frame.configure(bg=new_bg)
            for child in frame.winfo_children():
                if getattr(child, "cget", None) and child.cget("cursor") == "hand2":
                    continue
                try:
                    child.configure(bg=new_bg)
                except tk.TclError:
                    pass
        self.status_var.set(
            f"已選取：{self.reminders[idx]['name']}　← 可按「✏️ 編輯」或「🗑️ 刪除」")

    # ── 操作 ─────────────────────────────────────────────
    def _add(self):
        def on_save(data):
            self.reminders.append(data)
            self._sort()
            save_reminders(self.reminders)
            self._refresh_list()
            self.status_var.set(f"已新增：{data['name']}")
        ReminderDialog(self.root, on_save)

    def _edit(self):
        idx = self.selected_idx
        if idx is None or idx >= len(self.reminders):
            messagebox.showinfo("提示", "請先點選清單中的一筆提醒")
            return
        def on_save(data):
            self.reminders[idx] = data
            self._sort()
            save_reminders(self.reminders)
            self._refresh_list()
            self.status_var.set(f"已更新：{data['name']}")
        ReminderDialog(self.root, on_save, self.reminders[idx])

    def _delete(self):
        idx = self.selected_idx
        if idx is None or idx >= len(self.reminders):
            messagebox.showinfo("提示", "請先點選清單中的一筆提醒")
            return
        name = self.reminders[idx]["name"]
        if messagebox.askyesno("確認刪除", f"確定要刪除「{name}」嗎？"):
            self.reminders.pop(idx)
            self.selected_idx = None
            save_reminders(self.reminders)
            self._refresh_list()
            self.status_var.set(f"已刪除：{name}")

    def _test(self):
        idx  = self.selected_idx
        name = (self.reminders[idx]["name"]
                if idx is not None and idx < len(self.reminders) else "測試會議")
        show_alarm(name, self.root)

    def _shortcut(self):
        try:
            path = create_desktop_shortcut()
            messagebox.showinfo("✅ 完成",
                f"桌面捷徑已建立！\n\n{path}")
            self.status_var.set("✅ 桌面捷徑建立成功")
        except Exception as e:
            messagebox.showerror("建立失敗", f"無法建立捷徑：{e}")

    def _toggle_startup(self):
        enabled = self.startup_var.get()
        try:
            result = set_startup(enabled)
            if enabled:
                # 驗證是否真的建立成功
                if is_startup_enabled():
                    messagebox.showinfo(
                        "✅ 開機自動啟動設定成功",
                        f"工作排程器任務已建立！\n\n任務名稱：{TASK_NAME}\n執行指令：\n{result}\n\n重開機後會自動在系統匣執行。\n\n可在「工作排程器」中搜尋「{TASK_NAME}」確認。")
                    self.status_var.set("✅ 已設定開機自動啟動（工作排程器）")
                else:
                    raise RuntimeError("任務建立後驗證失敗，請以系統管理員身份執行程式再試一次。")
            else:
                self.status_var.set("⏹ 已取消開機自動啟動")
        except Exception as e:
            messagebox.showerror("設定失敗",
                f"錯誤訊息：{e}\n\n請嘗試：\n1. 右鍵點「會議提醒」捷徑 → 以系統管理員身份執行\n2. 重新勾選「開機自動啟動」")
            self.startup_var.set(not enabled)

    def _quit(self):
        if messagebox.askyesno("結束程式", "確定要結束會議提醒程式嗎？"):
            self._force_quit()

    def _force_quit(self):
        self.monitor_running = False
        self.tray.hide()
        self.root.destroy()

    def show_window(self):
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()
        self.tray.hide()

    def _on_close(self):
        self.root.withdraw()
        self.tray.show()

    # ── 監控 ─────────────────────────────────────────────
    def _start_monitor(self):
        def run():
            while self.monitor_running:
                now = datetime.now()
                for rem in self.reminders:
                    if not rem.get("enabled", True):
                        continue
                    try:
                        self._check_fire(rem, now)
                    except Exception:
                        pass
                time.sleep(1)
        threading.Thread(target=run, daemon=True).start()

    def _check_fire(self, rem, now):
        rec = rem.get("recurrence", "none")
        pre = rem.get("pre_min", 0)
        if rec == "none":
            dt = datetime.strptime(rem["datetime"], "%Y-%m-%d %H:%M")
            alarm_dt = dt - timedelta(minutes=pre)
            if alarm_dt <= now < alarm_dt + timedelta(seconds=60):
                last = rem.get("last_fired", "")
                if not last or (now - datetime.strptime(last, "%Y-%m-%d %H:%M")
                                ).total_seconds() > 120:
                    self._fire(rem, now)
        else:
            h, m  = map(int, rem.get("time", "0:0").split(":"))
            total = h * 60 + m - pre
            if total < 0:
                total = 0
            ah, am = divmod(total, 60)
            alarm_today = now.replace(hour=ah, minute=am, second=0, microsecond=0)
            if not (alarm_today <= now < alarm_today + timedelta(seconds=60)):
                return
            wd    = now.weekday()
            match = {"daily": True, "weekdays": wd < 5,
                     "weekly": wd in rem.get("weekdays", []),
                     "monthly": now.day == rem.get("monthday", 1)}.get(rec, False)
            if not match:
                return
            last = rem.get("last_fired", "")
            if last and (now - datetime.strptime(last, "%Y-%m-%d %H:%M")
                         ).total_seconds() < 23 * 3600:
                return
            self._fire(rem, now)

    def _fire(self, rem, now):
        rem["last_fired"] = now.strftime("%Y-%m-%d %H:%M")
        save_reminders(self.reminders)
        name = rem["name"]
        self.root.after(0, self.show_window)
        self.root.after(300, lambda n=name: show_alarm(n, self.root))


# ══════════════════════════════════════════════════════════
#  誇張警報視窗
# ══════════════════════════════════════════════════════════
class AlarmWindow:
    def __init__(self, meeting_name, parent):
        self.meeting_name = meeting_name
        self.parent       = parent
        self.dismissed    = False
        self._create()

    def _create(self):
        self.win = tk.Toplevel(self.parent)
        w  = self.win
        sw, sh = w.winfo_screenwidth(), w.winfo_screenheight()
        w.geometry(f"{sw}x{sh}+0+0")
        w.configure(bg=BG_DARK)
        w.title("🚨 緊急提醒！")
        w.attributes("-topmost", True)
        w.attributes("-fullscreen", True)
        w.overrideredirect(True)
        w.focus_force()
        w.grab_set()
        w.protocol("WM_DELETE_WINDOW", lambda: None)
        self._build(w, sw, sh)
        self._pulse()
        self._shake(w, sw, sh)
        self._force_top(w)

    def _build(self, w, sw, sh):
        self.border = tk.Frame(w, bg=ACCENT)
        self.border.place(relx=0.5, rely=0.5, anchor="center",
                          width=sw - 40, height=sh - 40)
        inner = tk.Frame(self.border, bg=BG_DARK)
        inner.place(x=8, y=8, width=sw - 56, height=sh - 56)
        tk.Label(inner, text="🚨", font=("Segoe UI Emoji", 80),
                 bg=BG_DARK, fg=ACCENT).pack(pady=(50, 0))
        tk.Label(inner, text="⚠️  緊急提醒  ⚠️",
                 font=("微軟正黑體", 52, "bold"), bg=BG_DARK, fg=GOLD).pack(pady=(8, 0))
        self.name_lbl = tk.Label(inner, text=f"【 {self.meeting_name} 】",
                                 font=("微軟正黑體", 72, "bold"), bg=BG_DARK, fg=ACCENT)
        self.name_lbl.pack(pady=(6, 0))
        tk.Label(inner, text="快去開會啦！！！",
                 font=("微軟正黑體", 40, "bold"), bg=BG_DARK, fg=WHITE).pack(pady=(8, 0))
        tk.Label(inner, text=datetime.now().strftime("現在時間：%Y/%m/%d  %H:%M"),
                 font=("微軟正黑體", 22), bg=BG_DARK, fg=GRAY).pack(pady=(14, 0))
        tk.Label(inner, text="按下「我知道了！」才能關閉此視窗",
                 font=("微軟正黑體", 18), bg=BG_DARK, fg=GRAY).pack(pady=(6, 24))
        tk.Button(inner, text="✅   我知道了！馬上去！",
                  font=("微軟正黑體", 28, "bold"), bg=GREEN, fg="#000",
                  relief="flat", padx=50, pady=22, cursor="hand2",
                  command=self._dismiss, activebackground="#00cc66").pack(pady=8)

    def _dismiss(self):
        self.dismissed = True
        self.win.grab_release()
        self.win.destroy()

    def _force_top(self, w):
        if self.dismissed:
            return
        try:
            w.attributes("-topmost", True)
            w.lift()
            w.focus_force()
            hwnd = ctypes.windll.user32.FindWindowW(None, "🚨 緊急提醒！")
            if hwnd:
                ctypes.windll.user32.SetWindowPos(hwnd, -1, 0, 0, 0, 0, 0x0003 | 0x0040)
        except Exception:
            pass
        w.after(300, lambda: self._force_top(w))

    def _pulse(self):
        colors = [ACCENT, ACCENT2, GOLD, ACCENT]
        idx = [0]
        def _p():
            if self.dismissed:
                return
            c = colors[idx[0] % len(colors)]
            self.border.configure(bg=c)
            self.name_lbl.configure(fg=c)
            idx[0] += 1
            self.win.after(300, _p)
        _p()

    def _shake(self, w, sw, sh):
        phase = [0]
        def _s():
            if self.dismissed:
                return
            ox = int(math.sin(phase[0] * 0.7) * 6)
            oy = int(math.cos(phase[0] * 1.1) * 4)
            w.geometry(f"{sw}x{sh}+{max(0,ox)}+{max(0,oy)}")
            phase[0] += 1
            w.after(40, _s)
        _s()


def show_alarm(name, parent):
    AlarmWindow(name, parent)


# ══════════════════════════════════════════════════════════
#  主程式
# ══════════════════════════════════════════════════════════
if __name__ == "__main__":
    # 單一執行個體檢查
    if not acquire_instance_lock():
        show_existing_window()
        sys.exit(0)

    root = tk.Tk()
    app  = MainApp(root)
    if "--minimized" in sys.argv:
        root.after(200, app._on_close)   # 開機啟動 → 直接縮到系統匣
    root.mainloop()
