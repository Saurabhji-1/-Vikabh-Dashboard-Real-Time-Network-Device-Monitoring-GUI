#!/usr/bin/env python3

"""
Vikabh Dashboard - single-file application.

Notes:
- Documents-backed data folder: prefers C:/Users/Administrator/Documents/Vikabh Dashboard,
  falls back to user's Documents, then application folder.
- Attempts to move existing devices.db from application folder into the Documents folder on first run.
- DB_FILE, LOG_FILE, MONITOR_FILE module-level constants point to Documents folder (or fallback).
- Robust write_log() with fallback to app folder.
- Hides child console windows on Windows when launching external viewers.
- Safe exception handling for DB and file operations.

Added features:
- Ping rate control (1-10 seconds) via a Spinbox in the toolbar that controls polling interval.
- "Ping Now" toolbar button for immediate refresh.
- "Open Data Folder" toolbar button to open the Documents-backed folder.
- "Export on Close" checkbox in Settings (if enabled, export occurs on exit).
- "Launch VNC" toolbar button (new) to launch TightVNC/RealVNC viewer for all selected devices.
"""

import os
import sys
import threading
import time
import subprocess
import socket
import sqlite3
import shutil
import json
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog, colorchooser

# Optional Pillow (PIL) for logo/icon generation
try:
    from PIL import Image, ImageDraw, ImageFont, ImageTk
    HAS_PIL = True
except Exception:
    HAS_PIL = False

APP_NAME = "Vikabh Dashboard"

# -------------------- Utilities (timestamps & logging) --------------------
LOG_LOCK = threading.Lock()

def now_ts():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Module-level placeholders updated after choosing data dir
DATA_DIR = None
DB_FILE = None
LOG_FILE = None
MONITOR_FILE = None

def write_log(line):
    """
    Append a timestamped line to LOG_FILE. If it fails, fallback to app folder log.
    """
    global LOG_FILE
    with LOG_LOCK:
        try:
            if not LOG_FILE:
                raise RuntimeError("LOG_FILE not initialized")
            with open(LOG_FILE, "a", encoding="utf-8", errors="ignore") as f:
                f.write(f"{now_ts()} - {line}\n")
        except Exception:
            try:
                fallback = os.path.join(BASE_DIR, "monitor.log")
                with open(fallback, "a", encoding="utf-8", errors="ignore") as f:
                    f.write(f"{now_ts()} - {line}\n")
                LOG_FILE = fallback
            except Exception:
                # final swallow to avoid crashing the app due to logging issues
                pass

# -------------------- Choose Documents folder & module paths -------------
def _choose_documents_vikabh_dir():
    """
    Choose preferred Documents/Vikabh Dashboard path.

    1) Try C:/Users/Administrator/Documents/Vikabh Dashboard on Windows (preferred).
    2) Otherwise use current user's Documents/Vikabh Dashboard.
    3) As last resort, fall back to the application folder.
    """
    # Candidate 1: Administrator Documents (Windows)
    try:
        if sys.platform.startswith("win"):
            admin_docs = os.path.join("C:/Users/Administrator/Documents", "Vikabh Dashboard")
            try:
                os.makedirs(admin_docs, exist_ok=True)
                test_path = os.path.join(admin_docs, ".vikabh_write_test")
                with open(test_path, "w", encoding="utf-8") as f:
                    f.write("ok")
                os.remove(test_path)
                return admin_docs
            except Exception:
                pass
    except Exception:
        pass

    # Candidate 2: current user's Documents
    try:
        home = os.path.expanduser("~")
        user_docs = os.path.join(home, "Documents", "Vikabh Dashboard")
        os.makedirs(user_docs, exist_ok=True)
        test_path2 = os.path.join(user_docs, ".vikabh_write_test")
        with open(test_path2, "w", encoding="utf-8") as f:
            f.write("ok")
        os.remove(test_path2)
        return user_docs
    except Exception:
        pass

    # Last resort: application folder
    return BASE_DIR

# Initialize data folder constants
DATA_DIR = _choose_documents_vikabh_dir()
DB_FILE = os.path.join(DATA_DIR, "devices.db")
LOG_FILE = os.path.join(DATA_DIR, "monitor.log")
MONITOR_FILE = os.path.join(DATA_DIR, "monitor.txt")

# -------------------- Ensure DB exists and migrate if needed -------------
def ensure_db_and_migrate():
    """
    Ensure Documents-backed DB exists and apply schema migration if needed.
    """
    global LOG_FILE, DB_FILE

    # Try to move existing DB from app folder to DATA_DIR (only if target missing)
    try:
        app_db = os.path.join(BASE_DIR, "devices.db")
        if os.path.exists(app_db) and not os.path.exists(DB_FILE):
            try:
                shutil.move(app_db, DB_FILE)
                write_log(f"Moved existing app DB from {app_db} to {DB_FILE}")
            except Exception:
                # fallback to copy if move fails
                try:
                    shutil.copy2(app_db, DB_FILE)
                    write_log(f"Copied existing app DB from {app_db} to {DB_FILE} (move failed)")
                except Exception:
                    write_log(f"Failed to move or copy existing DB from {app_db} to {DB_FILE}")
    except Exception as e:
        try:
            write_log(f"DB relocation check failed: {e}")
        except Exception:
            pass

    # Ensure DB schema exists at DB_FILE
    try:
        conn = sqlite3.connect(DB_FILE, timeout=5)
        cur = conn.cursor()
        cur.execute("""
        CREATE TABLE IF NOT EXISTS devices (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            device_name TEXT,
            host TEXT,
            method TEXT,
            port INTEGER DEFAULT 0,
            team TEXT DEFAULT '',
            enabled INTEGER DEFAULT 1,
            monitoring INTEGER DEFAULT 1,
            offline_time TEXT DEFAULT '',
            last_offline_time TEXT DEFAULT ''
        )
        """)
        cur.execute("""
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT
        )
        """)
        cur.execute("""
        CREATE TABLE IF NOT EXISTS teams (
            name TEXT PRIMARY KEY
        )
        """)
        # default settings
        cur.execute("INSERT OR IGNORE INTO settings(key, value) VALUES ('interval', '10')")
        cur.execute("INSERT OR IGNORE INTO settings(key, value) VALUES ('timeout', '2')")
        cur.execute("INSERT OR IGNORE INTO settings(key, value) VALUES ('export_on_close', '0')")
        conn.commit()
        conn.close()
    except Exception as e:
        try:
            write_log(f"Failed to create/verify DB schema at {DB_FILE}: {e}")
        except Exception:
            pass

    # Ensure a minimal log file exists at LOG_FILE, otherwise fallback to app folder
    try:
        if not os.path.exists(LOG_FILE):
            os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)
            with open(LOG_FILE, "w", encoding="utf-8") as f:
                f.write(f"{now_ts()} - Log created.\n")
    except Exception:
        try:
            fallback_log = os.path.join(BASE_DIR, "monitor.log")
            if not os.path.exists(fallback_log):
                with open(fallback_log, "w", encoding="utf-8") as f:
                    f.write(f"{now_ts()} - Log created (fallback).\n")
            LOG_FILE = fallback_log
            try:
                write_log("Logging using fallback monitor.log in app folder.")
            except Exception:
                pass
        except Exception:
            pass

# -------------------- Theme / Personalization ---------------------------
DEFAULT_PALETTES = {
    "Light": {
        "bg": "#FFFFFF","fg": "#000000","accent": "#007ACC","accent_text": "#FFFFFF",
        "row_alt": "#F7F9FC","status_online": "#008000","status_offline": "#C62828"
    },
    "Dark": {
        "bg": "#2B2B2B","fg": "#EAEAEA","accent": "#1E88E5","accent_text": "#FFFFFF",
        "row_alt": "#333333","status_online": "#66BB6A","status_offline": "#EF5350"
    },
    "Teal": {
        "bg": "#F0FFFF","fg": "#013220","accent": "#00796B","accent_text": "#FFFFFF",
        "row_alt": "#E0F2F1","status_online": "#00796B","status_offline": "#B71C1C"
    }
}

def get_setting(key, default=None):
    try:
        conn = sqlite3.connect(DB_FILE, timeout=5)
        cur = conn.cursor()
        cur.execute("SELECT value FROM settings WHERE key=?", (key,))
        r = cur.fetchone()
        conn.close()
        return r[0] if r else default
    except Exception:
        return default

def set_setting(key, val):
    try:
        conn = sqlite3.connect(DB_FILE, timeout=5)
        cur = conn.cursor()
        cur.execute("INSERT OR REPLACE INTO settings(key, value) VALUES (?,?)", (key, str(val)))
        conn.commit()
        conn.close()
    except Exception:
        pass

def load_theme():
    raw = get_setting("theme", None)
    if not raw:
        return DEFAULT_PALETTES["Light"]
    try:
        palette = json.loads(raw)
        # ensure required keys
        for k in ("bg","fg","accent","accent_text","row_alt","status_online","status_offline"):
            palette.setdefault(k, DEFAULT_PALETTES["Light"].get(k))
        return palette
    except Exception:
        return DEFAULT_PALETTES["Light"]

def save_theme(palette):
    set_setting("theme", json.dumps(palette))

def apply_theme_to_app(root, palette):
    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass
    try:
        root.configure(bg=palette.get("bg", "#FFFFFF"))
    except Exception:
        pass
    style.configure("TLabel", background=palette.get("bg"), foreground=palette.get("fg"))
    style.configure("TFrame", background=palette.get("bg"))
    style.configure("Treeview", background=palette.get("bg"), fieldbackground=palette.get("bg"), foreground=palette.get("fg"))
    style.configure("TButton", background=palette.get("accent"), foreground=palette.get("accent_text"))
    try:
        style.map("TButton", background=[("active", palette.get("accent"))], foreground=[("active", palette.get("accent_text"))])
    except Exception:
        pass
    style.configure("Treeview.Heading", background=palette.get("accent"), foreground=palette.get("accent_text"))
    try:
        style.map("Treeview", background=[('selected', palette.get("accent"))], foreground=[('selected', palette.get("accent_text"))])
    except Exception:
        pass
    try:
        if hasattr(root, "tree"):
            root.tree.tag_configure("online", foreground=palette.get("status_online"))
            root.tree.tag_configure("offline", foreground=palette.get("status_offline"))
    except Exception:
        pass

# -------------------- Networking helpers -------------------------------
def ping_host(host, timeout=2):
    if sys.platform.startswith("win"):
        cmd = ["ping", "-n", "1", "-w", str(int(timeout * 1000)), host]
    else:
        cmd = ["ping", "-c", "1", "-W", str(int(timeout)), host]
    try:
        t0 = time.time()
        run_kwargs = dict(capture_output=True, text=True, timeout=max(3, int(timeout) + 2))
        if sys.platform.startswith("win"):
            run_kwargs["creationflags"] = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)
            try:
                si = subprocess.STARTUPINFO()
                si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                run_kwargs["startupinfo"] = si
            except Exception:
                pass
        proc = subprocess.run(cmd, **run_kwargs)
        output = (proc.stdout or "") + "\n" + (proc.stderr or "")
        outl = output.lower()

        success = False
        latency = None

        if "reply from" in outl:
            success = True
            for token in outl.replace(",", " ").split():
                if "time=" in token or (token.startswith("time") and "ms" in token):
                    try:
                        val = token.replace("time=", "").replace("ms", "").replace("<", "").strip()
                        latency = int(float(val)); break
                    except Exception:
                        pass
                if token.endswith("ms"):
                    try:
                        latency = int(token.replace("ms", "").strip()); break
                    except Exception:
                        pass

        if not success and ("bytes from" in outl or "1 packets received" in outl or "1 received" in outl or "received, 0% packet loss" in outl):
            success = True
            for part in outl.replace(",", " ").split():
                if "time=" in part:
                    try:
                        val = part.split("time=")[1].replace("ms", "").strip()
                        latency = int(float(val)); break
                    except Exception:
                        pass
                if part.endswith("ms"):
                    try:
                        latency = int(float(part.replace("ms", "").strip())); break
                    except Exception:
                        pass

        fail_keywords = ["timed out", "request timed out", "unreachable", "could not find host", "ttl expired", "destination host unreachable", "general failure", "no route to host"]
        for kw in fail_keywords:
            if kw in outl:
                success = False
                break

        if proc.returncode != 0 and not success:
            success = False

        if success and latency is None:
            latency = int((time.time() - t0) * 1000)

        return bool(success), (latency if latency is not None else None)
    except subprocess.TimeoutExpired:
        return False, None
    except Exception:
        return False, None

def tcp_check(host, port=80, timeout=2):
    try:
        with socket.create_connection((host, int(port)), timeout=timeout):
            return True
    except Exception:
        return False

def detect_remote_services(host, timeout=2):
    services = []
    try:
        if tcp_check(host, 5900, timeout=timeout):
            services.append("VNC")
    except Exception:
        pass
    return ",".join(services)

def find_tightvnc_viewer():
    candidates = [
        "vncviewer", "vncviewer.exe",
        "tightvncviewer", "tightvncviewer.exe",
        "tvnviewer", "tvnviewer.exe",
        "xtightvncviewer", "xtightvncviewer.exe",
        "tigervnc", "tigervnc.exe", "vncviewer.exe"
    ]
    for name in candidates:
        p = shutil.which(name)
        if p:
            return p
    if sys.platform.startswith("win"):
        pf = os.environ.get("ProgramFiles", r"C:\Program Files")
        pf_x86 = os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")
        possible_dirs = [os.path.join(pf, "TightVNC"), os.path.join(pf_x86, "TightVNC"), os.path.join(pf, "RealVNC"), os.path.join(pf_x86, "RealVNC")]
        exe_names = ["tvnviewer.exe", "vncviewer.exe", "tightvncviewer.exe"]
        for d in possible_dirs:
            if not os.path.isdir(d):
                continue
            for exe in exe_names:
                candidate = os.path.join(d, exe)
                if os.path.exists(candidate):
                    return candidate
    return None

# -------------------- Export function ----------------------------------
def export_to_documents():
    """
    Ensure DATA_DIR exists and create monitor.txt (tail of LOG_FILE).
    Returns (db_path, monitor_path) on success or (None, None) on failure.
    """
    try:
        os.makedirs(DATA_DIR, exist_ok=True)
    except Exception:
        pass

    try:
        db_target = DB_FILE
        monitor_target = MONITOR_FILE

        if os.path.exists(LOG_FILE):
            try:
                with open(LOG_FILE, "r", encoding="utf-8", errors="ignore") as f:
                    lines = f.readlines()
                tail = lines[-500:] if len(lines) > 500 else lines
                with open(monitor_target, "w", encoding="utf-8", errors="ignore") as f:
                    f.write("Vikabh Dashboard Monitor Export\n")
                    f.write(f"Export Time: {now_ts()}\n\n")
                    f.writelines(tail)
            except Exception as e:
                write_log(f"Failed writing monitor file {monitor_target}: {e}")
                try:
                    with open(monitor_target, "w", encoding="utf-8") as f:
                        f.write("Vikabh Dashboard Monitor Export\n")
                        f.write(f"Export Time: {now_ts()}\n\nNo logs available.\n")
                except Exception:
                    pass
        else:
            try:
                with open(monitor_target, "w", encoding="utf-8") as f:
                    f.write("Vikabh Dashboard Monitor Export\n")
                    f.write(f"Export Time: {now_ts()}\n\nNo logs available.\n")
            except Exception:
                pass

        write_log(f"Exported DB and monitor to {DATA_DIR}")
        return db_target, monitor_target
    except Exception as e:
        write_log(f"Export failed: {e}")
        return None, None

# -------------------- Monitoring engine --------------------------------
class MonitorEngine(threading.Thread):
    def __init__(self, app):
        super().__init__(daemon=True)
        self.app = app
        self._stop_event = threading.Event()

    def stop(self):
        self._stop_event.set()

    def run(self):
        write_log("Monitor engine started.")
        while not self._stop_event.is_set():
            try:
                self.perform_checks()
            except Exception as e:
                write_log(f"Monitor engine error: {e}")
            try:
                interval = float(get_setting("interval", "10"))
            except Exception:
                interval = 10.0
            slept = 0.0
            while slept < interval and not self._stop_event.is_set():
                time.sleep(0.5)
                slept += 0.5
        write_log("Monitor engine stopped.")

    def perform_checks(self):
        try:
            conn = sqlite3.connect(DB_FILE, timeout=5)
            cur = conn.cursor()
            cur.execute("SELECT id, device_name, host, method, port, monitoring, offline_time, last_offline_time FROM devices WHERE enabled=1")
            rows = cur.fetchall()
            conn.close()
        except Exception:
            rows = []
        for row in rows:
            dev_id, name, host, method, port, monitoring, offline_time, last_offline_time = row
            if not monitoring:
                continue
            timeout = float(get_setting("timeout", "2"))
            status_ok = False
            latency_ms = None
            if method and method.lower().startswith("tcp"):
                p = int(port) if port else 80
                status_ok = tcp_check(host, p, timeout=timeout)
            else:
                status_ok, latency_ms = ping_host(host, timeout=timeout)
            remote_services = detect_remote_services(host, timeout=timeout)
            ts = now_ts()
            self.app.queue_device_update(dev_id, status_ok, latency_ms, ts, remote_services)

# -------------------- Main application UI -------------------------------
class DashboardApp(tk.Tk):
    CHECK_OFF = "☐"
    CHECK_ON  = "☑"

    def __init__(self):
        super().__init__()
        ensure_db_and_migrate()
        self.logo_path, self.ico_path = ensure_app_assets()
        self.title(APP_NAME)
        try:
            if os.path.exists(self.ico_path):
                self.iconbitmap(self.ico_path)
        except Exception:
            pass
        self.geometry("1250x700")
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        # state
        self.checked_ids = set()
        self._device_update_queue = []
        self._queue_lock = threading.Lock()
        self.monitor_engine = MonitorEngine(self)
        self._prev_interval = None

        # ping rate control variable (StringVar for Spinbox)
        self.ping_rate_var = tk.StringVar(value=str(clamp_interval(get_setting("interval", "10"))))

        # UI
        self.create_widgets()
        apply_theme_to_app(self, load_theme())
        export_to_documents()
        self.load_devices_into_tree()
        self.monitor_engine.start()
        self.after(500, self._process_queue)

    def create_widgets(self):
        menubar = tk.Menu(self)
        menu = tk.Menu(menubar, tearoff=0)
        menu.add_command(label="Settings", command=self.show_settings)
        menu.add_command(label="Manage (Device)", command=self.manage_device)
        menu.add_command(label="Personalize", command=self.open_personalize)
        menu.add_separator()
        menu.add_command(label="Export to Documents", command=lambda: (export_to_documents(), messagebox.showinfo("Export", f"Exported to {DATA_DIR}")))
        menu.add_separator()
        menu.add_command(label="Help", command=self.show_help)
        menubar.add_cascade(label="Menu", menu=menu)
        self.config(menu=menubar)

        topbar = ttk.Frame(self)
        topbar.pack(side="top", fill="x", padx=6, pady=6)

        left_toolbar = ttk.Frame(topbar)
        left_toolbar.pack(side="left", anchor="w")

        def tb_button(text, cmd, width=12):
            b = ttk.Button(left_toolbar, text=text, command=cmd, width=width)
            b.pack(side="left", padx=2)
            return b

        tb_button("Add Device", self.add_device)
        tb_button("Edit", self.edit_selected)
        tb_button("Start", self.start_selected)
        tb_button("Stop", self.stop_selected)
        tb_button("Logs", self.show_logs)
        tb_button("Ping Now", self.manual_refresh)
        tb_button("Launch VNC", self.launch_vnc_for_selected)  # NEW: multi-launch VNC button

        self.select_all_state = False
        self.select_all_btn = ttk.Button(left_toolbar, text="Select All", command=self.toggle_select_all, width=11)
        self.select_all_btn.pack(side="left", padx=4)

        # Open Data Folder button
        open_folder_btn = ttk.Button(left_toolbar, text="Open Data Folder", command=self.open_data_folder, width=16)
        open_folder_btn.pack(side="left", padx=4)

        logo_frame = ttk.Frame(topbar)
        logo_frame.pack(side="top", expand=True)
        if HAS_PIL and os.path.exists(self.logo_path):
            try:
                img = Image.open(self.logo_path)
                img = img.resize((180, 48), Image.LANCZOS)
                self.logo_imgtk = ImageTk.PhotoImage(img)
                ttk.Label(logo_frame, image=self.logo_imgtk).pack()
            except Exception:
                ttk.Label(logo_frame, text="Vikabh Dashboard", font=("Helvetica", 16, "bold")).pack()
        else:
            ttk.Label(logo_frame, text="Vikabh Dashboard", font=("Helvetica", 16, "bold")).pack()

        right_frame = ttk.Frame(topbar)
        right_frame.pack(side="right", anchor="e")

        # Ping rate control (1-10 seconds)
        ttk.Label(right_frame, text="Ping rate (s):").pack(side="left", padx=(0,4))
        # Use tk.Spinbox for compatibility
        self.ping_spin = tk.Spinbox(right_frame, from_=1, to=10, width=5, textvariable=self.ping_rate_var, command=self.on_ping_rate_change)
        self.ping_spin.pack(side="left", padx=(0,8))
        # Bind manual entry changes (user typed a value)
        self.ping_rate_var.trace_add("write", lambda *a: self._on_ping_var_write())

        ttk.Label(right_frame, text="Filter Team:").pack(side="left", padx=(0,4))
        self.team_filter_var = tk.StringVar(value="All")
        self.team_cb = ttk.Combobox(right_frame, textvariable=self.team_filter_var, state="readonly", width=20)
        self.team_cb.pack(side="left")
        self.team_cb.bind("<<ComboboxSelected>>", lambda e: self.apply_team_filter())
        self.reload_team_list()

        frame = ttk.Frame(self)
        frame.pack(fill="both", expand=True, padx=8, pady=8)

        columns = ("sr_no", "select", "device_name", "host", "method", "result", "lastcheck", "uptime", "status", "offline_time", "remote", "team")
        self.tree = ttk.Treeview(frame, columns=columns, show="headings", selectmode="extended")
        headings = {
            "sr_no": ("Sr. No.", 60),
            "select": ("Select", 70),
            "device_name": ("Device Name", 200),
            "host": ("Host", 140),
            "method": ("Method", 100),
            "result": ("Result", 120),
            "lastcheck": ("Last Check", 160),
            "uptime": ("Uptime", 80),
            "status": ("Status", 80),
            "offline_time": ("Last Offline Time", 160),
            "remote": ("Remote", 140),
            "team": ("Team", 120),
        }
        for k, (text, w) in headings.items():
            anchor = "center" if k in ("sr_no","select","uptime","status") else "w"
            self.tree.heading(k, text=text)
            self.tree.column(k, width=w, anchor=anchor)
        self.tree.pack(side="left", fill="both", expand=True)

        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")

        theme = load_theme()
        self.tree.tag_configure("online", foreground=theme.get("status_online", "green"))
        self.tree.tag_configure("offline", foreground=theme.get("status_offline", "red"))

        self.tree.bind("<Double-1>", lambda e: self.edit_selected())
        self.tree.bind("<Button-1>", self._on_tree_click)
        self.create_context_menu()

        bottom_frame = ttk.Frame(self)
        bottom_frame.pack(side="bottom", fill="x")
        self.status_var = tk.StringVar(value="Ready.")
        statusbar = ttk.Label(bottom_frame, textvariable=self.status_var, relief="sunken", anchor="w")
        statusbar.pack(side="left", fill="x", expand=True)
        creator = ttk.Label(bottom_frame, text="Saurabh Sharma", font=("Arial", 7))
        creator.pack(side="right", padx=6, pady=2)

    # ----------------- Ping rate helpers ---------------------------------
    def _on_ping_var_write(self):
        # when user types value into spinbox, keep it constrained
        val = self.ping_rate_var.get()
        try:
            ival = int(float(val))
        except Exception:
            return
        if ival < 1: ival = 1
        if ival > 10: ival = 10
        if str(ival) != val:
            # update canonical string
            self.ping_rate_var.set(str(ival))
        # setting writes to DB via on_ping_rate_change when user confirms (spin command) or we call explicitly

    def on_ping_rate_change(self):
        """Called by the Spinbox command when the value changes via arrow buttons"""
        val = self.ping_rate_var.get()
        try:
            ival = int(float(val))
        except Exception:
            messagebox.showerror("Invalid", "Ping rate must be an integer between 1 and 10.")
            return
        if ival < 1 or ival > 10:
            messagebox.showerror("Invalid", "Ping rate must be between 1 and 10 seconds.")
            self.ping_rate_var.set(str(clamp_interval(get_setting("interval", "10"))))
            return
        # store into same 'interval' setting so MonitorEngine uses it directly
        set_setting("interval", str(ival))
        write_log(f"Ping rate set to {ival}s via toolbar control.")
        self.status_var.set(f"Ping rate set to {ival}s.")

    def update_ping_rate_widget(self):
        """Update spinbox value to match the interval setting (used when code changes interval)"""
        try:
            val = clamp_interval(get_setting("interval", "10"))
            if str(val) != self.ping_rate_var.get():
                self.ping_rate_var.set(str(val))
        except Exception:
            pass

    # ---------------- Team functions -------------------------------------
    def reload_team_list(self):
        try:
            conn = sqlite3.connect(DB_FILE, timeout=5)
            cur = conn.cursor()
            cur.execute("SELECT name FROM teams ORDER BY name COLLATE NOCASE")
            rows = cur.fetchall()
            conn.close()
        except Exception:
            rows = []
        teams = ["All"] + [r[0] for r in rows]
        self.team_cb['values'] = teams
        if self.team_filter_var.get() not in teams:
            self.team_filter_var.set("All")

    def apply_team_filter(self):
        self.load_devices_into_tree()

    # ---------------- Device CRUD ---------------------------------------
    def load_devices_into_tree(self):
        for r in self.tree.get_children():
            self.tree.delete(r)
        team_filter = self.team_filter_var.get()
        try:
            conn = sqlite3.connect(DB_FILE, timeout=5)
            cur = conn.cursor()
            if team_filter and team_filter != "All":
                cur.execute("SELECT id, device_name, host, method, port, monitoring, offline_time, last_offline_time, team FROM devices WHERE enabled=1 AND (team=? OR team='')", (team_filter,))
            else:
                cur.execute("SELECT id, device_name, host, method, port, monitoring, offline_time, last_offline_time, team FROM devices WHERE enabled=1")
            rows = cur.fetchall()
            conn.close()
        except Exception:
            rows = []
        for idx, row in enumerate(rows, start=1):
            dev_id, name, host, method, port, monitoring, offline_time, last_offline_time, team = row
            method_display = (method or "Ping") + (f":{port}" if port else "")
            lastcheck = ""
            result = ""
            uptime = "100%"
            status_txt = "Online" if monitoring and (result == "OK") else ("Stopped" if not monitoring else "Unknown")
            tag = "offline"
            chk = self.CHECK_ON if dev_id in self.checked_ids else self.CHECK_OFF
            display_offline = last_offline_time or offline_time or ""
            remote_display = ""
            values = (str(idx), chk, name or "", host or "", method_display, result, lastcheck, uptime, status_txt, display_offline, remote_display, team or "")
            iid = str(dev_id)
            self.tree.insert("", "end", iid=iid, values=values, tags=(tag,))
        self._update_select_all_state()

    def add_device(self):
        dlg = DeviceDialog(self, title="Add Device")
        self.wait_window(dlg.top)
        if dlg.result:
            name, host, method, port, team = dlg.result
            try:
                conn = sqlite3.connect(DB_FILE, timeout=5)
                cur = conn.cursor()
                cur.execute(
                    "INSERT INTO devices(device_name, host, method, port, team, enabled, monitoring, offline_time, last_offline_time) VALUES(?,?,?,?,?,1,1,?,?)",
                    (name, host, method, port, team, "", "")
                )
                conn.commit()
                new_id = cur.lastrowid
                conn.close()
                write_log(f"Device added: {name} ({host}) team={team} method={method} port={port}")
                self.checked_ids.add(new_id)
                self.reload_team_list()
                self.load_devices_into_tree()
                self.status_var.set("Device added.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to add device: {e}")

    def edit_selected(self):
        ids = self.get_selected_ids()
        if not ids:
            messagebox.showinfo("Edit", "Select a device to edit (checkbox or row).")
            return
        dev_id = ids[0]
        try:
            conn = sqlite3.connect(DB_FILE, timeout=5)
            cur = conn.cursor()
            cur.execute("SELECT device_name, host, method, port, team FROM devices WHERE id=?", (dev_id,))
            r = cur.fetchone()
            conn.close()
        except Exception:
            r = None
        if not r:
            messagebox.showerror("Edit", "Device not found.")
            return
        name, host, method, port, team = r
        dlg = DeviceDialog(self, title="Edit Device", initial=(name, host, method, port, team))
        self.wait_window(dlg.top)
        if dlg.result:
            nd, nh, nm, np, nt = dlg.result
            try:
                conn = sqlite3.connect(DB_FILE, timeout=5)
                cur = conn.cursor()
                cur.execute("UPDATE devices SET device_name=?, host=?, method=?, port=?, team=? WHERE id=?", (nd, nh, nm, np, nt, dev_id))
                conn.commit()
                conn.close()
                write_log(f"Device edited: {nd} ({nh}) team={nt} method={nm} port={np}")
                self.load_devices_into_tree()
                self.status_var.set("Device edited.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to update device: {e}")

    def remove_selected(self):
        ids = self.get_selected_ids()
        if not ids:
            messagebox.showinfo("Remove", "Select device(s) to remove.")
            return
        if not messagebox.askyesno("Remove", f"Remove {len(ids)} device(s)?"):
            return
        try:
            conn = sqlite3.connect(DB_FILE, timeout=5)
            cur = conn.cursor()
            cur.executemany("DELETE FROM devices WHERE id=?", [(i,) for i in ids])
            conn.commit()
            conn.close()
            for i in ids:
                self.checked_ids.discard(i)
            write_log(f"Removed devices: {ids}")
            self.load_devices_into_tree()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to remove devices: {e}")

    def get_selected_ids(self):
        if self.checked_ids:
            return sorted(list(self.checked_ids))
        s = self.tree.selection()
        return [int(iid) for iid in s] if s else []

    # --------------- Context menu / Manage / Start / Stop ----------------
    def create_context_menu(self):
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label="Edit", command=self.edit_selected)
        menu.add_command(label="Start Monitoring", command=self.start_selected)
        menu.add_command(label="Stop Monitoring", command=self.stop_selected)
        menu.add_separator()
        menu.add_command(label="Launch TightVNC", command=self.launch_tightvnc_selected)
        menu.add_separator()
        menu.add_command(label="Remove", command=self.remove_selected)
        self.tree.bind("<Button-3>", lambda e: self._show_context_menu(e, menu))

    def _show_context_menu(self, event, menu):
        iid = self.tree.identify_row(event.y)
        if iid:
            self.tree.selection_set(iid)
            menu.post(event.x_root, event.y_root)

    def manage_device(self):
        ids = self.get_selected_ids()
        if not ids:
            messagebox.showinfo("Manage", "Select a device to manage.")
            return
        ManageDialog(self, ids[0])

    def start_selected(self):
        ids = self.get_selected_ids()
        if not ids:
            messagebox.showinfo("Start", "Select device(s) to start monitoring.")
            return
        if self._prev_interval is None:
            try:
                cur_iv = float(get_setting("interval", "10"))
            except:
                cur_iv = 10.0
            self._prev_interval = cur_iv
            set_setting("interval", "1")
            write_log(f"Fast-refresh enabled (1s). Previous interval saved: {self._prev_interval}")
            self.status_var.set("Fast refresh enabled (1s).")
            # update ping rate UI
            self.update_ping_rate_widget()
        try:
            conn = sqlite3.connect(DB_FILE, timeout=5)
            cur = conn.cursor()
            cur.executemany("UPDATE devices SET monitoring=1 WHERE id=?", [(i,) for i in ids])
            conn.commit()
            conn.close()
            write_log(f"Started monitoring for: {ids}")
            self.load_devices_into_tree()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to start monitoring: {e}")

    def stop_selected(self):
        ids = self.get_selected_ids()
        if not ids:
            messagebox.showinfo("Stop", "Select device(s) to stop monitoring.")
            return
        try:
            conn = sqlite3.connect(DB_FILE, timeout=5)
            cur = conn.cursor()
            cur.executemany("UPDATE devices SET monitoring=0 WHERE id=?", [(i,) for i in ids])
            conn.commit()
            conn.close()
            write_log(f"Stopped monitoring for: {ids}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to stop monitoring: {e}")
            return
        if self._prev_interval is not None:
            try:
                conn = sqlite3.connect(DB_FILE, timeout=5)
                cur = conn.cursor()
                cur.execute("SELECT COUNT(*) FROM devices WHERE monitoring=1 AND enabled=1")
                count = cur.fetchone()[0]
                conn.close()
            except Exception:
                count = 0
            if count == 0:
                set_setting("interval", str(self._prev_interval))
                write_log(f"Fast-refresh disabled. Restored interval: {self._prev_interval}")
                self._prev_interval = None
                self.status_var.set(f"Restored polling interval to {get_setting('interval')}.")
                # update ping rate UI
                self.update_ping_rate_widget()
            else:
                self.status_var.set(f"Stopped selected devices. Fast-refresh remains (other devices monitored).")
        else:
            self.status_var.set("Stopped selected devices.")
        self.load_devices_into_tree()

    # --------------- Queue processing / UI updates -----------------------
    def manual_refresh(self):
        threading.Thread(target=self.monitor_engine.perform_checks, daemon=True).start()
        self.status_var.set("Manual refresh started.")
        write_log("Manual refresh started.")

    def queue_device_update(self, dev_id, status_ok, latency_ms, ts, remote_services):
        with self._queue_lock:
            self._device_update_queue.append((dev_id, status_ok, latency_ms, ts, remote_services))

    def _process_queue(self):
        with self._queue_lock:
            q = list(self._device_update_queue)
            self._device_update_queue.clear()
        updated = False
        for dev_id, status_ok, latency_ms, ts, remote_services in q:
            self._apply_update_to_tree(dev_id, status_ok, latency_ms, ts, remote_services)
            updated = True
        if updated:
            self.status_var.set(f"Last update: {now_ts()}")
        self.after(800, self._process_queue)

    def _apply_update_to_tree(self, dev_id, status_ok, latency_ms, ts, remote_services):
        iid = str(dev_id)
        if not self.tree.exists(iid):
            self.load_devices_into_tree()
            return
        result = f"OK ({latency_ms} ms)" if status_ok and latency_ms is not None else ("OK" if status_ok else "DOWN")
        lastcheck = ts
        uptime = "100%" if status_ok else "0%"
        status_text = "Online" if status_ok else "Offline"
        tag = "online" if status_ok else "offline"
        try:
            conn = sqlite3.connect(DB_FILE, timeout=5)
            cur = conn.cursor()
            cur.execute("SELECT offline_time, last_offline_time, monitoring FROM devices WHERE id=?", (dev_id,))
            r = cur.fetchone()
            if r:
                current_offline_time, current_last_offline, monitoring = r
            else:
                current_offline_time, current_last_offline, monitoring = ("", "", 1)
            if not status_ok:
                if not current_offline_time:
                    cur.execute("UPDATE devices SET offline_time=?, last_offline_time=? WHERE id=?", (ts, ts, dev_id)); conn.commit()
            else:
                if current_offline_time:
                    cur.execute("UPDATE devices SET offline_time=? WHERE id=?", ("", dev_id)); conn.commit()
            cur.execute("SELECT last_offline_time FROM devices WHERE id=?", (dev_id,))
            rr = cur.fetchone()
            display_offline = (rr[0] if rr and rr[0] else "")
            conn.close()
        except Exception:
            display_offline = ""
        try:
            self.tree.set(iid, 'result', result)
            self.tree.set(iid, 'lastcheck', lastcheck)
            self.tree.set(iid, 'uptime', uptime)
            self.tree.set(iid, 'status', status_text)
            self.tree.set(iid, 'offline_time', display_offline)
            self.tree.set(iid, 'remote', remote_services or "")
            self.tree.item(iid, tags=(tag,))
        except Exception:
            existing_vals = list(self.tree.item(iid, "values"))
            sr_no = existing_vals[0] if len(existing_vals) >= 1 else ""
            chk = existing_vals[1] if len(existing_vals) >= 2 else (self.CHECK_ON if dev_id in self.checked_ids else self.CHECK_OFF)
            device_name = existing_vals[2] if len(existing_vals) >= 3 else ""
            host = existing_vals[3] if len(existing_vals) >= 4 else ""
            method_display = existing_vals[4] if len(existing_vals) >= 5 else ""
            team_val = existing_vals[-1] if len(existing_vals) >= 1 else ""
            new_values = (sr_no, chk, device_name, host, method_display, result, lastcheck, uptime, status_text, display_offline, remote_services or "", team_val)
            self.tree.item(iid, values=new_values, tags=(tag,))

    # ---------------- Logs / Settings / Help -----------------------------
    def show_logs(self):
        try:
            with open(LOG_FILE, "r", encoding="utf-8", errors="ignore") as f:
                data = f.read()
        except Exception:
            try:
                with open(os.path.join(BASE_DIR, "monitor.log"), "r", encoding="utf-8", errors="ignore") as f:
                    data = f.read()
            except Exception:
                data = ""
        top = tk.Toplevel(self)
        top.title("Logs")
        top.geometry("800x500")
        txt = tk.Text(top, wrap="none")
        txt.insert("1.0", data)
        txt.configure(state="disabled")
        txt.pack(fill="both", expand=True)
        def save():
            path = filedialog.asksaveasfilename(defaultextension=".log", filetypes=[("Log files","*.log"), ("All","*.*")])
            if path:
                with open(path, "w", encoding="utf-8") as f:
                    f.write(data)
                messagebox.showinfo("Saved", f"Logs saved to {path}")
        btn = ttk.Button(top, text="Save As...", command=save)
        btn.pack(side="bottom", pady=4)

    def show_settings(self):
        SettingsDialog(self)

    def show_help(self):
        msg = (
            f"{APP_NAME} Help\n\n"
            "- Use Menu (top-left) for Settings, Manage, Personalize, Export and Help.\n"
            "- Toolbar contains primary actions only (Add Device, Edit, Start, Stop, Logs, Ping Now, Open Data Folder, Launch VNC).\n"
            "- Add teams from the Add/Edit Device dialog using 'New Team'.\n"
            f"- Exported DB and monitor.txt are available in: {DATA_DIR}\n"
            "- Use 'Ping rate (s)' to set interval between polls (1-10 seconds).\n"
        )
        messagebox.showinfo("Help", msg)

    # ---------------- Treeview checkbox handling -------------------------
    def _on_tree_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        col = self.tree.identify_column(event.x)
        row = self.tree.identify_row(event.y)
        if not row:
            return
        if col == "#2":
            try:
                dev_id = int(row)
            except Exception:
                return "break"
            if dev_id in self.checked_ids:
                self.checked_ids.remove(dev_id)
            else:
                self.checked_ids.add(dev_id)
            curvals = list(self.tree.item(row, "values"))
            if len(curvals) >= 2:
                curvals[1] = self.CHECK_ON if dev_id in self.checked_ids else self.CHECK_OFF
                self.tree.item(row, values=curvals)
            self._update_select_all_state()
            return "break"
        return

    def toggle_select_all(self):
        all_iids = self.tree.get_children()
        all_ids = [int(iid) for iid in all_iids]
        if not all_ids:
            return
        select_all = not self.select_all_state
        if select_all:
            for i in all_ids:
                self.checked_ids.add(i)
        else:
            for i in all_ids:
                self.checked_ids.discard(i)
        for iid in all_iids:
            curvals = list(self.tree.item(iid, "values"))
            dev_id = int(iid)
            if len(curvals) >= 2:
                curvals[1] = self.CHECK_ON if dev_id in self.checked_ids else self.CHECK_OFF
                self.tree.item(iid, values=curvals)
        self.select_all_state = select_all
        self.select_all_btn.config(text="Unselect All" if select_all else "Select All")

    def _update_select_all_state(self):
        all_iids = self.tree.get_children()
        if not all_iids:
            self.select_all_state = False
            self.select_all_btn.config(text="Select All")
            return
        all_ids = [int(iid) for iid in all_iids]
        if all(dev_id in self.checked_ids for dev_id in all_ids) and all_ids:
            self.select_all_state = True
            self.select_all_btn.config(text="Unselect All")
        else:
            self.select_all_state = False
            self.select_all_btn.config(text="Select All")

    # ---------------- Launch VNC viewer without console ------------------
    def launch_tightvnc_selected(self):
        ids = self.get_selected_ids()
        if not ids:
            messagebox.showinfo("Launch TightVNC", "Select a device (checkbox or row) to launch TightVNC.")
            return
        dev_id = ids[0]
        try:
            conn = sqlite3.connect(DB_FILE, timeout=5)
            cur = conn.cursor()
            cur.execute("SELECT host FROM devices WHERE id=?", (dev_id,))
            r = cur.fetchone()
            conn.close()
        except Exception:
            r = None
        if not r:
            messagebox.showerror("Launch TightVNC", "Device not found.")
            return
        host = r[0]
        viewer = find_tightvnc_viewer()
        if not viewer:
            messagebox.showerror("Launch TightVNC", "No VNC/TightVNC viewer detected. Install one or place it on PATH.")
            return
        try:
            if sys.platform.startswith("win"):
                subprocess.Popen([viewer, host], creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000))
            else:
                subprocess.Popen([viewer, host])
            write_log(f"Launched VNC viewer for {host}")
            self.status_var.set(f"Launched VNC viewer for {host}")
        except Exception as e:
            messagebox.showerror("Launch TightVNC", f"Failed to launch VNC viewer: {e}")
            write_log(f"Failed to launch VNC viewer {viewer} for {host}: {e}")

    def launch_vnc_for_selected(self):
        """
        Launch the detected VNC viewer for all selected devices (checkbox selection preferred).
        """
        ids = self.get_selected_ids()
        if not ids:
            messagebox.showinfo("Launch VNC", "Select one or more devices (via checkbox or row selection) to launch VNC viewers.")
            return

        viewer = find_tightvnc_viewer()
        if not viewer:
            messagebox.showerror("Launch VNC", "No VNC/TightVNC viewer detected on PATH or standard install locations.")
            return

        # Resolve hosts from DB
        hosts = []
        failed_ids = []
        try:
            conn = sqlite3.connect(DB_FILE, timeout=5)
            cur = conn.cursor()
            cur.execute("SELECT id, host FROM devices WHERE id IN ({seq})".format(seq=",".join(["?"]*len(ids))), ids)
            rows = cur.fetchall()
            conn.close()
            id_to_host = {r[0]: r[1] for r in rows}
            for i in ids:
                h = id_to_host.get(i)
                if h:
                    hosts.append((i, h))
                else:
                    failed_ids.append(i)
        except Exception as e:
            messagebox.showerror("Launch VNC", f"Failed to read device hosts: {e}")
            write_log(f"Failed to read hosts for VNC launch: {e}")
            return

        if not hosts:
            messagebox.showerror("Launch VNC", "No hosts could be resolved for the selected devices.")
            return

        successes = []
        failures = []
        for dev_id, host in hosts:
            try:
                if sys.platform.startswith("win"):
                    subprocess.Popen([viewer, host], creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000))
                else:
                    subprocess.Popen([viewer, host])
                successes.append((dev_id, host))
                write_log(f"Launched VNC viewer for device {dev_id} ({host}) using {viewer}")
            except Exception as e:
                failures.append((dev_id, host, str(e)))
                write_log(f"Failed to launch VNC for device {dev_id} ({host}): {e}")

        # Build a summary message
        parts = []
        if successes:
            parts.append(f"Launched VNC for {len(successes)} device(s).")
        if failures or failed_ids:
            parts.append(f"Failed to launch for {len(failures) + len(failed_ids)} device(s).")
        summary = " ".join(parts)
        if not summary:
            summary = "No action performed."
        messagebox.showinfo("Launch VNC", summary)
        # update status bar
        self.status_var.set(summary)

    # ---------------- Personalize panel (Menu) ---------------------------
    def open_personalize(self):
        dlg = PersonalizeDialog(self, initial_palette=load_theme())
        self.wait_window(dlg.top)
        if getattr(dlg, "result", None):
            apply_theme_to_app(self, dlg.result)
            save_theme(dlg.result)
            self.load_devices_into_tree()

    # ---------------- Misc ----------------------------------------------
    def on_close(self):
        if messagebox.askokcancel("Quit", "Quit application?"):
            write_log("Application closing.")
            try:
                # if user enabled export on close, do export
                if get_setting("export_on_close", "0") == "1":
                    export_to_documents()
                    write_log("Auto-export on close executed.")
            except Exception:
                pass
            try:
                self.monitor_engine.stop()
            except Exception:
                pass
            self.destroy()

    def open_data_folder(self):
        """Open DATA_DIR in file explorer."""
        try:
            if not os.path.exists(DATA_DIR):
                os.makedirs(DATA_DIR, exist_ok=True)
            if sys.platform.startswith("win"):
                os.startfile(DATA_DIR)
            elif sys.platform.startswith("darwin"):
                subprocess.Popen(["open", DATA_DIR])
            else:
                subprocess.Popen(["xdg-open", DATA_DIR])
            write_log(f"Opened data folder: {DATA_DIR}")
        except Exception as e:
            messagebox.showerror("Open Folder", f"Failed to open data folder: {e}")
            write_log(f"Failed to open data folder {DATA_DIR}: {e}")

# -------------------- Dialogs & helpers --------------------------------
class DeviceDialog:
    """
    Add/Edit device dialog. Includes 'New Team' button to add team inline.
    """
    def __init__(self, parent, title="Device", initial=None):
        self.parent = parent
        self.top = tk.Toplevel(parent)
        self.top.title(title)
        self.top.transient(parent)
        self.top.grab_set()
        self.result = None

        ttk.Label(self.top, text="Device Name:").grid(row=0, column=0, sticky="e", padx=4, pady=4)
        self.e_name = ttk.Entry(self.top, width=40); self.e_name.grid(row=0, column=1, padx=4, pady=4)

        ttk.Label(self.top, text="Host (IP/hostname):").grid(row=1, column=0, sticky="e", padx=4, pady=4)
        self.e_host = ttk.Entry(self.top, width=40); self.e_host.grid(row=1, column=1, padx=4, pady=4)

        ttk.Label(self.top, text="Method:").grid(row=2, column=0, sticky="e", padx=4, pady=4)
        self.method_var = tk.StringVar(value="Ping")
        self.combo_method = ttk.Combobox(self.top, textvariable=self.method_var, values=["Ping", "TCP"], state="readonly", width=38)
        self.combo_method.grid(row=2, column=1, padx=4, pady=4)

        ttk.Label(self.top, text="Port (for TCP):").grid(row=3, column=0, sticky="e", padx=4, pady=4)
        self.e_port = ttk.Entry(self.top, width=20); self.e_port.grid(row=3, column=1, padx=4, pady=4, sticky="w")

        ttk.Label(self.top, text="Team:").grid(row=4, column=0, sticky="e", padx=4, pady=4)
        self.team_var = tk.StringVar()
        self.team_cb = ttk.Combobox(self.top, textvariable=self.team_var, state="readonly", width=30)
        self.team_cb.grid(row=4, column=1, padx=4, pady=4, sticky="w")
        ttk.Button(self.top, text="New Team", width=10, command=self.add_team_inline).grid(row=4, column=2, padx=4, pady=4)
        self.reload_team_list_into_cb()

        btn_frame = ttk.Frame(self.top); btn_frame.grid(row=5, column=0, columnspan=3, pady=8)
        ttk.Button(btn_frame, text="OK", command=self.on_ok).pack(side="left", padx=6)
        ttk.Button(btn_frame, text="Cancel", command=self.on_cancel).pack(side="left", padx=6)

        if initial:
            name, host, method, port, team = initial
            if name: self.e_name.insert(0, name)
            if host: self.e_host.insert(0, host)
            if method: self.method_var.set(method if method else "Ping")
            if port: self.e_port.insert(0, str(port))
            if team:
                # reload list and set team if present
                self.reload_team_list_into_cb()
                self.team_var.set(team)

    def reload_team_list_into_cb(self):
        try:
            conn = sqlite3.connect(DB_FILE, timeout=5)
            cur = conn.cursor()
            cur.execute("SELECT name FROM teams ORDER BY name COLLATE NOCASE")
            rows = cur.fetchall()
            conn.close()
        except Exception:
            rows = []
        teams = [""] + [r[0] for r in rows]
        self.team_cb['values'] = teams

    def add_team_inline(self):
        name = simpledialog.askstring("New Team", "Enter team/department name:", parent=self.top)
        if not name:
            return
        name = name.strip()
        if not name:
            return
        try:
            conn = sqlite3.connect(DB_FILE, timeout=5); cur = conn.cursor()
            cur.execute("INSERT INTO teams(name) VALUES(?)", (name,)); conn.commit()
            conn.close()
            write_log(f"Team added (inline): {name}")
        except sqlite3.IntegrityError:
            messagebox.showinfo("New Team", "Team already exists.")
        except Exception as e:
            messagebox.showerror("New Team", f"Failed to add team: {e}")
        self.reload_team_list_into_cb()
        self.team_var.set(name)
        try:
            self.top.master.reload_team_list()
        except Exception:
            pass

    def on_ok(self):
        name = self.e_name.get().strip()
        host = self.e_host.get().strip()
        method = self.method_var.get()
        port = self.e_port.get().strip()
        team = self.team_var.get().strip()
        if not name or not host:
            messagebox.showerror("Validation", "Device Name and Host are required.")
            return
        try:
            port_val = int(port) if port else 0
        except:
            messagebox.showerror("Validation", "Port must be an integer.")
            return
        self.result = (name, host, method, port_val, team)
        self.top.destroy()

    def on_cancel(self):
        self.top.destroy()

class ManageDialog:
    def __init__(self, parent, dev_id):
        self.top = tk.Toplevel(parent)
        self.top.title("Manage device")
        self.top.geometry("500x300")
        ttk.Label(self.top, text=f"Manage device ID: {dev_id}").pack(pady=8)
        ttk.Label(self.top, text="(Placeholder)").pack(pady=6)
        ttk.Button(self.top, text="Close", command=self.top.destroy).pack(pady=12)

class SettingsDialog:
    def __init__(self, parent):
        self.top = tk.Toplevel(parent)
        self.top.title("Settings")
        self.top.transient(parent)
        self.top.grab_set()
        ttk.Label(self.top, text="Polling interval (sec):").grid(row=0, column=0, sticky="e", padx=4, pady=6)
        self.e_interval = ttk.Entry(self.top, width=40); self.e_interval.grid(row=0, column=1, padx=4, pady=6)
        self.e_interval.insert(0, get_setting("interval", "10"))
        ttk.Label(self.top, text="Timeout (sec):").grid(row=1, column=0, sticky="e", padx=4, pady=6)
        self.e_timeout = ttk.Entry(self.top, width=40); self.e_timeout.grid(row=1, column=1, padx=4, pady=6)
        self.e_timeout.insert(0, get_setting("timeout", "2"))

        # Export on close option
        self.export_on_close_var = tk.IntVar(value=1 if get_setting("export_on_close", "0") == "1" else 0)
        ttk.Checkbutton(self.top, text="Export DB & monitor.txt on exit", variable=self.export_on_close_var).grid(row=2, column=0, columnspan=2, padx=4, pady=6, sticky="w")

        btn_frame = ttk.Frame(self.top); btn_frame.grid(row=3, column=0, columnspan=2, pady=8)
        ttk.Button(btn_frame, text="Save", command=self.save).pack(side="left", padx=6)
        ttk.Button(btn_frame, text="Cancel", command=self.top.destroy).pack(side="left", padx=6)

    def save(self):
        iv = self.e_interval.get().strip(); to = self.e_timeout.get().strip()
        try:
            fiv = float(iv); fto = float(to)
        except:
            messagebox.showerror("Validation", "Interval and timeout must be numbers."); return
        # clamp interval to sensible range (1-3600)
        if fiv < 1:
            fiv = 1
        if fiv > 3600:
            fiv = 3600
        set_setting("interval", str(int(fiv)))
        set_setting("timeout", str(fto))
        set_setting("export_on_close", "1" if self.export_on_close_var.get() == 1 else "0")
        write_log(f"Settings changed: interval={int(fiv)}, timeout={fto}, export_on_close={get_setting('export_on_close')}")
        # If top-level has ping rate widget, update it
        try:
            self.top.master.update_ping_rate_widget()
        except Exception:
            pass
        self.top.destroy()

class PersonalizeDialog:
    def __init__(self, parent, initial_palette=None):
        self.parent = parent
        self.top = tk.Toplevel(parent)
        self.top.title("Personalize")
        self.top.transient(parent)
        self.top.grab_set()
        self.result = None

        self.initial = initial_palette or load_theme()
        self.current = dict(self.initial)

        ttk.Label(self.top, text="Choose a palette:").grid(row=0, column=0, padx=8, pady=6, sticky="w")
        self.listbox = tk.Listbox(self.top, height=5, exportselection=False)
        for k in DEFAULT_PALETTES.keys():
            self.listbox.insert("end", k)
        self.listbox.grid(row=1, column=0, rowspan=4, padx=8, pady=4, sticky="ns")
        self.listbox.bind("<<ListboxSelect>>", self.on_palette_select)

        preview_frame = ttk.Frame(self.top, padding=8)
        preview_frame.grid(row=0, column=1, rowspan=5, padx=8, pady=4, sticky="nsew")
        preview_frame.columnconfigure(0, weight=1)

        ttk.Label(preview_frame, text="Accent color:").grid(row=0, column=0, sticky="w")
        self.btn_accent = ttk.Button(preview_frame, text="Pick...", command=lambda: self.choose_color("accent"))
        self.btn_accent.grid(row=0, column=1, sticky="e")

        ttk.Label(preview_frame, text="Background color:").grid(row=1, column=0, sticky="w")
        self.btn_bg = ttk.Button(preview_frame, text="Pick...", command=lambda: self.choose_color("bg"))
        self.btn_bg.grid(row=1, column=1, sticky="e")

        ttk.Label(preview_frame, text="Text color:").grid(row=2, column=0, sticky="w")
        self.btn_fg = ttk.Button(preview_frame, text="Pick...", command=lambda: self.choose_color("fg"))
        self.btn_fg.grid(row=2, column=1, sticky="e")

        # Use plain tk.Label for preview to allow bg/fg changes reliably
        self.preview = tk.Label(preview_frame, text="Preview Area", anchor="center", padx=6, pady=6)
        self.preview.grid(row=3, column=0, columnspan=2, pady=(12,6), sticky="ew")

        btn_frame = ttk.Frame(self.top)
        btn_frame.grid(row=5, column=0, columnspan=2, pady=8)
        ttk.Button(btn_frame, text="Save", command=self.on_save).pack(side="left", padx=6)
        ttk.Button(btn_frame, text="Cancel", command=self.on_cancel).pack(side="left", padx=6)
        ttk.Button(btn_frame, text="Reset", command=self.on_reset).pack(side="left", padx=6)

        # Initialize selection and preview
        self.listbox.selection_clear(0, 'end')
        for i, k in enumerate(DEFAULT_PALETTES.keys()):
            p = DEFAULT_PALETTES.get(k)
            if p and p.get("bg") == self.current.get("bg"):
                self.listbox.selection_set(i)
                break
        self.update_preview()

    def on_palette_select(self, evt=None):
        sel = self.listbox.curselection()
        if not sel:
            return
        key = list(DEFAULT_PALETTES.keys())[sel[0]]
        self.current = dict(DEFAULT_PALETTES[key])
        self.update_preview()

    def choose_color(self, key):
        initial = self.current.get(key, "#ffffff")
        col = colorchooser.askcolor(initialcolor=initial, parent=self.top)
        if col and col[1]:
            self.current[key] = col[1]
            self.update_preview()

    def update_preview(self):
        accent = self.current.get("accent", "#007ACC")
        bg = self.current.get("bg", "#FFFFFF")
        fg = self.current.get("fg", "#000000")
        try:
            self.preview.configure(background=bg, foreground=fg, text=f"Accent: {accent}\nBG: {bg}\nFG: {fg}")
        except Exception:
            pass
        try:
            apply_theme_to_app(self.parent, self.current)
        except Exception:
            pass

    def on_save(self):
        save_theme(self.current)
        self.result = self.current
        try:
            apply_theme_to_app(self.parent, self.current)
        except Exception:
            pass
        self.top.destroy()

    def on_cancel(self):
        try:
            apply_theme_to_app(self.parent, self.initial)
        except Exception:
            pass
        self.top.destroy()

    def on_reset(self):
        self.current = dict(DEFAULT_PALETTES["Light"])
        self.update_preview()

# -------------------- App assets (logo & icon) -------------------------
def ensure_app_assets():
    """
    Generate a simple VD logo (logo.png) and vd.ico if Pillow is available.
    Returns (logo_path, ico_path).
    """
    base_dir = BASE_DIR
    logo_path = os.path.join(base_dir, "logo.png")
    ico_path = os.path.join(base_dir, "vd.ico")

    if not HAS_PIL:
        write_log("Pillow not available: skipping logo/icon generation.")
        return logo_path, ico_path

    try:
        W, H = 360, 96
        bg = (30, 144, 255, 255)
        fg = (255, 255, 255, 255)
        img = Image.new("RGBA", (W, H), (0,0,0,0))
        draw = ImageDraw.Draw(img)
        draw.rounded_rectangle((0,0,W,H), radius=18, fill=bg)
        try:
            f = ImageFont.truetype("arial.ttf", 56)
        except Exception:
            f = ImageFont.load_default()
        tw, th = draw.textsize("VD", font=f)
        draw.text(((W-tw)/2, (H-th)/2), "VD", font=f, fill=fg)
        img.save(logo_path, format="PNG")
        write_log(f"Generated logo at {logo_path}")
    except Exception as e:
        write_log(f"Logo generation failed: {e}")

    try:
        sizes = [(256,256),(128,128),(64,64),(48,48),(32,32)]
        images = []
        for (w,h) in sizes:
            img2 = Image.new("RGBA", (w,h), (0,0,0,0))
            draw2 = ImageDraw.Draw(img2)
            draw2.ellipse((0,0,w,h), fill=bg)
            try:
                f2 = ImageFont.truetype("arial.ttf", int(w*0.45))
            except Exception:
                f2 = ImageFont.load_default()
            txt = "VD"
            tw, th = draw2.textsize(txt, font=f2)
            draw2.text(((w-tw)/2, (h-th)/2 - h*0.03), txt, font=f2, fill=fg)
            images.append(img2)
        images[0].save(ico_path, format="ICO", sizes=[(s[0],s[1]) for s in sizes])
        write_log(f"Generated icon at {ico_path}")
    except Exception as e:
        write_log(f"Icon generation failed: {e}")

    return logo_path, ico_path

# -------------------- Utility helpers ----------------------------------
def clamp_interval(val):
    """Return integer interval clamped to 1..10 for toolbar display; accepts string or numeric."""
    try:
        ival = int(float(val))
    except Exception:
        ival = 10
    if ival < 1:
        ival = 1
    if ival > 10:
        # for toolbar display we clamp to 10, but DB may contain larger interval (Settings allows larger).
        ival = 10
    return ival

# -------------------- Entrypoint ---------------------------------------
def main():
    ensure_db_and_migrate()
    app = DashboardApp()
    app.mainloop()

if __name__ == "__main__":
    main()
