"""
QuickLaunchBar.py  –  v1.1
Zeigt alle Verknüpfungen aus dem Quick Launch Ordner als Icon-Leiste.
Benötigt: pip install pillow pywin32

Änderungen v1.1:
  - Settings werden nicht mehr in der Registry gespeichert,
    sondern in einer settings.json neben der EXE (portabel, kein Installer nötig)
  - Migration: vorhandene Registry-Werte werden beim ersten Start automatisch übernommen
"""

import os
import sys
import json
import threading
import subprocess
import ctypes
import tkinter as tk
import tkinter.messagebox
import tkinter.ttk as ttk
from PIL import Image, ImageTk
import win32com.client
import win32gui
import win32ui
import win32con
import win32api

VERSION = "1.2"

# ── Single Instance Guard ─────────────────────────────────────────────────────
_MUTEX_NAME = "QuickLaunchBar_SingleInstance_Mutex"
_mutex = ctypes.windll.kernel32.CreateMutexW(None, False, _MUTEX_NAME)
if ctypes.windll.kernel32.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
    ctypes.windll.kernel32.CloseHandle(_mutex)
    sys.exit(0)


# ── Settings (JSON) ───────────────────────────────────────────────────────────

def _settings_path() -> str:
    """settings.json liegt immer neben der EXE / dem Skript."""
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "settings.json")

DEFAULTS = {
    "Columns":     8,
    "MaxRows":     0,
    "IconSize":    32,
    "IconSpacing": 2,
    "TaskbarPos":  "bottom-right",
    "OffsetX":     8,
    "OffsetY":     50,
    "Order":       [],   # manuelle Icon-Reihenfolge (Dateinamen)
}

_settings: dict = {}

def _load_settings():
    global _settings
    path = _settings_path()
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                _settings = json.load(f)
        except Exception:
            _settings = {}
    else:
        _settings = {}
        _migrate_from_registry()   # einmalige Migration
    # Fehlende Schlüssel mit Defaults auffüllen
    for k, v in DEFAULTS.items():
        _settings.setdefault(k, v)

def _save_settings():
    try:
        with open(_settings_path(), "w", encoding="utf-8") as f:
            json.dump(_settings, f, indent=2, ensure_ascii=False)
    except Exception as ex:
        print(f"settings save error: {ex}")

def cfg_get(name: str):
    return _settings.get(name, DEFAULTS.get(name))

def cfg_set(name: str, value):
    _settings[name] = value
    _save_settings()

def _migrate_from_registry():
    """Liest alte Registry-Werte und übernimmt sie in die neue JSON."""
    try:
        import winreg
        REG_KEY = r"Software\QuickLaunchBar"
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_KEY)
        mapping = {
            "Columns":     ("Columns",     int),
            "MaxRows":     ("MaxRows",     int),
            "IconSize":    ("IconSize",    int),
            "IconSpacing": ("IconSpacing", int),
            "TaskbarPos":  ("TaskbarPos",  str),
            "OffsetX":     ("OffsetX",     int),
            "OffsetY":     ("OffsetY",     int),
        }
        for reg_name, (json_name, typ) in mapping.items():
            try:
                val, _ = winreg.QueryValueEx(key, reg_name)
                _settings[json_name] = typ(val)
            except Exception:
                pass
        winreg.CloseKey(key)
        print("Registry-Werte erfolgreich nach settings.json migriert.")
    except Exception:
        pass  # Keine Registry-Werte vorhanden – kein Problem


# Settings beim Start laden
_load_settings()


# ── Konstanten ────────────────────────────────────────────────────────────────

QUICK_LAUNCH = os.path.expandvars(
    r"%APPDATA%\Microsoft\Internet Explorer\Quick Launch"
)

PAD      = 6
BG       = "#1c1c1c"
BTN_NORM = "#2d2d2d"
BTN_HOVER= "#464646"
BTN_PRESS= "#5f5f5f"
BORDER   = "#4b4b4b"


# ── Icon extraction ───────────────────────────────────────────────────────────

def best_icon(path: str, icon_size: int = 32):
    """Lädt Icon direkt aus der .lnk wie Windows Explorer."""
    try:
        from ctypes import wintypes
        shell32 = ctypes.windll.shell32

        class SHFILEINFO(ctypes.Structure):
            _fields_ = [
                ("hIcon",         wintypes.HANDLE),
                ("iIcon",         ctypes.c_int),
                ("dwAttributes",  wintypes.DWORD),
                ("szDisplayName", ctypes.c_wchar * 260),
                ("szTypeName",    ctypes.c_wchar * 80),
            ]

        SHGFI_ICON      = 0x000000100
        SHGFI_LARGEICON = 0x000000000

        info = SHFILEINFO()
        res  = shell32.SHGetFileInfoW(
            path, 0, ctypes.byref(info), ctypes.sizeof(info),
            SHGFI_ICON | SHGFI_LARGEICON
        )
        if res and info.hIcon:
            hicon      = info.hIcon
            hdc_screen = win32gui.GetDC(0)
            hdc        = win32ui.CreateDCFromHandle(hdc_screen)
            hdc_mem    = hdc.CreateCompatibleDC()
            hbmp       = win32ui.CreateBitmap()
            hbmp.CreateCompatibleBitmap(hdc, icon_size, icon_size)
            hdc_mem.SelectObject(hbmp)
            hdc_mem.FillSolidRect((0, 0, icon_size, icon_size), 0x1c1c1c)
            win32gui.DrawIconEx(hdc_mem.GetSafeHdc(), 0, 0, hicon,
                                icon_size, icon_size, 0, None, win32con.DI_NORMAL)
            bmpinfo = hbmp.GetInfo()
            bmpdata = hbmp.GetBitmapBits(True)
            img = Image.frombuffer(
                "RGBA",
                (bmpinfo["bmWidth"], bmpinfo["bmHeight"]),
                bmpdata, "raw", "BGRA", 0, 1
            )
            win32gui.ReleaseDC(0, hdc_screen)
            win32gui.DestroyIcon(hicon)
            return img.resize((icon_size, icon_size), Image.LANCZOS)
    except Exception as e:
        print(f"best_icon EXCEPTION '{path}': {e}")
    return None


# ── Icon Button ───────────────────────────────────────────────────────────────

class IconButton(tk.Frame):
    def __init__(self, parent, name, img, btn_size, on_click, on_right_click, **kw):
        super().__init__(
            parent,
            width=btn_size, height=btn_size,
            bg=BTN_NORM,
            highlightthickness=1,
            highlightbackground=BORDER,
            **kw
        )
        self.pack_propagate(False)
        self._img_ref   = img
        self._name      = name
        self._on_click  = on_click
        self._on_rclick = on_right_click

        lbl = tk.Label(self, bg=BTN_NORM, cursor="hand2")
        if img:
            lbl.configure(image=img)
            lbl._img = img
        else:
            lbl.configure(text=name[:4], fg="white", font=("Segoe UI", 7))
        lbl.place(relx=0.5, rely=0.5, anchor="center")

        for w in (self, lbl):
            w.bind("<Enter>",           self._hover_on)
            w.bind("<Leave>",           self._hover_off)
            w.bind("<ButtonPress-1>",   self._press)
            w.bind("<ButtonRelease-1>", self._release)
            w.bind("<Button-3>",        self._right)

        self._tip     = None
        self._tip_job = None
        for w in (self, lbl):
            w.bind("<Enter>", self._tip_schedule, add="+")
            w.bind("<Leave>", self._tip_hide,     add="+")

    def _hover_on(self, e):
        self.configure(bg=BTN_HOVER)
        for c in self.winfo_children(): c.configure(bg=BTN_HOVER)
    def _hover_off(self, e):
        self.configure(bg=BTN_NORM)
        for c in self.winfo_children(): c.configure(bg=BTN_NORM)
    def _press(self, e):
        self.configure(bg=BTN_PRESS)
        for c in self.winfo_children(): c.configure(bg=BTN_PRESS)
    def _release(self, e):
        self._hover_on(e)
        self._on_click()
    def _right(self, e):
        self._on_rclick(e)

    def _tip_schedule(self, e):
        self._tip_cancel()
        self._tip_destroy()
        self._tip_job = self.after(400, self._tip_show)

    def _tip_cancel(self):
        if self._tip_job:
            self.after_cancel(self._tip_job)
            self._tip_job = None

    def _tip_destroy(self):
        if self._tip:
            self._tip.destroy()
            self._tip = None

    def _tip_show(self):
        self._tip_job = None
        if self._tip:
            return
        x = self.winfo_rootx() + self.winfo_width() // 2
        y = self.winfo_rooty() - 28
        self._tip = tk.Toplevel(self)
        self._tip.wm_overrideredirect(True)
        self._tip.wm_attributes("-topmost", True)
        self._tip.wm_geometry(f"+{x}+{y}")
        tk.Label(self._tip, text=self._name,
                 bg="#2d2d2d", fg="white",
                 font=("Segoe UI", 9), padx=6, pady=3,
                 relief="flat").pack()

    def _tip_hide(self, e):
        self._tip_cancel()
        self._tip_destroy()


# ── Main window ───────────────────────────────────────────────────────────────

class QuickLaunchBar:

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Quick Launch")
        self.root.configure(bg=BG)
        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)
        self.root.resizable(False, False)

        self.root.bind("<FocusOut>", self._on_focus_out)
        self.root.bind("<Escape>",   lambda e: self.root.withdraw())

        frame = tk.Frame(self.root, bg=BORDER, padx=1, pady=1)
        frame.pack(fill=tk.BOTH, expand=True)

        inner = tk.Frame(frame, bg=BG, padx=PAD, pady=PAD)
        inner.pack(fill=tk.BOTH, expand=True)

        inner.bind("<ButtonPress-1>", self._drag_start)
        inner.bind("<B1-Motion>",     self._drag_move)

        self._btn_frame = tk.Frame(inner, bg=BG)
        self._btn_frame.pack()

        self._reload_cfg()
        self._images = {}
        self._load_shortcuts()
        self._position_window()
        self.root.withdraw()   # beim Start versteckt – nur Tray sichtbar
        self._setup_tray()

    def _reload_cfg(self):
        """Liest alle Einstellungen aus dem Settings-Dict."""
        self._cols         = cfg_get("Columns")
        self._max_rows     = cfg_get("MaxRows")
        self._icon_size    = cfg_get("IconSize")
        self._icon_spacing = cfg_get("IconSpacing")
        self._tb_pos       = cfg_get("TaskbarPos")
        self._offset_x     = cfg_get("OffsetX")
        self._offset_y     = cfg_get("OffsetY")

    # ── Position & Drag ───────────────────────────────────────────────────────

    def _position_window(self):
        self.root.update_idletasks()
        w  = self.root.winfo_reqwidth()
        h  = self.root.winfo_reqheight()
        mx = self.root.winfo_pointerx()
        my = self.root.winfo_pointery()

        try:
            import screeninfo
            monitors = screeninfo.get_monitors()
            mon = next((m for m in monitors
                        if m.x <= mx < m.x + m.width
                        and m.y <= my < m.y + m.height), monitors[0])
            mx0, my0, mw, mh = mon.x, mon.y, mon.width, mon.height
        except Exception:
            mx0, my0 = 0, 0
            mw = self.root.winfo_screenwidth()
            mh = self.root.winfo_screenheight()

        ox, oy, pos = self._offset_x, self._offset_y, self._tb_pos
        if   pos == "top-left":     x, y = mx0 + ox,          my0 + oy
        elif pos == "top-right":    x, y = mx0 + mw - w - ox, my0 + oy
        elif pos == "bottom-left":  x, y = mx0 + ox,          my0 + mh - h - oy
        else:                       x, y = mx0 + mw - w - ox, my0 + mh - h - oy

        self.root.geometry(f"+{x}+{y}")

    def _drag_start(self, e):
        self._dx = e.x_root - self.root.winfo_x()
        self._dy = e.y_root - self.root.winfo_y()

    def _drag_move(self, e):
        self.root.geometry(f"+{e.x_root - self._dx}+{e.y_root - self._dy}")

    def _on_focus_out(self, e):
        self.root.after(150, self._check_focus)

    def _check_focus(self):
        try:
            if self.root.focus_get() is None:
                self.root.withdraw()
        except Exception:
            self.root.withdraw()

    # ── Load shortcuts ────────────────────────────────────────────────────────

    def _load_shortcuts(self):
        for w in self._btn_frame.winfo_children():
            w.destroy()

        if not os.path.exists(QUICK_LAUNCH):
            tk.Label(self._btn_frame,
                     text="Quick Launch\nfolder not found",
                     bg=BG, fg="#aaaaaa").pack()
            return

        EXTENSIONS = (".lnk", ".rdp", ".exe", ".bat", ".cmd", ".url")
        all_files = sorted(
            f for f in os.listdir(QUICK_LAUNCH)
            if os.path.splitext(f)[1].lower() in EXTENSIONS
            and os.path.isfile(os.path.join(QUICK_LAUNCH, f))
            and f.lower() != "desktop.ini"
        )

        # Gespeicherte Reihenfolge anwenden
        order = cfg_get("Order")
        ordered = [f for f in order if f in all_files]
        rest    = [f for f in all_files if f not in ordered]
        files   = ordered + rest

        icon_size = self._icon_size
        btn_size  = icon_size + 8
        sp        = self._icon_spacing

        for idx, filename in enumerate(files):
            row = idx // self._cols
            if self._max_rows > 0 and row >= self._max_rows:
                break
            lnk  = os.path.join(QUICK_LAUNCH, filename)
            name = os.path.splitext(filename)[0]

            pil_img = best_icon(lnk, icon_size)
            img = ImageTk.PhotoImage(pil_img) if pil_img else None
            if img:
                self._images[filename] = img

            btn = IconButton(
                self._btn_frame, name, img, btn_size,
                on_click       = lambda p=lnk: self._launch(p),
                on_right_click = lambda e, p=lnk, n=name: self._ctx(e, p, n),
            )
            btn.grid(row=row, column=idx % self._cols, padx=sp, pady=sp)

    def _launch(self, lnk_path):
        try:
            os.startfile(lnk_path)
            self.root.withdraw()
        except Exception as ex:
            tkinter.messagebox.showerror("Error", str(ex))

    def _ctx(self, event, lnk_path, name):
        menu = tk.Menu(self.root, tearoff=0,
                       bg="#2d2d2d", fg="white",
                       activebackground="#464646",
                       activeforeground="white")
        menu.add_command(label="Open",
                         command=lambda: self._launch(lnk_path))
        menu.add_command(label="Show in Explorer",
                         command=lambda: subprocess.Popen(
                             f'explorer /select,"{lnk_path}"'))
        menu.add_separator()
        menu.add_command(label="Remove",
                         command=lambda: self._remove(lnk_path, name))
        menu.tk_popup(event.x_root, event.y_root)

    def _remove(self, lnk_path, name):
        if tkinter.messagebox.askyesno(
                "Quick Launch Bar", f"Delete '{name}' from Quick Launch?"):
            try:
                os.remove(lnk_path)
                self._load_shortcuts()
            except Exception as ex:
                tkinter.messagebox.showerror("Error", str(ex))

    # ── Tray ──────────────────────────────────────────────────────────────────

    def _setup_tray(self):
        t = threading.Thread(target=self._tray_thread, daemon=True)
        t.start()

    def _tray_thread(self):
        WM_TRAY = win32con.WM_USER + 20
        TRAY_ID = 1

        def wnd_proc(hwnd, msg, wparam, lparam):
            if msg == WM_TRAY:
                if lparam in (win32con.WM_LBUTTONUP,
                              win32con.WM_LBUTTONDBLCLK):
                    self.root.after(0, self._show)
                elif lparam == win32con.WM_RBUTTONUP:
                    self.root.after(0, self._tray_menu)
            elif msg == win32con.WM_DESTROY:
                win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE,
                    (hwnd, TRAY_ID, win32gui.NIF_ICON, 0, 0, ""))
                win32gui.PostQuitMessage(0)
            return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)

        cls = "QLTray"
        wc  = win32gui.WNDCLASS()
        wc.lpfnWndProc   = wnd_proc
        wc.hInstance     = win32api.GetModuleHandle(None)
        wc.lpszClassName = cls
        win32gui.RegisterClass(wc)

        hwnd = win32gui.CreateWindow(cls, "Tray", 0, 0, 0, 0, 0,
                                     0, 0, wc.hInstance, None)
        self._tray_hwnd = hwnd
        self._tray_id   = TRAY_ID

        exe      = sys.executable if getattr(sys, "frozen", False) else __file__
        base_dir = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(exe)))
        ico_path = os.path.join(base_dir, "quicklaunch.ico")

        hicon = None
        if os.path.exists(ico_path):
            try:
                hicon = win32gui.LoadImage(
                    0, ico_path, win32con.IMAGE_ICON, 16, 16,
                    win32con.LR_LOADFROMFILE)
            except Exception:
                hicon = None

        if not hicon:
            try:
                hicon = win32gui.LoadImage(
                    win32api.GetModuleHandle(None), 1,
                    win32con.IMAGE_ICON, 16, 16, win32con.LR_DEFAULTCOLOR)
            except Exception:
                hicon = None

        if not hicon:
            try:
                large, small = win32gui.ExtractIconEx(exe, 0)
                hicon = small[0] if small else (large[0] if large else None)
                for h in (large[1:] if large else []) + (small[1:] if small else []):
                    try: win32gui.DestroyIcon(h)
                    except: pass
            except Exception:
                hicon = None

        if not hicon:
            hicon = win32gui.LoadIcon(0, win32con.IDI_APPLICATION)

        win32gui.Shell_NotifyIcon(win32gui.NIM_ADD, (
            hwnd, TRAY_ID,
            win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP,
            WM_TRAY, hicon, f"Quick Launch Bar v{VERSION}"))

        win32gui.PumpMessages()

    def _show(self):
        self._position_window()
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()

    def _tray_menu(self):
        menu = tk.Menu(self.root, tearoff=0,
                       bg="#2d2d2d", fg="white",
                       activebackground="#464646",
                       activeforeground="white")
        menu.add_command(label=f"Quick Launch Bar v{VERSION}", state="disabled")
        menu.add_separator()
        menu.add_command(label="Reload",      command=self._load_shortcuts)
        menu.add_command(label="Settings...", command=self._show_settings)
        menu.add_separator()
        menu.add_command(label="Exit",        command=self._quit)
        menu.tk_popup(self.root.winfo_pointerx(),
                      self.root.winfo_pointery())

    # ── Settings window ───────────────────────────────────────────────────────

    def _show_settings(self):
        win = tk.Toplevel(self.root)
        win.title("Settings")
        win.configure(bg="#2d2d2d")
        win.resizable(False, False)
        win.attributes("-topmost", True)

        sw = win.winfo_screenwidth()
        sh = win.winfo_screenheight()

        def _center():
            win.update_idletasks()
            w = win.winfo_reqwidth()
            h = win.winfo_reqheight()
            win.geometry(f"+{(sw-w)//2}+{(sh-h)//2}")
        win.after(10, _center)

        pad = dict(padx=12, pady=2)

        def spinbox(parent, var, from_, to):
            return tk.Spinbox(parent, from_=from_, to=to, textvariable=var, width=6,
                              bg="#3d3d3d", fg="white", buttonbackground="#3d3d3d",
                              insertbackground="white")

        tk.Label(win, text="Columns:", bg="#2d2d2d", fg="white",
                 font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w", **pad)
        var_cols = tk.IntVar(value=cfg_get("Columns"))
        spinbox(win, var_cols, 1, 20).grid(row=0, column=1, sticky="w", **pad)

        tk.Label(win, text="Max Rows:", bg="#2d2d2d", fg="white",
                 font=("Segoe UI", 9)).grid(row=1, column=0, sticky="w", **pad)
        var_rows = tk.IntVar(value=cfg_get("MaxRows"))
        spinbox(win, var_rows, 0, 20).grid(row=1, column=1, sticky="w", **pad)
        tk.Label(win, text="(0 = auto)", bg="#2d2d2d", fg="#888888",
                 font=("Segoe UI", 8)).grid(row=1, column=2, sticky="w")

        tk.Label(win, text="Icon Size:", bg="#2d2d2d", fg="white",
                 font=("Segoe UI", 9)).grid(row=2, column=0, sticky="w", **pad)
        var_icon = tk.StringVar(value=str(cfg_get("IconSize")))
        ttk.Combobox(win, textvariable=var_icon,
                     values=["8", "12", "16", "20", "24", "28", "32", "48"],
                     width=6, state="readonly").grid(row=2, column=1, sticky="w", **pad)

        tk.Label(win, text="Icon Spacing (px):", bg="#2d2d2d", fg="white",
                 font=("Segoe UI", 9)).grid(row=3, column=0, sticky="w", **pad)
        var_spacing = tk.IntVar(value=cfg_get("IconSpacing"))
        spinbox(win, var_spacing, 0, 20).grid(row=3, column=1, sticky="w", **pad)

        tk.Label(win, text="Taskbar Position:", bg="#2d2d2d", fg="white",
                 font=("Segoe UI", 9)).grid(row=4, column=0, sticky="w", **pad)
        tb_options = {
            "Top Left":     "top-left",
            "Top Right":    "top-right",
            "Bottom Left":  "bottom-left",
            "Bottom Right": "bottom-right",
        }
        pos_display   = {v: k for k, v in tb_options.items()}
        var_pos_label = tk.StringVar(
            value=pos_display.get(cfg_get("TaskbarPos"), "Bottom Right"))
        ttk.Combobox(win, textvariable=var_pos_label,
                     values=list(tb_options.keys()),
                     width=12, state="readonly").grid(row=4, column=1, columnspan=2,
                                                      sticky="w", **pad)

        tk.Label(win, text="Offset X (px):", bg="#2d2d2d", fg="white",
                 font=("Segoe UI", 9)).grid(row=5, column=0, sticky="w", **pad)
        var_ox = tk.IntVar(value=cfg_get("OffsetX"))
        spinbox(win, var_ox, 0, 500).grid(row=5, column=1, sticky="w", **pad)

        tk.Label(win, text="Offset Y (px):", bg="#2d2d2d", fg="white",
                 font=("Segoe UI", 9)).grid(row=6, column=0, sticky="w", **pad)
        var_oy = tk.IntVar(value=cfg_get("OffsetY"))
        spinbox(win, var_oy, 0, 500).grid(row=6, column=1, sticky="w", **pad)

        tk.Label(win, text="⚠ If columns/rows are too small,\n   not all shortcuts may be shown.",
                 bg="#2d2d2d", fg="#aaaaaa",
                 font=("Segoe UI", 8), justify="left").grid(
                 row=7, column=0, columnspan=3, padx=12, pady=(4, 2), sticky="w")

        def on_ok():
            cfg_set("Columns",     var_cols.get())
            cfg_set("MaxRows",     var_rows.get())
            cfg_set("IconSize",    int(var_icon.get()))
            cfg_set("IconSpacing", var_spacing.get())
            cfg_set("TaskbarPos",  tb_options[var_pos_label.get()])
            cfg_set("OffsetX",     var_ox.get())
            cfg_set("OffsetY",     var_oy.get())
            self._reload_cfg()
            self._load_shortcuts()
            win.destroy()

        tk.Button(win, text="OK", command=on_ok, width=8,
                  bg="#464646", fg="white", relief="flat",
                  activebackground="#5a5a5a",
                  cursor="hand2").grid(row=8, column=0, columnspan=3, pady=8)

    def _quit(self):
        """Tray-Icon sauber entfernen bevor das Fenster geschlossen wird."""
        try:
            win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, (
                self._tray_hwnd, self._tray_id,
                win32gui.NIF_ICON, 0, 0, ""))
        except Exception:
            pass
        self.root.quit()

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = QuickLaunchBar()
    app.run()