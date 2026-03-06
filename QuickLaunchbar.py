"""
QuickLaunchBar.py
Displays all shortcuts from the Quick Launch folder as an icon bar.
Requires: pip install pillow pywin32

Changes v1.1:  Settings moved from registry to portable settings.json
Changes v1.2:  Window starts hidden on startup; tray icon removed cleanly on exit
Changes v1.3:  Drag & drop icon reordering with ghost image and blue drop indicator
Changes v1.4:  Ctrl+Scroll to resize icons; extended icon size list
Changes v1.5:  Click app name in tray menu to open Quick Launch folder in Explorer
Changes v1.6:  Tray context menu closes correctly when clicking outside
Changes v1.7:  Background color configurable in Settings (Auto vs Manual with color picker)
Changes v1.8:  Icon background: Auto (brightness offset %) or Manual (color picker)
Changes v1.9:  Border color now visually distinct from button background
Changes v2.0:  Hover color: Auto (derived) or Manual (color picker)
"""

import os
import sys
import json
import threading
import subprocess
import ctypes
import ctypes.wintypes
import tkinter as tk
import tkinter.messagebox
import tkinter.ttk as ttk
from PIL import Image, ImageTk
import win32com.client
import win32gui
import win32ui
import win32con
import win32api

VERSION = "2.0"

# ── Single Instance Guard ─────────────────────────────────────────────────────
_MUTEX_NAME = "QuickLaunchBar_SingleInstance_Mutex"
_mutex = ctypes.windll.kernel32.CreateMutexW(None, False, _MUTEX_NAME)
if ctypes.windll.kernel32.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
    ctypes.windll.kernel32.CloseHandle(_mutex)
    sys.exit(0)


# ── Settings (JSON) ───────────────────────────────────────────────────────────

def _settings_path() -> str:
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
    "Order":       [],
    "BgColor":     "#000000",
    "BtnStyle":    "auto",
    "IconBgStyle":  "auto",
    "IconBgOffset": 15,
    "IconBgColor":  "#1c1c1c",
    "HoverStyle":   "auto",
    "HoverColor":   "#464646",
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
        _migrate_from_registry()
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
    except Exception:
        pass

_load_settings()


# ── Konstanten ────────────────────────────────────────────────────────────────

QUICK_LAUNCH = os.path.expandvars(
    r"%APPDATA%\Microsoft\Internet Explorer\Quick Launch"
)

PAD = 6

def _get_hover_color() -> str:
    """Berechnet die Hover-Farbe."""
    if _settings.get("HoverStyle", "auto") == "manual":
        return _settings.get("HoverColor", "#464646") or "#464646"
    # Auto: BTN_NORM nochmal eine Stufe heller/dunkler
    bg = (_settings.get("BgColor", "#000000") or "#000000").lstrip("#")
    r, g, b = int(bg[0:2],16), int(bg[2:4],16), int(bg[4:6],16)
    brightness = (r*299 + g*587 + b*114) // 1000
    step = 30
    if brightness < 128:
        r, g, b = min(255,r+step*2), min(255,g+step*2), min(255,b+step*2)
    else:
        r, g, b = max(0,r-step*2), max(0,g-step*2), max(0,b-step*2)
    return f"#{r:02x}{g:02x}{b:02x}"

def _get_icon_bg() -> str:
    """Berechnet die Icon-Hintergrundfarbe aus Settings."""
    style = _settings.get("IconBgStyle", "auto")
    if style == "manual":
        return _settings.get("IconBgColor", "#1c1c1c") or "#1c1c1c"
    bg = (_settings.get("BgColor", "#000000") or "#000000").lstrip("#")
    r, g, b = int(bg[0:2],16), int(bg[2:4],16), int(bg[4:6],16)
    offset = int(_settings.get("IconBgOffset", 15))
    brightness = (r*299 + g*587 + b*114) // 1000
    step = int(255 * offset / 100)
    if brightness < 128:
        r, g, b = min(255,r+step), min(255,g+step), min(255,b+step)
    else:
        r, g, b = max(0,r-step), max(0,g-step), max(0,b-step)
    return f"#{r:02x}{g:02x}{b:02x}"

def _get_bg() -> str:
    return _settings.get("BgColor", "#000000") or "#000000"

def _derive_colors(bg_hex: str):
    h = (bg_hex or "#000000").lstrip("#")
    r, g, b = int(h[0:2],16), int(h[2:4],16), int(h[4:6],16)
    def clamp(v): return max(0, min(255, v))
    def to_hex(r,g,b): return f"#{clamp(r):02x}{clamp(g):02x}{clamp(b):02x}"
    brightness = (r*299 + g*587 + b*114) // 1000
    step = 30 if brightness < 128 else -30
    return (
        to_hex(r+step,   g+step,   b+step),    # BTN_NORM
        to_hex(r+step*2, g+step*2, b+step*2),  # BTN_HOVER
        to_hex(r+step*3, g+step*3, b+step*3),  # BTN_PRESS
        to_hex(r+step*2, g+step*2, b+step*2),  # BORDER
    )

BG        = _get_bg()
BTN_NORM  = "#2d2d2d"
BTN_HOVER = "#464646"
BTN_PRESS = "#5f5f5f"
BORDER    = "#4b4b4b"
DRAG_IND  = "#3a7bd5"   # Blauer Drop-Indikator
DRAG_THRESHOLD = 5      # px bis Drag startet


# ── Icon extraction ───────────────────────────────────────────────────────────

def best_icon(path: str, icon_size: int = 32):
    try:
        shell32 = ctypes.windll.shell32

        class SHFILEINFO(ctypes.Structure):
            _fields_ = [
                ("hIcon",         ctypes.wintypes.HANDLE),
                ("iIcon",         ctypes.c_int),
                ("dwAttributes",  ctypes.wintypes.DWORD),
                ("szDisplayName", ctypes.c_wchar * 260),
                ("szTypeName",    ctypes.c_wchar * 80),
            ]

        info = SHFILEINFO()
        res  = shell32.SHGetFileInfoW(
            path, 0, ctypes.byref(info), ctypes.sizeof(info),
            0x000000100 | 0x000000000   # SHGFI_ICON | SHGFI_LARGEICON
        )
        if res and info.hIcon:
            hicon = info.hIcon

            class BITMAPINFOHEADER(ctypes.Structure):
                _fields_ = [
                    ("biSize",          ctypes.c_uint32),
                    ("biWidth",         ctypes.c_int32),
                    ("biHeight",        ctypes.c_int32),
                    ("biPlanes",        ctypes.c_uint16),
                    ("biBitCount",      ctypes.c_uint16),
                    ("biCompression",   ctypes.c_uint32),
                    ("biSizeImage",     ctypes.c_uint32),
                    ("biXPelsPerMeter", ctypes.c_int32),
                    ("biYPelsPerMeter", ctypes.c_int32),
                    ("biClrUsed",       ctypes.c_uint32),
                    ("biClrImportant",  ctypes.c_uint32),
                ]

            bih = BITMAPINFOHEADER()
            bih.biSize = ctypes.sizeof(BITMAPINFOHEADER)
            bih.biWidth = icon_size
            bih.biHeight = -icon_size
            bih.biPlanes = 1
            bih.biBitCount = 32
            bih.biCompression = 0

            hdc_screen = win32gui.GetDC(0)
            hdc        = win32ui.CreateDCFromHandle(hdc_screen)
            hdc_mem    = hdc.CreateCompatibleDC()

            pbits = ctypes.c_void_p()
            hbmp = ctypes.windll.gdi32.CreateDIBSection(
                hdc_mem.GetSafeHdc(), ctypes.byref(bih), 0,
                ctypes.byref(pbits), None, 0
            )
            old_bmp = ctypes.windll.gdi32.SelectObject(hdc_mem.GetSafeHdc(), hbmp)
            ctypes.windll.gdi32.PatBlt(hdc_mem.GetSafeHdc(), 0, 0, icon_size, icon_size, 0x00000042)
            win32gui.DrawIconEx(hdc_mem.GetSafeHdc(), 0, 0, hicon,
                                icon_size, icon_size, 0, None, win32con.DI_NORMAL)

            buf = (ctypes.c_uint8 * (icon_size * icon_size * 4))()
            ctypes.memmove(buf, pbits, icon_size * icon_size * 4)
            icon_img = Image.frombuffer("RGBA", (icon_size, icon_size),
                                        bytes(buf), "raw", "BGRA", 0, 1).copy()

            ctypes.windll.gdi32.SelectObject(hdc_mem.GetSafeHdc(), old_bmp)
            ctypes.windll.gdi32.DeleteObject(hbmp)
            win32gui.ReleaseDC(0, hdc_screen)
            win32gui.DestroyIcon(hicon)

            icon_bg = _get_icon_bg()
            _hex = icon_bg.lstrip("#")
            br, bg_, bb = int(_hex[0:2],16), int(_hex[2:4],16), int(_hex[4:6],16)
            bg_img = Image.new("RGBA", (icon_size, icon_size), (br, bg_, bb, 255))
            bg_img.paste(icon_img, mask=icon_img.split()[3])
            return bg_img.resize((icon_size, icon_size), Image.LANCZOS)
    except Exception as e:
        print(f"best_icon EXCEPTION '{path}': {e}")
    return None


# ── Icon Button ───────────────────────────────────────────────────────────────

class IconButton(tk.Frame):
    def __init__(self, parent, name, img, btn_size, on_click, on_right_click,
                 on_drag_start, on_drag_motion, on_drag_end, **kw):
        super().__init__(
            parent,
            width=btn_size, height=btn_size,
            bg=BTN_NORM,
            highlightthickness=1,
            highlightbackground=BORDER,
            **kw
        )
        self.pack_propagate(False)
        self._img_ref      = img
        self._name         = name
        self._on_click     = on_click
        self._on_rclick    = on_right_click
        self._on_drag_start  = on_drag_start
        self._on_drag_motion = on_drag_motion
        self._on_drag_end    = on_drag_end

        # Drag-State
        self._press_x   = 0
        self._press_y   = 0
        self._dragging  = False

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
            w.bind("<B1-Motion>",       self._motion)
            w.bind("<ButtonRelease-1>", self._release)
            w.bind("<Button-3>",        self._right)

        self._tip     = None
        self._tip_job = None
        for w in (self, lbl):
            w.bind("<Enter>", self._tip_schedule, add="+")
            w.bind("<Leave>", self._tip_hide,     add="+")

    # ── Hover ────────────────────────────────────────────────────────────────
    def _hover_on(self, e):
        if not self._dragging:
            self.configure(bg=BTN_HOVER)
            for c in self.winfo_children(): c.configure(bg=BTN_HOVER)
    def _hover_off(self, e):
        if not self._dragging:
            self.configure(bg=BTN_NORM)
            for c in self.winfo_children(): c.configure(bg=BTN_NORM)

    # ── Click / Drag ─────────────────────────────────────────────────────────
    def _press(self, e):
        self._press_x  = e.x_root
        self._press_y  = e.y_root
        self._dragging = False
        self.configure(bg=BTN_PRESS)
        for c in self.winfo_children(): c.configure(bg=BTN_PRESS)

    def _motion(self, e):
        if self._dragging:
            self._on_drag_motion(e.x_root, e.y_root)
            return
        dx = abs(e.x_root - self._press_x)
        dy = abs(e.y_root - self._press_y)
        if dx > DRAG_THRESHOLD or dy > DRAG_THRESHOLD:
            self._dragging = True
            self._tip_hide(e)          # Tooltip sofort weg
            self.configure(bg=BTN_NORM)
            for c in self.winfo_children(): c.configure(bg=BTN_NORM)
            self._on_drag_start(self, e.x_root, e.y_root)

    def _release(self, e):
        if self._dragging:
            self._dragging = False
            self._on_drag_end(e.x_root, e.y_root)
        else:
            self._hover_on(e)
            self._on_click()

    def _right(self, e):
        self._on_rclick(e)

    # ── Tooltip ──────────────────────────────────────────────────────────────
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
        if self._tip or self._dragging:
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

    def _tip_hide(self, e=None):
        self._tip_cancel()
        self._tip_destroy()

    def set_indicator(self, show: bool):
        """Blauen Rahmen als Drop-Indikator ein/ausschalten."""
        if show:
            self.configure(highlightthickness=2, highlightbackground=DRAG_IND)
        else:
            self.configure(highlightthickness=1, highlightbackground=BORDER)


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
        self.root.bind("<Control-MouseWheel>", self._on_ctrl_scroll)

        frame = tk.Frame(self.root, bg=BORDER, padx=1, pady=1)
        frame.pack(fill=tk.BOTH, expand=True)

        inner = tk.Frame(frame, bg=BG, padx=PAD, pady=PAD)
        inner.pack(fill=tk.BOTH, expand=True)

        inner.bind("<ButtonPress-1>",   self._drag_start)
        inner.bind("<B1-Motion>",       self._drag_move)

        self._btn_frame = tk.Frame(inner, bg=BG)
        self._btn_frame.pack()

        # Drag & Drop State
        self._ghost       = None   # Toplevel Geisterbild
        self._ghost_img   = None
        self._drag_btn    = None   # Button der gerade gezogen wird
        self._drag_idx    = None   # Index des gezogenen Icons
        self._drop_idx    = None   # aktueller Ziel-Index
        self._buttons     = []     # Liste aller IconButtons in Reihenfolge
        self._filenames   = []     # parallel zu _buttons: Dateinamen

        self._reload_cfg()
        self._apply_bg()
        self._images = {}
        self._load_shortcuts()
        self._position_window()
        self.root.withdraw()
        self._setup_tray()

    def _reload_cfg(self):
        global BG, BTN_NORM, BTN_HOVER, BTN_PRESS, BORDER
        BG = _get_bg()
        if cfg_get("BtnStyle") == "manual":
            BTN_NORM, BTN_HOVER, BTN_PRESS, BORDER = _derive_colors(BG)
        else:
            BTN_NORM, BTN_HOVER, BTN_PRESS, BORDER = "#2d2d2d", "#464646", "#5f5f5f", "#4b4b4b"
        BTN_HOVER = _get_hover_color()
        self._cols         = cfg_get("Columns")
        self._max_rows     = cfg_get("MaxRows")
        self._icon_size    = cfg_get("IconSize")
        self._icon_spacing = cfg_get("IconSpacing")
        self._tb_pos       = cfg_get("TaskbarPos")
        self._offset_x     = cfg_get("OffsetX")
        self._offset_y     = cfg_get("OffsetY")

    def _apply_bg(self):
        self.root.configure(bg=BG)
        for w in self.root.winfo_children():
            try: w.configure(bg=BG)
            except: pass
            for ww in w.winfo_children():
                try: ww.configure(bg=BG)
                except: pass
                for www in ww.winfo_children():
                    try: www.configure(bg=BG)
                    except: pass

    # ── Position & Window-Drag ────────────────────────────────────────────────

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

    def _on_ctrl_scroll(self, e):
        """Ctrl + Mausrad → Icon-Größe dynamisch ändern."""
        SIZES = [8, 12, 16, 20, 24, 28, 32, 36, 40, 44, 48]
        current = self._icon_size
        if current not in SIZES:
            # Nächstgelegenen Wert finden
            current = min(SIZES, key=lambda s: abs(s - current))
        idx = SIZES.index(current)
        if e.delta > 0:
            idx = min(idx + 1, len(SIZES) - 1)   # größer
        else:
            idx = max(idx - 1, 0)                  # kleiner
        new_size = SIZES[idx]
        if new_size != self._icon_size:
            cfg_set("IconSize", new_size)
            self._icon_size = new_size
            self._load_shortcuts()

    def _open_quick_launch(self, e=None):
        """Opens the Quick Launch folder in Windows Explorer."""
        subprocess.Popen(f'explorer "{QUICK_LAUNCH}"')

    def _on_focus_out(self, e):
        self.root.after(150, self._check_focus)

    def _check_focus(self):
        try:
            if self.root.focus_get() is None and self._ghost is None:
                self.root.withdraw()
        except Exception:
            self.root.withdraw()

    # ── Load shortcuts ────────────────────────────────────────────────────────

    def _load_shortcuts(self):
        for w in self._btn_frame.winfo_children():
            w.destroy()
        self._buttons   = []
        self._filenames = []

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

        order   = cfg_get("Order")
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
                on_click        = lambda p=lnk: self._launch(p),
                on_right_click  = lambda e, p=lnk, n=name: self._ctx(e, p, n),
                on_drag_start   = lambda b=None, x=0, y=0, i=idx: self._icon_drag_start(i, b, x, y),
                on_drag_motion  = self._icon_drag_motion,
                on_drag_end     = self._icon_drag_end,
            )
            btn.grid(row=row, column=idx % self._cols, padx=sp, pady=sp)
            self._buttons.append(btn)
            self._filenames.append(filename)

    # ── Icon Drag & Drop ──────────────────────────────────────────────────────

    def _icon_drag_start(self, idx, btn, x, y):
        self._drag_idx = idx
        self._drag_btn = btn
        self._drop_idx = idx

        # Geisterbild erstellen: Icon halbtransparent
        icon_size = self._icon_size
        btn_size  = icon_size + 8
        filename  = self._filenames[idx]

        pil_img = None
        if filename in self._images:
            # Aus dem ImageTk zurück zu PIL
            try:
                img_tk = self._images[filename]
                # Neue PIL Image mit Transparenz aus dem gecachten Icon
                lnk = os.path.join(QUICK_LAUNCH, filename)
                pil_img = best_icon(lnk, icon_size)
            except Exception:
                pass

        self._ghost = tk.Toplevel(self.root)
        self._ghost.wm_overrideredirect(True)
        self._ghost.wm_attributes("-topmost", True)
        self._ghost.wm_attributes("-alpha", 0.6)
        self._ghost.configure(bg=BTN_HOVER)

        ghost_frame = tk.Frame(
            self._ghost,
            width=btn_size, height=btn_size,
            bg=BTN_HOVER,
            highlightthickness=1,
            highlightbackground=DRAG_IND
        )
        ghost_frame.pack_propagate(False)
        ghost_frame.pack()

        if pil_img:
            self._ghost_img = ImageTk.PhotoImage(pil_img)
            lbl = tk.Label(ghost_frame, image=self._ghost_img, bg=BTN_HOVER)
            lbl._img = self._ghost_img
        else:
            lbl = tk.Label(ghost_frame,
                           text=self._filenames[idx][:4],
                           fg="white", bg=BTN_HOVER,
                           font=("Segoe UI", 7))
        lbl.place(relx=0.5, rely=0.5, anchor="center")

        offset = btn_size // 2
        self._ghost.wm_geometry(f"+{x - offset}+{y - offset}")

        # Drag-Button leicht ausblenden
        btn.configure(bg="#1c1c1c")
        for c in btn.winfo_children(): c.configure(bg="#1c1c1c")

    def _icon_drag_motion(self, x, y):
        if self._ghost is None:
            return

        # Geisterbild bewegen
        btn_size = self._icon_size + 8
        offset   = btn_size // 2
        self._ghost.wm_geometry(f"+{x - offset}+{y - offset}")

        # Alten Indikator löschen
        if self._drop_idx is not None and self._drop_idx < len(self._buttons):
            self._buttons[self._drop_idx].set_indicator(False)

        # Ziel-Index bestimmen: welcher Button liegt unter der Maus?
        new_drop = self._find_drop_index(x, y)
        self._drop_idx = new_drop

        # Neuen Indikator setzen
        if new_drop is not None and new_drop < len(self._buttons):
            self._buttons[new_drop].set_indicator(True)

    def _icon_drag_end(self, x, y):
        # Indikator entfernen
        if self._drop_idx is not None and self._drop_idx < len(self._buttons):
            self._buttons[self._drop_idx].set_indicator(False)

        # Geisterbild zerstören
        if self._ghost:
            self._ghost.destroy()
            self._ghost     = None
            self._ghost_img = None

        # Reihenfolge neu setzen wenn sinnvoll
        drag_idx = self._drag_idx
        drop_idx = self._drop_idx

        self._drag_btn  = None
        self._drag_idx  = None
        self._drop_idx  = None

        if drag_idx is not None and drop_idx is not None and drag_idx != drop_idx:
            # Filenames umsortieren
            files = list(self._filenames)
            item  = files.pop(drag_idx)
            files.insert(drop_idx, item)
            cfg_set("Order", files)
            self._load_shortcuts()
        else:
            # Drag abgebrochen – Button-Farbe zurücksetzen
            if drag_idx is not None and drag_idx < len(self._buttons):
                self._buttons[drag_idx].configure(bg=BTN_NORM)
                for c in self._buttons[drag_idx].winfo_children():
                    c.configure(bg=BTN_NORM)

    def _find_drop_index(self, x, y) -> int | None:
        """Gibt den Index des Buttons zurück der am nächsten zur Mausposition liegt."""
        best_idx  = None
        best_dist = float("inf")
        for i, btn in enumerate(self._buttons):
            if i == self._drag_idx:
                continue
            try:
                bx = btn.winfo_rootx() + btn.winfo_width()  // 2
                by = btn.winfo_rooty() + btn.winfo_height() // 2
                dist = (x - bx) ** 2 + (y - by) ** 2
                if dist < best_dist:
                    best_dist = dist
                    best_idx  = i
            except Exception:
                pass
        return best_idx

    # ── Launch & Context Menu ─────────────────────────────────────────────────

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
            WM_TRAY, hicon, f"Quick Launch Bar v{VERSION}\nClick to open Quick Launch folder"))

        win32gui.PumpMessages()

    def _show(self):
        self._position_window()
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()

    def _tray_menu(self):
        x = self.root.winfo_pointerx()
        y = self.root.winfo_pointery()

        menu = tk.Menu(self.root, tearoff=0,
                       bg="#2d2d2d", fg="white",
                       activebackground="#464646",
                       activeforeground="white")
        menu.add_command(label=f"Quick Launch Bar v{VERSION}  📂",
                         foreground="#aaaaaa",
                         command=self._open_quick_launch)
        menu.add_separator()
        menu.add_command(label="Reload",      command=self._load_shortcuts)
        menu.add_command(label="Settings...", command=self._show_settings)
        menu.add_separator()
        menu.add_command(label="Exit",        command=self._quit)

        # Hauptfenster zuerst verstecken, dann Hilfsfenster für Fokus
        self.root.withdraw()
        helper = tk.Toplevel(self.root)
        helper.overrideredirect(True)
        helper.geometry("1x1+0+0")
        helper.attributes("-alpha", 0)
        helper.focus_force()
        menu.bind("<Unmap>", lambda e: helper.destroy())
        menu.post(x, y)

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
                     values=["8", "12", "16", "20", "24", "28", "32", "36", "40", "44", "48"],
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

        rb_style = dict(bg="#2d2d2d", fg="white", selectcolor="#464646",
                        activebackground="#2d2d2d", activeforeground="white",
                        font=("Segoe UI", 9), cursor="hand2")

        # ── Icon Background Color ─────────────────────────────────────────
        tk.Label(win, text="Icon Background:", bg="#2d2d2d", fg="white",
                 font=("Segoe UI", 9)).grid(row=7, column=0, sticky="w", **pad)

        var_iconbg_style  = tk.StringVar(value=cfg_get("IconBgStyle")  or "auto")
        var_iconbg_offset = tk.IntVar(value=int(cfg_get("IconBgOffset") or 15))
        var_iconbg_color  = tk.StringVar(value=cfg_get("IconBgColor")  or "#1c1c1c")

        iconbg_row = tk.Frame(win, bg="#2d2d2d")
        iconbg_row.grid(row=7, column=1, columnspan=2, sticky="w", padx=12, pady=2)

        iconbg_preview = tk.Frame(iconbg_row, width=22, height=16,
                                  highlightthickness=1, highlightbackground="#666")

        def _update_iconbg_preview(*_):
            if var_iconbg_style.get() == "manual":
                iconbg_preview.configure(bg=var_iconbg_color.get())
            else:
                bg = (cfg_get("BgColor") or "#000000").lstrip("#")
                r2,g2,b2 = int(bg[0:2],16),int(bg[2:4],16),int(bg[4:6],16)
                brightness = (r2*299+g2*587+b2*114)//1000
                try: step = int(255 * var_iconbg_offset.get() / 100)
                except: step = 38
                if brightness < 128:
                    r2,g2,b2 = min(255,r2+step),min(255,g2+step),min(255,b2+step)
                else:
                    r2,g2,b2 = max(0,r2-step),max(0,g2-step),max(0,b2-step)
                iconbg_preview.configure(bg=f"#{r2:02x}{g2:02x}{b2:02x}")

        def pick_iconbg():
            from tkinter.colorchooser import askcolor
            result = askcolor(color=var_iconbg_color.get(), parent=win, title="Icon Background Color")
            if result[1]:
                var_iconbg_color.set(result[1])
                _update_iconbg_preview()

        offset_spin = tk.Spinbox(iconbg_row, from_=0, to=100,
                                 textvariable=var_iconbg_offset, width=4,
                                 bg="#3d3d3d", fg="white", buttonbackground="#3d3d3d",
                                 insertbackground="white",
                                 command=_update_iconbg_preview)
        offset_spin.bind("<KeyRelease>", _update_iconbg_preview)

        btn_iconbg_pick = tk.Button(iconbg_row, text="Pick...", command=pick_iconbg, width=6,
                                    bg="#464646", fg="white", relief="flat",
                                    activebackground="#5a5a5a", cursor="hand2")
        pct_label = tk.Label(iconbg_row, text="%", bg="#2d2d2d", fg="#aaaaaa",
                             font=("Segoe UI", 9))

        def on_iconbg_style(*_):
            is_manual = var_iconbg_style.get() == "manual"
            offset_spin.configure(state="disabled" if is_manual else "normal")
            pct_label.configure(fg="#555555" if is_manual else "#aaaaaa")
            btn_iconbg_pick.configure(state="normal" if is_manual else "disabled")
            _update_iconbg_preview()

        tk.Radiobutton(iconbg_row, text="Auto", variable=var_iconbg_style,
                       value="auto",   command=on_iconbg_style, **rb_style).pack(side="left")
        offset_spin.pack(side="left", padx=(6,0))
        pct_label.pack(side="left", padx=(2,8))
        tk.Radiobutton(iconbg_row, text="Manual", variable=var_iconbg_style,
                       value="manual", command=on_iconbg_style, **rb_style).pack(side="left", padx=(0,4))
        iconbg_preview.pack(side="left", padx=(0,4))
        btn_iconbg_pick.pack(side="left")
        on_iconbg_style()

        # ── Window Background Color ───────────────────────────────────────
        tk.Label(win, text="Background Color:", bg="#2d2d2d", fg="white",
                 font=("Segoe UI", 9)).grid(row=8, column=0, sticky="w", **pad)

        var_btn_style = tk.StringVar(value=cfg_get("BtnStyle") or "auto")
        var_bg        = tk.StringVar(value=cfg_get("BgColor") or "#000000")

        bg_row = tk.Frame(win, bg="#2d2d2d")
        bg_row.grid(row=8, column=1, columnspan=2, sticky="w", padx=12, pady=2)

        bg_preview = tk.Frame(bg_row, width=22, height=16, bg=var_bg.get(),
                              highlightthickness=1, highlightbackground="#666")

        def pick_bg():
            from tkinter.colorchooser import askcolor
            result = askcolor(color=var_bg.get(), parent=win, title="Window Color")
            if result[1]:
                var_bg.set(result[1])
                bg_preview.configure(bg=result[1])
                _update_iconbg_preview()

        btn_bg_pick = tk.Button(bg_row, text="Pick...", command=pick_bg, width=6,
                                bg="#464646", fg="white", relief="flat",
                                activebackground="#5a5a5a", cursor="hand2")

        def on_btnstyle(*_):
            pass  # Auto/Manual für Button-Farben – future use

        bg_preview.pack(side="left", padx=(0,4))
        btn_bg_pick.pack(side="left")

        # ── Hover Color ───────────────────────────────────────────────────
        tk.Label(win, text="Hover Color:", bg="#2d2d2d", fg="white",
                 font=("Segoe UI", 9)).grid(row=9, column=0, sticky="w", **pad)

        var_hover_style = tk.StringVar(value=cfg_get("HoverStyle") or "auto")
        var_hover_color = tk.StringVar(value=cfg_get("HoverColor") or "#464646")

        hover_row = tk.Frame(win, bg="#2d2d2d")
        hover_row.grid(row=9, column=1, columnspan=2, sticky="w", padx=12, pady=2)

        hover_preview = tk.Frame(hover_row, width=22, height=16,
                                 highlightthickness=1, highlightbackground="#666")

        def _update_hover_preview(*_):
            if var_hover_style.get() == "manual":
                hover_preview.configure(bg=var_hover_color.get())
            else:
                bg = (cfg_get("BgColor") or "#000000").lstrip("#")
                r2,g2,b2 = int(bg[0:2],16),int(bg[2:4],16),int(bg[4:6],16)
                brightness = (r2*299+g2*587+b2*114)//1000
                step = 30
                if brightness < 128:
                    r2,g2,b2 = min(255,r2+step*2),min(255,g2+step*2),min(255,b2+step*2)
                else:
                    r2,g2,b2 = max(0,r2-step*2),max(0,g2-step*2),max(0,b2-step*2)
                hover_preview.configure(bg=f"#{r2:02x}{g2:02x}{b2:02x}")

        def pick_hover():
            from tkinter.colorchooser import askcolor
            result = askcolor(color=var_hover_color.get(), parent=win, title="Hover Color")
            if result[1]:
                var_hover_color.set(result[1])
                _update_hover_preview()

        btn_hover_pick = tk.Button(hover_row, text="Pick...", command=pick_hover, width=6,
                                   bg="#464646", fg="white", relief="flat",
                                   activebackground="#5a5a5a", cursor="hand2")

        rb_hover = dict(bg="#2d2d2d", fg="white", selectcolor="#464646",
                        activebackground="#2d2d2d", activeforeground="white",
                        font=("Segoe UI", 9), cursor="hand2")

        def on_hover_style(*_):
            is_manual = var_hover_style.get() == "manual"
            btn_hover_pick.configure(state="normal" if is_manual else "disabled")
            _update_hover_preview()

        tk.Radiobutton(hover_row, text="Auto", variable=var_hover_style,
                       value="auto",   command=on_hover_style, **rb_hover).pack(side="left")
        tk.Radiobutton(hover_row, text="Manual", variable=var_hover_style,
                       value="manual", command=on_hover_style, **rb_hover).pack(side="left", padx=(8,4))
        hover_preview.pack(side="left", padx=(0,4))
        btn_hover_pick.pack(side="left")
        on_hover_style()
        _update_hover_preview()

        tk.Label(win, text="⚠ If columns/rows are too small,\n   not all shortcuts may be shown.",
                 bg="#2d2d2d", fg="#aaaaaa",
                 font=("Segoe UI", 8), justify="left").grid(
                 row=10, column=0, columnspan=3, padx=12, pady=(4, 2), sticky="w")

        def on_ok():
            cfg_set("Columns",      var_cols.get())
            cfg_set("MaxRows",      var_rows.get())
            cfg_set("IconSize",     int(var_icon.get()))
            cfg_set("IconSpacing",  var_spacing.get())
            cfg_set("TaskbarPos",   tb_options[var_pos_label.get()])
            cfg_set("OffsetX",      var_ox.get())
            cfg_set("OffsetY",      var_oy.get())
            cfg_set("IconBgStyle",  var_iconbg_style.get())
            cfg_set("IconBgOffset", var_iconbg_offset.get())
            cfg_set("IconBgColor",  var_iconbg_color.get())
            cfg_set("HoverStyle",   var_hover_style.get())
            cfg_set("HoverColor",   var_hover_color.get())
            cfg_set("BgColor",      var_bg.get())
            self._reload_cfg()
            self._apply_bg()
            self._load_shortcuts()
            win.destroy()

        tk.Button(win, text="OK", command=on_ok, width=8,
                  bg="#464646", fg="white", relief="flat",
                  activebackground="#5a5a5a",
                  cursor="hand2").grid(row=11, column=0, columnspan=3, pady=8)

    # ── Quit ──────────────────────────────────────────────────────────────────

    def _quit(self):
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