"""
QuickLaunchBar.py
Zeigt alle Verknüpfungen aus dem Quick Launch Ordner als Icon-Leiste.
Benötigt: pip install pillow pywin32
"""

import os
import sys
import threading
import subprocess
import winreg
import tkinter as tk
import tkinter.messagebox
from PIL import Image, ImageTk
import win32com.client
import win32gui
import win32ui
import win32con
import win32api

VERSION = "1.0"

REG_KEY = r"Software\QuickLaunchBar"

def reg_get(name: str, default):
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_KEY)
        val, _ = winreg.QueryValueEx(key, name)
        winreg.CloseKey(key)
        return val
    except Exception:
        return default

def reg_set(name: str, value):
    try:
        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, REG_KEY)
        if isinstance(value, int):
            winreg.SetValueEx(key, name, 0, winreg.REG_DWORD, value)
        else:
            winreg.SetValueEx(key, name, 0, winreg.REG_SZ, str(value))
        winreg.CloseKey(key)
    except Exception:
        pass


QUICK_LAUNCH = os.path.expandvars(
    r"%APPDATA%\Microsoft\Internet Explorer\Quick Launch"
)

BTN_SIZE  = 36
ICON_SIZE = 32
PAD       = 6
BG        = "#1c1c1c"
BTN_NORM  = "#2d2d2d"
BTN_HOVER = "#464646"
BTN_PRESS = "#5f5f5f"
BORDER    = "#4b4b4b"


# ── Icon extraction ───────────────────────────────────────────────────────────

def get_icon(path: str, index: int = 0):
    try:
        large, small = win32gui.ExtractIconEx(path, index, 10)
        if not large:
            print(f"ExtractIconEx no icons: '{path}'")
            return None

        hicon = large[0]
        for h in large[1:] + small:
            try: win32gui.DestroyIcon(h)
            except: pass

        hdc_screen = win32gui.GetDC(0)
        hdc        = win32ui.CreateDCFromHandle(hdc_screen)
        hdc_mem    = hdc.CreateCompatibleDC()
        hbmp       = win32ui.CreateBitmap()
        hbmp.CreateCompatibleBitmap(hdc, ICON_SIZE, ICON_SIZE)
        hdc_mem.SelectObject(hbmp)
        hdc_mem.FillSolidRect((0, 0, ICON_SIZE, ICON_SIZE), 0x1c1c1c)
        win32gui.DrawIconEx(hdc_mem.GetSafeHdc(), 0, 0, hicon,
                            ICON_SIZE, ICON_SIZE, 0, None, win32con.DI_NORMAL)

        bmpinfo = hbmp.GetInfo()
        bmpdata = hbmp.GetBitmapBits(True)
        img = Image.frombuffer(
            "RGBA",
            (bmpinfo["bmWidth"], bmpinfo["bmHeight"]),
            bmpdata, "raw", "BGRA", 0, 1
        )

        win32gui.DestroyIcon(hicon)
        win32gui.ReleaseDC(0, hdc_screen)

        return img.resize((ICON_SIZE, ICON_SIZE), Image.LANCZOS)  # PIL Image
    except Exception as e:
        print(f"get_icon EXCEPTION '{path}': {e}")
        return None


def _shell_icon(path: str):
    try:
        import ctypes
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
            hicon = info.hIcon
            hdc_screen = win32gui.GetDC(0)
            hdc        = win32ui.CreateDCFromHandle(hdc_screen)
            hdc_mem    = hdc.CreateCompatibleDC()
            hbmp       = win32ui.CreateBitmap()
            hbmp.CreateCompatibleBitmap(hdc, ICON_SIZE, ICON_SIZE)
            hdc_mem.SelectObject(hbmp)
            hdc_mem.FillSolidRect((0, 0, ICON_SIZE, ICON_SIZE), 0x1c1c1c)
            win32gui.DrawIconEx(hdc_mem.GetSafeHdc(), 0, 0, hicon,
                                ICON_SIZE, ICON_SIZE, 0, None, win32con.DI_NORMAL)
            bmpinfo = hbmp.GetInfo()
            bmpdata = hbmp.GetBitmapBits(True)
            img = Image.frombuffer(
                "RGBA",
                (bmpinfo["bmWidth"], bmpinfo["bmHeight"]),
                bmpdata, "raw", "BGRA", 0, 1
            )
            win32gui.ReleaseDC(0, hdc_screen)
            win32gui.DestroyIcon(hicon)
            return img.resize((ICON_SIZE, ICON_SIZE), Image.LANCZOS)  # PIL Image
    except Exception as e:
        print(f"_shell_icon EXCEPTION '{path}': {e}")
    return None


def resolve_lnk(lnk_path: str):
    """Gibt (target_path, icon_path, icon_index) zurück."""
    try:
        shell    = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortcut(lnk_path)
        loc      = shortcut.IconLocation
        if "," in loc:
            parts      = loc.rsplit(",", 1)
            icon_path  = parts[0].strip()
            try:    icon_index = int(parts[1].strip())
            except: icon_index = 0
        else:
            icon_path  = loc.strip()
            icon_index = 0

        # TargetPath kann leer sein wenn Argumente enthalten sind
        target = shortcut.TargetPath.strip()
        if not target:
            # Aus FullName / Arguments extrahieren
            target = shortcut.FullName.strip()

        return target, icon_path, icon_index
    except Exception:
        return "", "", 0


def best_icon(lnk_path: str):
    """Lädt Icon direkt aus der .lnk wie Windows Explorer – kein Target-Parsing nötig."""
    return _shell_icon(lnk_path)


def _shell_icon(path: str):
    try:
        import ctypes
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
            hicon = info.hIcon
            hdc_screen = win32gui.GetDC(0)
            hdc        = win32ui.CreateDCFromHandle(hdc_screen)
            hdc_mem    = hdc.CreateCompatibleDC()
            hbmp       = win32ui.CreateBitmap()
            hbmp.CreateCompatibleBitmap(hdc, ICON_SIZE, ICON_SIZE)
            hdc_mem.SelectObject(hbmp)
            hdc_mem.FillSolidRect((0, 0, ICON_SIZE, ICON_SIZE), 0x1c1c1c)
            win32gui.DrawIconEx(hdc_mem.GetSafeHdc(), 0, 0, hicon,
                                ICON_SIZE, ICON_SIZE, 0, None, win32con.DI_NORMAL)
            bmpinfo = hbmp.GetInfo()
            bmpdata = hbmp.GetBitmapBits(True)
            img = Image.frombuffer(
                "RGBA",
                (bmpinfo["bmWidth"], bmpinfo["bmHeight"]),
                bmpdata, "raw", "BGRA", 0, 1
            )
            win32gui.ReleaseDC(0, hdc_screen)
            win32gui.DestroyIcon(hicon)
            return img.resize((ICON_SIZE, ICON_SIZE), Image.LANCZOS)
    except Exception:
        pass
    return None


# ── Icon Button (Frame+Label, zuverlässiges Image-Rendering) ──────────────────

class IconButton(tk.Frame):
    def __init__(self, parent, name, img, on_click, on_right_click, **kw):
        super().__init__(
            parent,
            width=BTN_SIZE, height=BTN_SIZE,
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
            lbl._img = img  # extra GC-Schutz direkt am Label
        else:
            lbl.configure(text=name[:4], fg="white", font=("Segoe UI", 7))
        lbl.place(relx=0.5, rely=0.5, anchor="center")

        for w in (self, lbl):
            w.bind("<Enter>",           self._hover_on)
            w.bind("<Leave>",           self._hover_off)
            w.bind("<ButtonPress-1>",   self._press)
            w.bind("<ButtonRelease-1>", self._release)
            w.bind("<Button-3>",        self._right)

        self._tip = None
        for w in (self, lbl):
            w.bind("<Enter>", self._tip_show, add="+")
            w.bind("<Leave>", self._tip_hide, add="+")

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

    def _tip_show(self, e):
        if self._tip:
            return
        x = self.winfo_rootx() + BTN_SIZE // 2
        y = self.winfo_rooty() - 30
        self._tip = tk.Toplevel(self)
        self._tip.wm_overrideredirect(True)
        self._tip.wm_geometry(f"+{x}+{y}")
        tk.Label(self._tip, text=self._name,
                 bg="#2d2d2d", fg="white",
                 font=("Segoe UI", 9), padx=6, pady=3,
                 relief="flat").pack()

    def _tip_hide(self, e):
        # Nur verstecken wenn Maus wirklich den Button verlassen hat
        x, y = self.winfo_rootx(), self.winfo_rooty()
        if not (x <= e.x_root <= x + BTN_SIZE and y <= e.y_root <= y + BTN_SIZE):
            if self._tip:
                self._tip.destroy()
                self._tip = None


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

        # Border frame
        frame = tk.Frame(self.root, bg=BORDER, padx=1, pady=1)
        frame.pack(fill=tk.BOTH, expand=True)

        inner = tk.Frame(frame, bg=BG, padx=PAD, pady=PAD)
        inner.pack(fill=tk.BOTH, expand=True)

        # Drag via inner frame
        inner.bind("<ButtonPress-1>", self._drag_start)
        inner.bind("<B1-Motion>",     self._drag_move)

        self._btn_frame = tk.Frame(inner, bg=BG)
        self._btn_frame.pack()
        self._cols     = reg_get("Columns", 8)
        self._max_rows = reg_get("MaxRows", 0)

        self._images  = {}  # key=filename, value=ImageTk – niemals löschen (GC-Schutz)
        self._load_shortcuts()
        self._position_window()
        self._setup_tray()

    def _position_window(self):
        self.root.update_idletasks()
        w = self.root.winfo_reqwidth()
        h = self.root.winfo_reqheight()

        # Monitor wo die Maus gerade ist
        mx = self.root.winfo_pointerx()
        my = self.root.winfo_pointery()

        # Arbeitsbereich des aktuellen Monitors via tkinter
        # Fallback: alle Monitore via screeninfo wenn vorhanden
        try:
            import screeninfo
            for m in screeninfo.get_monitors():
                if m.x <= mx < m.x + m.width and m.y <= my < m.y + m.height:
                    x = m.x + m.width  - w - 10
                    y = m.y + m.height - h - 58
                    self.root.geometry(f"+{x}+{y}")
                    return
        except ImportError:
            pass

        # Fallback: primärer Monitor
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        self.root.geometry(f"+{sw - w - 10}+{sh - h - 58}")

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
        # _images NICHT leeren – GC-Schutz!

        if not os.path.exists(QUICK_LAUNCH):
            tk.Label(self._btn_frame,
                     text="Quick Launch\nnicht gefunden",
                     bg=BG, fg="#aaaaaa").pack()
            return

        EXTENSIONS = (".lnk", ".rdp", ".exe", ".bat", ".cmd", ".url")
        files = sorted(
            f for f in os.listdir(QUICK_LAUNCH)
            if os.path.splitext(f)[1].lower() in EXTENSIONS
            and os.path.isfile(os.path.join(QUICK_LAUNCH, f))
            and f.lower() != "desktop.ini"
        )

        for col, filename in enumerate(files):
            row = col // self._cols
            # MaxRows begrenzen
            if self._max_rows > 0 and row >= self._max_rows:
                break
            lnk      = os.path.join(QUICK_LAUNCH, filename)
            name     = os.path.splitext(filename)[0]
            pil_img  = best_icon(lnk)
            img = ImageTk.PhotoImage(pil_img) if pil_img else None
            if img:
                self._images[filename] = img

            btn = IconButton(
                self._btn_frame, name, img,
                on_click       = lambda p=lnk: self._launch(p),
                on_right_click = lambda e, p=lnk, n=name: self._ctx(e, p, n),
            )
            btn.grid(row=row, column=col % self._cols, padx=2, pady=2)

    def _launch(self, lnk_path):
        try:
            os.startfile(lnk_path)
            self.root.withdraw()
        except Exception as ex:
            tkinter.messagebox.showerror("Fehler", str(ex))

    def _ctx(self, event, lnk_path, name):
        menu = tk.Menu(self.root, tearoff=0,
                       bg="#2d2d2d", fg="white",
                       activebackground="#464646",
                       activeforeground="white")
        menu.add_command(label="Öffnen",
                         command=lambda: self._launch(lnk_path))
        menu.add_command(label="Im Explorer zeigen",
                         command=lambda: subprocess.Popen(
                             f'explorer /select,"{lnk_path}"'))
        menu.add_separator()
        menu.add_command(label="Entfernen",
                         command=lambda: self._remove(lnk_path, name))
        menu.tk_popup(event.x_root, event.y_root)

    def _remove(self, lnk_path, name):
        if tkinter.messagebox.askyesno(
                "Quick Launch", f"'{name}' löschen?"):
            try:
                os.remove(lnk_path)
                self._load_shortcuts()
            except Exception as ex:
                tkinter.messagebox.showerror("Fehler", str(ex))

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

        hwnd  = win32gui.CreateWindow(cls, "Tray", 0, 0, 0, 0, 0,
                                      0, 0, wc.hInstance, None)
        exe = sys.executable if getattr(sys, "frozen", False) else __file__
        try:
            hicon = win32gui.LoadImage(
                0,
                exe,
                win32con.IMAGE_ICON,
                16, 16,
                win32con.LR_LOADFROMFILE
            )
        except Exception as ex:
            import ctypes
            # Fallback: ExtractIconEx direkt
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
        menu.add_command(label="Neu laden",  command=self._load_shortcuts)
        menu.add_command(label="Einstellungen...", command=self._show_settings)
        menu.add_separator()
        menu.add_command(label="Beenden",    command=self.root.quit)
        menu.tk_popup(self.root.winfo_pointerx(),
                      self.root.winfo_pointery())

    def _show_settings(self):
        win = tk.Toplevel(self.root)
        win.title("Einstellungen")
        win.configure(bg="#2d2d2d")
        win.resizable(False, False)
        win.attributes("-topmost", True)

        # Center on screen
        win.update_idletasks()
        sw = win.winfo_screenwidth()
        sh = win.winfo_screenheight()
        w, h = 260, 140
        win.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

        pad = dict(padx=12, pady=6)

        tk.Label(win, text="Spalten:", bg="#2d2d2d", fg="white",
                 font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w", **pad)
        var_cols = tk.IntVar(value=reg_get("Columns", 8))
        tk.Spinbox(win, from_=1, to=20, textvariable=var_cols, width=6,
                   bg="#3d3d3d", fg="white", buttonbackground="#3d3d3d",
                   insertbackground="white").grid(row=0, column=1, sticky="w", **pad)

        tk.Label(win, text="Zeilen (max):", bg="#2d2d2d", fg="white",
                 font=("Segoe UI", 9)).grid(row=1, column=0, sticky="w", **pad)
        var_rows = tk.IntVar(value=reg_get("MaxRows", 0))
        tk.Spinbox(win, from_=0, to=20, textvariable=var_rows, width=6,
                   bg="#3d3d3d", fg="white", buttonbackground="#3d3d3d",
                   insertbackground="white").grid(row=1, column=1, sticky="w", **pad)
        tk.Label(win, text="(0 = automatisch)", bg="#2d2d2d", fg="#888888",
                 font=("Segoe UI", 8)).grid(row=1, column=2, sticky="w")

        def on_ok():
            reg_set("Columns", var_cols.get())
            reg_set("MaxRows", var_rows.get())
            self._cols = var_cols.get()
            self._max_rows = var_rows.get()
            self._load_shortcuts()
            win.destroy()

        tk.Button(win, text="OK", command=on_ok, width=8,
                  bg="#464646", fg="white", relief="flat",
                  activebackground="#5a5a5a",
                  cursor="hand2").grid(row=2, column=0, columnspan=3, pady=12)

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = QuickLaunchBar()
    app.run()
