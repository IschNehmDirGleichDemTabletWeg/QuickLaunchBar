# QuickLaunchBar

A lightweight Quick Launch toolbar for Windows 11, built with Python and tkinter.

Displays shortcuts from the Windows Quick Launch folder as a compact icon grid — always on top, auto-hides on focus loss, lives in the system tray.

![QuickLaunchBar Screenshot](screenshot.png)

## Features

- Reads all `.lnk`, `.rdp`, `.exe`, `.url` files from the Quick Launch folder
- Icons loaded via Windows Shell API (custom icons, shell namespaces, all supported)
- Configurable grid layout (columns × rows) stored in registry
- Auto-hides when focus is lost
- System tray icon — left click to show, right click for menu
- Draggable borderless window
- Appears on the monitor where your mouse cursor is
- Dark theme

## Requirements

- Windows 10/11
- Python 3.10+
- Dependencies:

```
pip install pillow pywin32 screeninfo
```

## Usage

```
python QuickLaunchBar.py
```

Or build a standalone EXE:

```
pip install pyinstaller
pyinstaller --onefile --windowed --icon=quicklaunch.ico --name QuickLaunchBar QuickLaunchBar.py
```

## Quick Launch folder

```
%APPDATA%\Microsoft\Internet Explorer\Quick Launch
```

## Settings

Right-click the tray icon → **Einstellungen** to configure columns and max rows.  
Settings are saved to:

```
HKEY_CURRENT_USER\Software\QuickLaunchBar
```

| Registry Value | Type  | Default | Description         |
|----------------|-------|---------|---------------------|
| Columns        | DWORD | 8       | Icons per row       |
| MaxRows        | DWORD | 0       | Max rows (0 = auto) |

## Auto-start with Windows

1. Press `Win+R` → `shell:startup`
2. Place a shortcut to `QuickLaunchBar.exe` there

## License

MIT
