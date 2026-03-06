# QuickLaunchBar

A lightweight Quick Launch toolbar for Windows 11, built with Python and tkinter.

Displays shortcuts from the Windows Quick Launch folder as a compact icon grid — always on top, auto-hides on focus loss, lives in the system tray.

## Features

- Reads all `.lnk`, `.rdp`, `.exe`, `.bat`, `.cmd`, `.url` files from the Quick Launch folder
- Icons loaded via Windows Shell API with true alpha transparency (DIB section, RGBA compositing)
- Configurable grid layout (columns × rows)
- Auto-hides when focus is lost
- System tray icon — left click to show, right click for menu
- Draggable borderless window
- Appears on the monitor where your mouse cursor is
- Drag & drop icon reordering — order saved automatically
- `Ctrl + Scroll` to resize icons on the fly
- Fully customizable colors (background, icon background, hover, borders)
- Portable — all settings saved to `settings.json` next to the EXE, no registry

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
pyinstaller --onefile --windowed --icon=quicklaunch.ico --add-data "quicklaunch.ico;." --name QuickLaunchBar QuickLaunchBar.py
```

## Quick Launch folder

```
%APPDATA%\Microsoft\Internet Explorer\Quick Launch
```

## Settings

Right-click the tray icon → **Settings...** to configure all options.  
Settings are saved to `settings.json` next to the script / EXE.

| Setting | Default | Description |
|---|---|---|
| Columns | 8 | Icons per row |
| Max Rows | 0 | Max rows (0 = auto) |
| Icon Size | 32 | Icon size in pixels (8–48) |
| Icon Spacing | 2 | Spacing between icons in pixels |
| Taskbar Position | bottom-right | Where the bar appears (4 corners) |
| Offset X / Y | 8 / 50 | Distance from screen edge in pixels |
| Icon Background | Auto 15% | Auto = derived from background color by brightness offset; Manual = custom color picker |
| Background Color | #000000 | Window background color |
| Hover Color | Auto | Auto = derived from background color; Manual = custom color picker |

## Auto-start with Windows

1. Press `Win+R` → `shell:startup`
2. Place a shortcut to `QuickLaunchBar.exe` (or `QuickLaunchBar.py`) there

## Changelog

| Version | Changes |
|---|---|
| v1.1 | Settings moved from registry to portable `settings.json` |
| v1.2 | Window starts hidden on startup; tray icon removed cleanly on exit |
| v1.3 | Drag & drop icon reordering with ghost image and blue drop indicator |
| v1.4 | `Ctrl+Scroll` to resize icons; extended icon size list |
| v1.5 | Click app name in tray menu to open Quick Launch folder in Explorer |
| v1.6 | Tray context menu closes correctly when clicking outside |
| v1.7 | Background color configurable in Settings |
| v1.8 | Icon background: Auto (brightness offset %) or Manual (color picker) |
| v1.9 | Border color now visually distinct from button background |
| v2.0 | Hover color: Auto (derived) or Manual (color picker) |

## License

MIT
# QuickLaunchBar

A lightweight Quick Launch toolbar for Windows 11, built with Python and tkinter.
<img width="27" height="30" alt="grafik" src="https://github.com/user-attachments/assets/1cdf1422-1cc9-4e31-9e82-7fbd99c3389e" />


Displays shortcuts from the Windows Quick Launch folder as a compact icon grid — always on top, auto-hides on focus loss, lives in the system tray.

<img width="707" height="283" alt="grafik" src="https://github.com/user-attachments/assets/a4c2fde3-ebc9-4b98-b314-c6bdb74170ec" />




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
