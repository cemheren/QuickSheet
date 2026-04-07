# QuickSheet

A lightweight interactive spreadsheet that lives on your desktop — built with C# and zero external dependencies. Use it as a terminal spreadsheet or replace your desktop wallpaper with a fully interactive grid.

![.NET 9](https://img.shields.io/badge/.NET-9.0-purple)
![License](https://img.shields.io/badge/license-MIT-blue)
![Windows](https://img.shields.io/badge/platform-Windows-0078D6?logo=windows)
![Linux](https://img.shields.io/badge/platform-Linux-FCC624?logo=linux&logoColor=black)

<!-- Add your own screenshots here:
![Console mode](docs/screenshots/console.png)
![Desktop mode on Windows](docs/screenshots/desktop-windows.png)
![Desktop mode on Linux](docs/screenshots/desktop-linux.png)
-->

## Features

### Spreadsheet

- **Excel-like grid** that scales to your terminal or screen size
- **Inline editing** — click a cell or just start typing
- **Arrow key navigation** between cells
- **Dynamic column widths** computed to fill the screen exactly
- **Auto-sum** (Σ) per column and **auto-product** (Π) per row in the status bar
- **CSV import/export** — load and save standard CSV files

### Selection & Clipboard

- **Multi-cell selection** — Shift+Arrow to extend, Ctrl+Click to toggle individual cells
- **Multi-line copy/paste** — copies selected cells as newline-separated text; paste splits lines across rows
- **System clipboard integration** — Ctrl+C/X/V use the OS clipboard (Windows & X11)
- **Row operations** — insert, delete, shift rows up/down

### Desktop Integration

- **Desktop file browser** — files from your `~/Desktop` folder appear in the grid
- **Double-click or Enter to open** files, folders, and hyperlinks
- **Clickable hyperlinks** — `http://` and `https://` URLs are highlighted and open in your browser
- **Runnable command cells** — prefix a cell with `r: ` to make it executable (e.g. `r: firefox`, `r: bash "echo hello"`)
- **Semi-transparent background** — see your wallpaper through the grid (85% opacity on Windows, ~80% on Linux)

### Reliability

- **Autosave every 5 seconds** — never lose work, even on a crash
- **Atomic file writes** — saves to a temp file first, then renames, to prevent corruption

## Getting Started

```bash
# Clone the repo
git clone https://github.com/cemheren/QuickSheet.git
cd QuickSheet

# Run in console mode
dotnet run

# Load a CSV file
dotnet run -- data.csv

# Run as a desktop background
dotnet run -- --desktop
dotnet run -- --desktop data.csv
```

Requires [.NET 9 SDK](https://dotnet.microsoft.com/download/dotnet/9.0).

## Desktop Mode

### Windows

Run with `--desktop` to replace the Windows desktop with an interactive spreadsheet.
The form fills the working area (behind the taskbar) and resists Win+D minimization,
so it appears as your desktop when all other windows are minimized.

<!-- ![Windows desktop mode](docs/screenshots/desktop-windows.png) -->

- **System tray icon** — right-click for Save and Exit
- **Win+D resistant** — all apps minimize but QuickSheet stays visible
- **Click to select** cells, then use all keyboard shortcuts
- **Semi-transparent** — your wallpaper shows through at 85% opacity
- **Desktop files** from `~/Desktop` are loaded into the grid automatically
- **Open anything** — double-click or press Enter on files, URLs, or `r:` commands
- The taskbar remains fully visible and functional

### Linux (X11)

Run with `--desktop` to place a spreadsheet window at the desktop layer using
`_NET_WM_WINDOW_TYPE_DESKTOP` — the same mechanism used by GNOME's desktop icons.
Uses raw X11 P/Invoke with zero NuGet dependencies.

<!-- ![Linux desktop mode](docs/screenshots/desktop-linux.png) -->

```bash
# Requires an X11 session (select "GNOME on Xorg" at the login screen)
dotnet run -- --desktop
dotnet run -- --desktop data.csv
```

- Works on any X11-based desktop environment (GNOME on Xorg, KDE, XFCE, etc.)
- **No external dependencies** — uses system libraries (`libX11.so.6`, `libXft.so.2`) via P/Invoke
- **ARGB transparency** — uses a 32-bit visual when available for true alpha blending
- **Desktop files** from `~/Desktop` are loaded into the grid (XDG fallback supported)
- **Right-click to open** files, hyperlinks, and runnable commands
- **X11 clipboard integration** — copy/paste works with other X11 applications
- **Terminal-aware commands** — `r:` shell commands open in your system terminal (ptyxis, kgx, gnome-terminal, etc.)

> **Wayland note:** True desktop-layer windows require X11. On a Wayland session,
> the app will attempt to run via XWayland but may not behave as a true desktop layer.
> Select "GNOME on Xorg" at the login screen for full support.

## Keyboard Shortcuts

### General

| Shortcut | Action |
|----------|--------|
| Arrow Keys | Navigate cells |
| Tab | Move to next cell |
| Enter | Move down / open selected item (desktop mode) |
| Backspace | Delete last character |
| Delete | Clear cell |
| Ctrl+S | Save to CSV |
| Ctrl+H | Show help |
| Ctrl+Q | Quit |

### Clipboard & Selection

| Shortcut | Action |
|----------|--------|
| Ctrl+C | Copy cell(s) to clipboard |
| Ctrl+X | Cut cell(s) to clipboard |
| Ctrl+V | Paste from clipboard (multi-line splits across rows) |
| Shift+Arrow | Extend selection |
| Ctrl+Click | Toggle individual cell selection |

### Row Operations

| Shortcut | Action |
|----------|--------|
| Ctrl+D | Delete current row |
| Ctrl+O | Insert row (shift down) |
| Ctrl+P | Remove row (shift up) |

### Desktop Mode

| Action | Trigger |
|--------|---------|
| Open file / URL / command | Double-click, Enter, or right-click (Linux) |
| Run a command cell | Enter on a cell prefixed with `r: ` |

## Architecture

All platform-specific code is isolated behind the `IDesktopHost` interface in `Platform/`.

```
Program.cs              → Entry point, CLI args, mode selection
SpreadsheetApp.cs       → Console-mode UI (raw System.Console)
GridManager.cs          → Pure data layer (no platform deps)
Platform/
  IDesktopHost.cs       → Cross-platform desktop interface
  Windows/
    WindowsDesktopHost.cs → WinForms + Win32 P/Invoke
    DesktopForm.cs        → Windows desktop window
  Linux/
    LinuxDesktopHost.cs   → X11 bootstrap
    DesktopWindow.cs      → X11 rendering, input, clipboard
    X11Methods.cs         → libX11/libXft P/Invoke bindings
```

## License

MIT
