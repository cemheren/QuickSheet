# QuickSheet

A lightweight interactive spreadsheet application for the terminal, built with C# and raw `System.Console` — zero dependencies.

![.NET 9](https://img.shields.io/badge/.NET-9.0-purple)
![License](https://img.shields.io/badge/license-MIT-blue)

## Features

- **Excel-like grid** that scales to your terminal size
- **Inline editing** — just start typing into any cell
- **Arrow key navigation** between cells
- **Dynamic column widths** that expand to fit content
- **Auto-sum** (Σ) per column and **auto-product** (Π) per row shown in the status bar
- **Clipboard** — cut, copy, paste cells
- **Row operations** — insert, delete, shift rows

## Getting Started

```bash
# Clone the repo
git clone https://github.com/cemheren/QuickSheet.git
cd QuickSheet

# Run (console mode)
dotnet run

# Load a CSV file
dotnet run -- data.csv

# Run as desktop background (Windows)
dotnet run -- --desktop
dotnet run -- --desktop data.csv
```

Requires [.NET 9 SDK](https://dotnet.microsoft.com/download/dotnet/9.0).

## Desktop Mode

### Windows

Run with `--desktop` to replace the Windows desktop with an interactive spreadsheet.
The form fills the working area (behind the taskbar) and resists Win+D minimization,
so it appears as your desktop when all other windows are minimized.

- **System tray icon** — Right-click for Save and Exit
- Click directly on the grid to focus it, then use all standard keyboard shortcuts
- Press **Win+D** — all apps minimize but QuickSheet stays visible as your desktop
- The taskbar remains fully visible and functional

### Linux (Fedora / X11)

Run with `--desktop` to replace the X11 desktop with an interactive spreadsheet.
Uses `_NET_WM_WINDOW_TYPE_DESKTOP` to sit below all other windows — the same
mechanism used by GNOME's desktop icons.

```bash
# Requires an X11 session (select "GNOME on Xorg" at the login screen)
dotnet run -- --desktop
dotnet run -- --desktop data.csv
```

- Works on any X11-based desktop environment (GNOME on Xorg, KDE, XFCE, etc.)
- No external dependencies — uses system libraries (`libX11`, `libXft`) via P/Invoke
- Desktop files from `~/Desktop` are shown in the grid
- Right-click to open files/hyperlinks
- Clipboard integrates with X11 selections

> **Wayland note:** True desktop-layer windows require X11. On a Wayland session,
> the app will attempt to run via XWayland but may not behave as a true desktop layer.
> Select "GNOME on Xorg" at the login screen for full support.

> **Cross-platform note:** All platform-specific code is isolated behind the
> `IDesktopHost` interface in `Platform/`. Windows uses WinForms + Win32 P/Invoke,
> Linux uses raw X11 P/Invoke.

## Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| Arrow Keys | Navigate cells |
| Tab | Next cell |
| Enter | Move down one row |
| Backspace | Delete last character |
| Delete | Clear cell |
| Ctrl+C | Copy cell |
| Ctrl+X | Cut cell |
| Ctrl+V | Paste cell |
| Ctrl+D | Delete row |
| Ctrl+O | Insert row (shift down) |
| Ctrl+P | Remove row (shift up) |
| Ctrl+S | Save to CSV |
| Ctrl+H | Show help |
| Ctrl+Q | Quit |

## License

MIT
