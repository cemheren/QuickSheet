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

Run with `--desktop` to replace the Windows desktop with an interactive spreadsheet.
The form fills the working area (behind the taskbar) and resists Win+D minimization,
so it appears as your desktop when all other windows are minimized.

- **System tray icon** — Right-click for Save and Exit
- Click directly on the grid to focus it, then use all standard keyboard shortcuts
- Press **Win+D** — all apps minimize but QuickSheet stays visible as your desktop
- The taskbar remains fully visible and functional

> **Cross-platform note:** All Windows-specific code is isolated behind the
> `IDesktopHost` interface in `Platform/`. A Linux implementation can be added
> without touching the core `GridManager` or console UI.

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
