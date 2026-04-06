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

## Desktop Background Mode

Run with `--desktop` to embed the spreadsheet behind your desktop icons as a live wallpaper.

- **Alt + \`** — Toggle between desktop (view-only) and focused (interactive) mode
- **System tray icon** — Right-click for Toggle Focus, Save, and Exit
- All standard keyboard shortcuts work in focused mode

The status bar shows `[DESKTOP]` or `[FOCUSED]` to indicate the current mode.

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
