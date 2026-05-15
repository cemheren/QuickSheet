# Keyboard Shortcuts

Complete reference for QuickSheet keyboard shortcuts. Works in both **console mode**
and **desktop mode** (wallpaper) unless noted otherwise.

## Navigation

| Key | Action |
|-----|--------|
| `↑` `↓` `←` `→` | Move one cell |
| `Tab` | Move to next column |
| `Enter` | Edit current cell (enter edit mode) |
| `Escape` | Cancel edit / close help overlay |
| `Ctrl+G` | Go to cell — type a reference like `A1`, `C10` |

## Editing

| Key | Action |
|-----|--------|
| `Enter` | Start editing the current cell |
| `Escape` | Cancel edit, discard changes |
| `Backspace` | Delete last character (in edit mode) / clear cell (in nav mode) |
| `Delete` | Clear cell content |
| Any character | Start typing into the current cell |

## Clipboard

| Key | Action |
|-----|--------|
| `Ctrl+C` | Copy cell |
| `Ctrl+X` | Cut cell |
| `Ctrl+V` | Paste cell |

## Row Operations

| Key | Action |
|-----|--------|
| `Ctrl+D` | Delete current row |
| `Ctrl+O` | Insert row above (shift down) |
| `Ctrl+P` | Remove row (shift up) |

## Search & Sort

| Key | Action |
|-----|--------|
| `Ctrl+F` | Find — substring search across all cells |
| `Ctrl+B` | Sort by current column (toggles ascending ↔ descending) |

## Undo / Redo

| Key | Action |
|-----|--------|
| `Ctrl+Z` | Undo last action |
| `Ctrl+Y` | Redo |

## File & System

| Key | Action |
|-----|--------|
| `Ctrl+S` | Save to CSV |
| `Ctrl+T` | Cycle theme (Dark → Light → Solarized → Nord → …) |
| `Ctrl+H` | Show / hide help overlay |
| `Ctrl+Q` | Quit |

## Cell Prefixes

Type these at the start of a cell value to activate special behavior:

| Prefix | Behavior | Example |
|--------|----------|---------|
| `r: <cmd>` | **Run command** — press Enter on the cell to execute | `r: echo hello` |
| `i: <cmd>` | **Inline process** — live subprocess output streams into cells below | `i: top -b -n1` |
| `s: <values>` | **Sparkline** — renders a mini bar chart from comma-separated numbers | `s: 3,1,4,1,5,9` |
| `L: <path>` | **Load file** — imports contents of a text file | `L: data.csv` |
| `ext: <source>` | **Extension** — installs and activates an extension | `ext: github:cemheren/quicksheet-weather` |
| `http://` / `https://` | **Hyperlink** — auto-detected, opens in browser on Enter | `https://github.com` |
| `{A1::C10}` | **Cell range reference** — displays referenced range inline | `{A1::B5}` |

## Built-in Math

QuickSheet automatically computes:

- **Σ (Column Sum)** — shown in the status bar for numeric columns
- **Π (Row Product)** — shown in the status bar for numeric rows

## Desktop Mode Notes

Desktop mode (`--desktop`) supports the same shortcuts but input is handled via
platform-native key events (WinForms on Windows, X11 on Linux) rather than
`System.Console`. A few differences:

- **No terminal cursor** — navigation is visual only
- **Auto-saves every 5 seconds** to the loaded CSV file
- **`Ctrl+Q`** closes the desktop overlay and restores the normal wallpaper

---

*See also: [60-Second Tour](tour.md) · [Extensions](extensions.md) · [Recipes](recipes.md)*
