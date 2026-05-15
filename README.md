# QuickSheet

**Your desktop is a spreadsheet.**

QuickSheet replaces your wallpaper with a transparent, interactive grid. Pin notes, launch apps, paste links, track numbers — all without opening a window. The idea: keep something lightweight always present in the background, instead of a static wallpaper you never interact with.

![QuickSheet running as the desktop wallpaper — cells holding runnable commands behind every open window](desktop-wallpaper-commands.png)

![.NET 9](https://img.shields.io/badge/.NET-9.0-purple)
![License](https://img.shields.io/badge/license-MIT-blue)
![Windows](https://img.shields.io/badge/platform-Windows-0078D6?logo=windows)
![Linux](https://img.shields.io/badge/platform-Linux-FCC624?logo=linux&logoColor=black)

## Quick Start

```bash
git clone https://github.com/cemheren/QuickSheet.git
cd QuickSheet
dotnet run -c Release --project ExcelConsole.csproj -- --desktop
```

Requires [.NET 9 SDK](https://dotnet.microsoft.com/download/dotnet/9.0). Release mode is recommended — it feels noticeably snappier. Drop `--desktop` to launch in plain TUI mode inside any terminal.

Tip: add it to your startup applications so your notes, links, and launchers are there every time you log in.

Want to export your data? Render any CSV as a GitHub-flavored Markdown table:

```bash
dotnet run --project ExcelConsole.csproj -- mydata.csv --export-md mydata.md
```

New here? The 60-second tour is in [docs/tour.md](docs/tour.md). Or jump straight to [docs/recipes.md](docs/recipes.md) for ready-to-paste dashboard layouts.

## A note on the code

This is a side project, and a lot of it was written with AI assistance. The hard rule the project keeps is **zero NuGet dependencies** — all native interop (X11, WinForms, ConPTY) is hand-written P/Invoke. That keeps the supply-chain surface area minimal: clone, build, run.

## Desktop environment for developers 
### Type directly into cells 
For quick notes, todos. Everything autosaves every 5 seconds. You can point and click to any part of your desktop to take some quick notes, use it as a buffer etc without losing focus of the application you are running. 
  
### Launcher
<!-- ![Launcher example](docs/screenshots/use-case-launcher.png) -->
Prefix any cell with `r: ` to make it a runnable command — e.g. `r: code .` or `r: firefox`.
You can open or launch multiple repos with a single operation. Multi select cells, and hit enter to run. 

I've used it to start the repos I want to work on for the day, and launch copilot with some saved prompts like summarize emails. Not sure how others solve this problem, but this to me is simpler than running startup scripts. 

![QuickSheet cells with runnable commands and links on the desktop](desktop-wallpaper-commands.png)

### Hyper-Link Dashboard
<!-- ![Links example](docs/screenshots/use-case-links.png) -->
Paste URLs into cells. They're highlighted and open in your browser on Enter/double-click.
Similar to the launcher funcitonality you can open and run multiple by selecting multiple cells. I was going for a emacs buffer type of feel to save and run multiple cells. 

![Hyperlink dashboard — clickable URLs organized in a grid](hyperlink-dashboard.png)

### Lightweight Data Tracking
<!-- ![Data example](docs/screenshots/use-case-data.png) -->
Auto-sum (Σ) per column and auto-product (Π) per row in the status bar. Import/export CSV. I hate opening the calculator for simple operations. This helps with that. 

![Lightweight data tracking with auto-sum per column](data-tracking-autosum.png)

### Sparklines in a cell
Prefix a cell with `s: 1,2,3,4,5,6` to render the values as a unicode bar sparkline (`▁▂▃▄▅▆`). Handy for tracking a small series next to other notes — paste a row of numbers, get a tiny chart, no extra column.

You can also point at a range of cells: `s: A1::A10` pulls numeric values from the referenced grid range and renders them. Non-numeric cells in the range are skipped.

<!-- 
## Add your own sections here!
Some ideas:
- Daily standups / sprint tracking
- Monitoring dashboard with r: commands
- Project-specific bookmarks
-->

### Desktop files
Desktop files are added to cells (padded to right), which can be used in a multi-select way. Helpful for finding/launching multiple files. Not sure about usability of these yet, likely I'm going to tweak this. 

![Desktop files rendered as clickable cells for quick access](desktop-files-grid.png)

## Extensions — Make Your Desktop Do More

QuickSheet has a lightweight extension system that lets you bring new capabilities right into your grid. Extensions are standalone programs that communicate with QuickSheet over a simple JSON-lines protocol — install one in seconds and it just works.

> **Full list:** see [docs/extensions.md](docs/extensions.md) for the live directory of available extensions and the protocol spec.

### Install an extension in one cell

Type `ext: github:user/repo` into any cell and press Enter. QuickSheet clones the repo, reads its manifest, and starts the extension automatically. That's it — no package managers, no config files.

### Example: GitHub Copilot on your desktop

The [quicksheet-copilot-ext](https://github.com/cemheren/quicksheet-copilot-ext) extension brings AI directly into your spreadsheet grid. Ask questions, generate structured data, or summarize cell ranges — all without leaving your desktop.

```
ext: github:cemheren/quicksheet-copilot-ext
copilot: test, 1, 1
```

![Copilot extension running on the desktop](copilot-ext-example.png)

You can reference cell ranges with `{A1::C10}` syntax and Copilot receives the actual cell contents as context:

```
copilot: summarize this data {B1::E50}, 3, 1
copilot: generate random test data with name and age, 2, 10
```

### Example: Weather forecast widget

The [quicksheet-weather](https://github.com/cemheren/quicksheet-weather) extension turns a cell into a live 7-day weather forecast:

```
ext: github:cemheren/quicksheet-weather
wthr: Seattle, 2, 7
```

![Weather extension showing a 7-day forecast on the desktop](weather-ext-example.png)

### Example: TLS certificate checker

The [quicksheet-tls-ext](https://github.com/cemheren/quicksheet-tls-ext) extension turns a cell into a live TLS cert expiry/issuer readout — useful as an ambient SRE dashboard on the wallpaper:

```
ext: github:cemheren/quicksheet-tls-ext
tls: github.com, 1, 4
```

### Example: Crypto price quotes

The [quicksheet-price-ext](https://github.com/cemheren/quicksheet-price-ext) extension turns a cell into a live CoinGecko price quote with 24h change — a one-liner ambient portfolio dashboard:

```
ext: github:cemheren/quicksheet-price-ext
price: btc, 1, 2
```

### Example: Dictionary lookups

The [quicksheet-define-ext](https://github.com/cemheren/quicksheet-define-ext) extension puts an inline dictionary in any cell — handy for writers and people learning a language:

```
ext: github:cemheren/quicksheet-define-ext
def: laconic, 1, 4
```

### Example: Mortgage calculator

The [quicksheet-mortgage-ext](https://github.com/cemheren/quicksheet-mortgage-ext) extension turns a cell into a live amortization calculator — useful for side-by-side loan comparisons:

```
ext: github:cemheren/quicksheet-mortgage-ext
mort: 500000, 6.5, 30, 1, 4
```

### Example: Pomodoro timer

The [quicksheet-pomodoro](https://github.com/cemheren/quicksheet-pomodoro) extension adds a live countdown timer to your desktop — perfect for focus sessions:

```
ext: github:cemheren/quicksheet-pomodoro
pomo: 25, 3, 2
```

Shows a live-updating countdown with progress bar: `🍅 FOCUS  23:41  [████░░░░░░]`. Supports `pomo: break` (5 min) and `pomo: long` (15 min) for the full Pomodoro technique.

### Example: System monitor

The [quicksheet-sysmon](https://github.com/cemheren/quicksheet-sysmon) extension turns your desktop into a live system dashboard — CPU, RAM, disk usage with visual bars:

```
ext: github:cemheren/quicksheet-sysmon
sys: all
```

Shows color-coded metrics with progress bars: `🟢 CPU 12.3% [██░░░░░░░░░░░░░░░░░░]`. Supports `sys: cpu`, `sys: mem`, `sys: disk` for individual metrics. Refreshes every 2 seconds.

### 📋 Todo Manager

The [quicksheet-todo](https://github.com/cemheren/quicksheet-todo) extension brings task management into your cells — add tasks with priorities and due dates, track completion:

```
ext: github:cemheren/quicksheet-todo
todo: add !high @2026-05-20 Fix login bug
todo: list
```

Supports 4 priority levels (`!low`, `!normal`, `!high`, `!critical`), due date tracking with overdue warnings, and persistent storage across sessions.

### Example: Calendar events

The [quicksheet-cal](https://github.com/cemheren/quicksheet-cal) extension reads `.ics` calendar files and shows upcoming events grouped by date — a glanceable schedule on your desktop:

```
ext: github:cemheren/quicksheet-cal
cal: ~/calendar.ics
cal: today
cal: week
```

Parses standard iCalendar (RFC 5545) files. Auto-scans common calendar directories (Evolution, Thunderbird, KDE, Calcurse) when no path is given.

### Example: Stock ticker

The [quicksheet-stock-ext](https://github.com/cemheren/quicksheet-stock-ext) extension pulls live stock quotes from Stooq — no API key needed. A column of `stock:` cells is a watchlist that lives on the wallpaper:

```
ext: github:cemheren/quicksheet-stock-ext
stock: AAPL, 1, 3
```

US tickers default to `.us`. For other markets: `stock: bp.uk`, `stock: 7203.jp`, `stock: spy.us`.

### Example: HTTP ping monitor

The [quicksheet-ping-ext](https://github.com/cemheren/quicksheet-ping-ext) extension turns cells into a no-config status page — HTTP status code and latency at a glance:

```
ext: github:cemheren/quicksheet-ping-ext
ping: https://example.com, 1, 3
```

Shows `✓` for 2xx/3xx, `⚠` for 4xx, `✗` for 5xx. Pair with `L: <cell>, 1m` for a one-minute uptime poll on the wallpaper.

### Example: Self-employment tax estimator

The [quicksheet-1099-ext](https://github.com/cemheren/quicksheet-1099-ext) extension estimates US self-employment tax — pure math, no network:

```
ext: github:cemheren/quicksheet-1099-ext
1099: 80000, 1, 5
```

Shows SE tax estimate, quarterly payment amount, and a reminder that it's not tax advice.

### Example: Gravatar lookup

The [quicksheet-grav-ext](https://github.com/cemheren/quicksheet-grav-ext) extension looks up Gravatar profiles by email — display name, location, and avatar URL:

```
ext: github:cemheren/quicksheet-grav-ext
grav: jane@example.com, 1, 4
```

### Example: Thesaurus

The [quicksheet-thes-ext](https://github.com/cemheren/quicksheet-thes-ext) extension provides inline synonyms via the free Datamuse API — pairs nicely with the dictionary extension:

```
ext: github:cemheren/quicksheet-thes-ext
thes: laconic, 1, 6
```

### Example: DOI citation lookup

The [quicksheet-cite-ext](https://github.com/cemheren/quicksheet-cite-ext) extension resolves DOIs to author/title/venue citations via Crossref — no browser round-trip needed:

```
ext: github:cemheren/quicksheet-cite-ext
cite: 10.1145/3623476.3623525, 1, 4
```

Accepts plain DOIs, `doi:...` form, or full `https://doi.org/...` URLs.

### Example: MX record checker

The [quicksheet-mxck-ext](https://github.com/cemheren/quicksheet-mxck-ext) extension looks up MX records via Google DNS-over-HTTPS — handy for email debugging:

```
ext: github:cemheren/quicksheet-mxck-ext
mxck: example.com, 1, 5
```

### Example: Budget envelope tracker

The [quicksheet-budget](https://github.com/cemheren/quicksheet-budget) extension visualizes spending per category with progress bars — a glanceable budget dashboard on your desktop:

```
ext: github:cemheren/quicksheet-budget
budget: Groceries, 800, 623
budget: Software, 200, 89
```

Shows color-coded status (🟢≤50%, 🟡≤75%, 🟠≤90%, 🔴>90%), visual fill bar, and remaining balance. Pairs with cell references for live updates as you log expenses.

### Build your own

Extensions are regular .NET (or any language) programs that read/write JSON lines on stdin/stdout. The protocol is intentionally minimal:

1. QuickSheet sends `{"type":"init"}` → extension replies with `{"type":"register", "prefix":"xyz", ...}`
2. When a user activates a cell matching the prefix, QuickSheet sends `{"type":"activate", ...}` → extension replies with `{"type":"write", "cells":[...]}` to fill the grid.

Ship a `quicksheet-extension.json` manifest in your repo and you're done. See the [weather extension](https://github.com/cemheren/quicksheet-weather) for a minimal working example, or the [Copilot extension](https://github.com/cemheren/quicksheet-copilot-ext) for something more advanced.

## Keyboard Shortcuts

### Navigation

| Shortcut | Action |
|----------|--------|
| Arrow Keys | Navigate cells |
| Shift+Arrow | Extend selection / multi-select |
| Tab | Move right, wrap to next row |
| Enter | Activate cell (open URL / run command) |

### Editing

| Shortcut | Action |
|----------|--------|
| F2 | Edit cell in-place |
| Backspace | Delete last character / clear selection |
| Delete | Clear cell or selection |
| Escape | Cancel edit |
| Type any character | Start editing the selected cell |

### View & Layout

| Shortcut | Action |
|----------|--------|
| F1 | Toggle raw vs resolved cell display |
| F3 | Rebuild grid / recalculate layout |
| F4 | Decrease column width |
| F5 | Increase column width |

### Clipboard & Search

| Shortcut | Action |
|----------|--------|
| Ctrl+C | Copy cell(s) |
| Ctrl+Shift+C | Copy resolved / displayed content |
| Ctrl+X | Cut cell(s) |
| Ctrl+V | Paste (supports multi-line) |
| Ctrl+F | Search |

### Undo & History

| Shortcut | Action |
|----------|--------|
| Ctrl+Z | Undo last action |
| Ctrl+Y | Redo |

### Appearance

| Shortcut | Action |
|----------|--------|
| Ctrl+T | Cycle theme (Dark → Light → Nord → Solarized → Matrix) |

### File & Row Operations

| Shortcut | Action |
|----------|--------|
| Ctrl+S | Save to CSV |
| Ctrl+D | Delete row |
| Ctrl+O | Insert row below |
| Ctrl+P | Insert row above |
| Ctrl+H | Show help overlay |
| Ctrl+Q | Quit |

## License

MIT
