# QuickSheet in 60 seconds

A guided tour of what makes QuickSheet different from "yet another TUI spreadsheet". Read this once and you've seen everything that matters.

> If you haven't installed it yet: see the [Quick Start](../README.md#quick-start). It's `dotnet run -c Release --project ExcelConsole.csproj -- --desktop`.

## 1. It's a desktop wallpaper

QuickSheet replaces your wallpaper with a transparent, click-through-aware grid. Your notes, launchers, and live data sit *behind* every window, always at hand, never stealing focus, until you click the grid to start typing.

## 2. Cell prefixes are the whole feature set

Every cell is plain text by default. Add a prefix and the cell *does something*.

| Prefix              | What it does                                                      | Example                            |
|---------------------|-------------------------------------------------------------------|------------------------------------|
| (none)              | Plain text. Autosaves every 5 seconds.                            | `groceries`                        |
| `r: `               | Runnable command. Press Enter to launch.                          | `r: code .`                        |
| `i: `               | Inline subprocess. Output streams back into the cell, live.       | `i: ping -c 1 example.com`         |
| `s: `               | Sparkline. Numbers render as unicode bars (▁▂▃▄▅▆▇█). Range form: `s: A1::A10`. | `s: 4,7,9,3,8,12`                  |
| `c:color: `         | Cell color. Highlights background with a named color.             | `c:red: URGENT`                    |
| `d: `               | Countdown to date. Shows days remaining/elapsed with label.       | `d: 2025-12-31 Release day`        |
| `L: `               | Loop a target cell on an interval.                                | `L: A10, 5m`                       |
| `ext: `             | Install an extension repo. One line.                              | `ext: github:Deskworks/quicksheet-weather` |
| `config: `          | Persistent settings (currently `theme=…`). Auto-created when you press `Ctrl+T`; survives restart via the CSV itself — no sidecar file. | `config: theme=Nord`               |
| `http://` `https://`| Hyperlink. Highlighted, opens on Enter.                           | `https://news.ycombinator.com`     |

Multi-select cells and hit Enter to fire all of them at once. That's how launching "my dev environment" works: pick the row of `r: ` cells, Enter.

## 3. Math without a calculator

The status bar shows:
- **Σ (sum)** of the column the cursor's in.
- **Π (product)** of the row.
- **File name**, **modified indicator** (●), and **non-empty cell count**.

No formulas to write. Type some numbers, glance at the bottom.

## 3½. Undo, redo, themes, go-to

A few keyboard shortcuts that round out the editing experience:

| Shortcut | What it does |
|----------|-------------|
| `Ctrl+Z` | Undo (up to 200 steps, action-grouped) |
| `Ctrl+Y` | Redo |
| `Ctrl+T` | Cycle theme: Dark, Light, Nord, Solarized, SolarizedLight, Matrix, Dracula, Synthwave, Gruvbox, Monokai, HotdogStand. Choice persists via a `config:` cell in the CSV. |
| `Ctrl+G` | Go-to cell — type a reference like `C5` and jump there |
| `Ctrl+H` | Help overlay with all shortcuts |

## 4. References

Wrap a range in `{A1::C10}` inside any text or runnable cell and QuickSheet inlines the live cell contents on activation. Combined with `r: ` and `i: `, this is how you build small per-row commands:

```
r: kubectl describe pod {A2::A2}
```

## 5. Extensions are git repos

`ext: github:user/repo` clones the repo, reads its `quicksheet-extension.json` manifest, and starts a subprocess that talks JSON-lines with QuickSheet. Then a new prefix appears.

Currently public:
- [`quicksheet-copilot-ext`](https://github.com/Deskworks/quicksheet-copilot-ext) — AI in a cell.
- [`quicksheet-weather`](https://github.com/Deskworks/quicksheet-weather) — 7-day forecast.
- [`quicksheet-tls-ext`](https://github.com/Deskworks/quicksheet-tls-ext) — TLS certificate checker.
- [`quicksheet-pomodoro`](https://github.com/Deskworks/quicksheet-pomodoro) — focus timer.
- [`quicksheet-price-ext`](https://github.com/Deskworks/quicksheet-price-ext) — crypto price quotes (CoinGecko).
- [`quicksheet-define-ext`](https://github.com/Deskworks/quicksheet-define-ext) — inline dictionary lookups.
- [`quicksheet-mortgage-ext`](https://github.com/Deskworks/quicksheet-mortgage-ext) — mortgage payment calculator.
- [`quicksheet-mxck-ext`](https://github.com/Deskworks/quicksheet-mxck-ext) — MX record check (DNS-over-HTTPS).
- [`quicksheet-ping-ext`](https://github.com/Deskworks/quicksheet-ping-ext) — HTTP status code + latency probe.
- [`quicksheet-cite-ext`](https://github.com/Deskworks/quicksheet-cite-ext) — DOI → citation (Crossref).
- [`quicksheet-thes-ext`](https://github.com/Deskworks/quicksheet-thes-ext) — thesaurus (Datamuse).
- [`quicksheet-stock-ext`](https://github.com/Deskworks/quicksheet-stock-ext) — stock ticker quotes (Stooq).
- [`quicksheet-1099-ext`](https://github.com/Deskworks/quicksheet-1099-ext) — US self-employment tax estimate.
- [`quicksheet-mileage-ext`](https://github.com/Deskworks/quicksheet-mileage-ext) — IRS standard-mileage deduction calculator.
- [`quicksheet-margin-ext`](https://github.com/Deskworks/quicksheet-margin-ext) — break-even point + contribution margin.
- [`quicksheet-depr-ext`](https://github.com/Deskworks/quicksheet-depr-ext) — straight-line + MACRS depreciation schedules.
- [`quicksheet-grav-ext`](https://github.com/Deskworks/quicksheet-grav-ext) — Gravatar profile + avatar URL.
- [`quicksheet-sysmon`](https://github.com/Deskworks/quicksheet-sysmon) — live CPU/RAM/disk/uptime monitor.
- [`quicksheet-todo`](https://github.com/Deskworks/quicksheet-todo) — task management with priorities + due dates.
- [`quicksheet-cal`](https://github.com/Deskworks/quicksheet-cal) — upcoming calendar events from .ics files.
- [`quicksheet-budget`](https://github.com/Deskworks/quicksheet-budget) — budget envelope visualizer with progress bars.
- [`quicksheet-qtr`](https://github.com/Deskworks/quicksheet-qtr) — IRS quarterly tax deadline countdown.
- [`quicksheet-fx`](https://github.com/Deskworks/quicksheet-fx) — live currency conversion (200+ currencies, ECB rates).
- [`quicksheet-rate`](https://github.com/Deskworks/quicksheet-rate) — freelance hourly rate calculator (taxes, benefits, billable time).

[Write your own](../README.md#build-your-own) in whatever language you want — the protocol is two message types.

## 5½. Headless mode & export

QuickSheet isn't just interactive — you can pipe data through it:

```bash
# Convert CSV to Markdown table (no UI launched)
dotnet run --project ExcelConsole.csproj -- data.csv --export-md output.md
```

This makes QuickSheet useful in scripts and CI pipelines, not just at the desktop.

## 6. The hard rules

- **Zero NuGet dependencies.** All native interop (X11, WinForms, ConPTY) is hand-written P/Invoke. Clone, build, run. No supply chain.
- **CSV is the format.** Open the file in Excel or vim. Same data.
- **Windows and Linux.** Windows uses WinForms + the WorkerW desktop trick. Linux uses raw X11 (`_NET_WM_WINDOW_TYPE_DESKTOP`). Wayland prints a warning.

## What this is good for

- A live dashboard of personal numbers (weights, runs, finances, todo counts) that doesn't require a webapp.
- A wallpaper-pinned launcher for the apps you open every day.
- An ambient SRE display: TLS expiries, ping latencies, alert counts.
- A scratchpad with computed columns that survives reboots as plain CSV.
- A freelancer finance dashboard: budget envelopes, quarterly tax deadlines, currency conversion, hourly rate calculator, and SE tax estimates — all in cells on your desktop.

## What this isn't

- An Excel replacement. No charts, pivot tables, or complex formulas.
- A database. CSV only.
- Polished. It's a side project. PRs and issues welcome — see [CONTRIBUTING.md](../CONTRIBUTING.md).
