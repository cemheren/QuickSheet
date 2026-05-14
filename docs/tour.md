# QuickSheet in 60 seconds

A guided tour of what makes QuickSheet different from "yet another TUI spreadsheet". Read this once and you've seen everything that matters.

> If you haven't installed it yet: see the [Quick Start](../README.md#quick-start). It's `dotnet run -c Release --project ExcelConsole.csproj -- --desktop`.

## 1. It's a desktop wallpaper

Run with `--desktop` and the grid becomes your wallpaper — transparent, click-through-aware, but interactive when you focus it. Your notes, launchers, and live data sit *behind* every window, always at hand, never stealing focus.

Drop the flag and you get a plain terminal TUI inside any shell. Same data, two surfaces.

## 2. Cell prefixes are the whole feature set

Every cell is plain text by default. Add a prefix and the cell *does something*.

| Prefix              | What it does                                                      | Example                            |
|---------------------|-------------------------------------------------------------------|------------------------------------|
| (none)              | Plain text. Autosaves every 5 seconds.                            | `groceries`                        |
| `r: `               | Runnable command. Press Enter to launch.                          | `r: code .`                        |
| `i: `               | Inline subprocess. Output streams back into the cell, live.       | `i: ping -c 1 example.com`         |
| `s: `               | Sparkline. Numbers render as unicode bars (▁▂▃▄▅▆▇█). Range form: `s: A1::A10`. | `s: 4,7,9,3,8,12`                  |
| `L: `               | Loop a target cell on an interval.                                | `L: A10, 5m`                       |
| `ext: `             | Install an extension repo. One line.                              | `ext: github:cemheren/quicksheet-weather` |
| `http://` `https://`| Hyperlink. Highlighted, opens on Enter.                           | `https://news.ycombinator.com`     |

Multi-select cells and hit Enter to fire all of them at once. That's how launching "my dev environment" works: pick the row of `r: ` cells, Enter.

## 3. Math without a calculator

The status bar shows:
- **Σ (sum)** of the column the cursor's in.
- **Π (product)** of the row.

No formulas to write. Type some numbers, glance at the bottom.

## 4. References

Wrap a range in `{A1::C10}` inside any text or runnable cell and QuickSheet inlines the live cell contents on activation. Combined with `r: ` and `i: `, this is how you build small per-row commands:

```
r: kubectl describe pod {A2::A2}
```

## 5. Extensions are git repos

`ext: github:user/repo` clones the repo, reads its `quicksheet-extension.json` manifest, and starts a subprocess that talks JSON-lines with QuickSheet. Then a new prefix appears.

Currently public:
- [`quicksheet-copilot-ext`](https://github.com/cemheren/quicksheet-copilot-ext) — AI in a cell.
- [`quicksheet-weather`](https://github.com/cemheren/quicksheet-weather) — 7-day forecast.
- [`quicksheet-tls-ext`](https://github.com/cemheren/quicksheet-tls-ext) — TLS certificate checker.
- [`quicksheet-pomodoro`](https://github.com/cemheren/quicksheet-pomodoro) — focus timer.
- [`quicksheet-price-ext`](https://github.com/cemheren/quicksheet-price-ext) — crypto price quotes (CoinGecko).
- [`quicksheet-define-ext`](https://github.com/cemheren/quicksheet-define-ext) — inline dictionary lookups.
- [`quicksheet-mortgage-ext`](https://github.com/cemheren/quicksheet-mortgage-ext) — mortgage payment calculator.
- [`quicksheet-mxck-ext`](https://github.com/cemheren/quicksheet-mxck-ext) — MX record check (DNS-over-HTTPS).
- [`quicksheet-ping-ext`](https://github.com/cemheren/quicksheet-ping-ext) — HTTP status code + latency probe.
- [`quicksheet-cite-ext`](https://github.com/cemheren/quicksheet-cite-ext) — DOI → citation (Crossref).
- [`quicksheet-thes-ext`](https://github.com/cemheren/quicksheet-thes-ext) — thesaurus (Datamuse).
- [`quicksheet-stock-ext`](https://github.com/cemheren/quicksheet-stock-ext) — stock ticker quotes (Stooq).

[Write your own](../README.md#build-your-own) in whatever language you want — the protocol is two message types.

## 6. The hard rules

- **Zero NuGet dependencies.** All native interop (X11, WinForms, ConPTY) is hand-written P/Invoke. Clone, build, run. No supply chain.
- **CSV is the format.** Open the file in Excel or vim. Same data.
- **Windows and Linux.** Windows uses WinForms + the WorkerW desktop trick. Linux uses raw X11 (`_NET_WM_WINDOW_TYPE_DESKTOP`). Wayland prints a warning.

## What this is good for

- A live dashboard of personal numbers (weights, runs, finances, todo counts) that doesn't require a webapp.
- A wallpaper-pinned launcher for the apps you open every day.
- An ambient SRE display: TLS expiries, ping latencies, alert counts.
- A scratchpad with computed columns that survives reboots as plain CSV.

## What this isn't

- An Excel replacement. No charts, pivot tables, or complex formulas.
- A database. CSV only.
- Polished. It's a side project. PRs and issues welcome — see [CONTRIBUTING.md](../CONTRIBUTING.md).
