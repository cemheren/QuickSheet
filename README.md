# QuickSheet

**Your desktop is a spreadsheet.**

QuickSheet replaces your wallpaper with a transparent, interactive grid. Pin notes, launch apps, paste links, track numbers — all without opening a window. The idea: keep something lightweight always present in the background, instead of a static wallpaper you never interact with.

The data is a CSV file. Cells can run shell commands. Same file on Windows or Linux.

![QuickSheet running as the desktop wallpaper — cells holding runnable commands behind every open window](desktop-wallpaper-commands.png)

![.NET 9](https://img.shields.io/badge/.NET-9.0-purple)
![License](https://img.shields.io/badge/license-MIT-blue)
![Windows](https://img.shields.io/badge/platform-Windows-0078D6?logo=windows)
![Linux](https://img.shields.io/badge/platform-Linux-FCC624?logo=linux&logoColor=black)
![GitHub release](https://img.shields.io/github/v/release/cemheren/QuickSheet?color=green)
![Zero Dependencies](https://img.shields.io/badge/dependencies-0-brightgreen)
![Extensions](https://img.shields.io/badge/extensions-69%2B-orange)
![Cell prefixes](https://img.shields.io/badge/cell_prefixes-9-blue)

## Why this exists

Most developers have a second monitor — or at least a desktop — that shows a static wallpaper doing nothing. QuickSheet turns that dead space into something useful:

- **Always-on scratchpad.** Click anywhere on the desktop to jot a note. No window to find, no app to open. Autosaves every 5 seconds.
- **App launcher.** Prefix a cell with `r: code .` and hit Enter. Multi-select cells to launch your whole morning stack in one keystroke.
- **Link dashboard.** Paste URLs into cells. They're highlighted and open on Enter — a personal start page that lives behind your windows.
- **Live data.** Column sums, sparklines, inline subprocesses (`i: top`), and 69+ extensions for weather, stocks, RSS, system monitoring, and more.
- **Zero dependencies.** Clone → `dotnet build` → run. No NuGet packages, no npm, no Docker. The entire supply chain is the .NET SDK.

If you spend your day in a terminal or IDE and want your desktop to *do* something, QuickSheet is for you.

## Quick Start

```bash
git clone https://github.com/cemheren/QuickSheet.git
cd QuickSheet
dotnet run -c Release --project ExcelConsole.csproj -- --desktop
```

Requires [.NET 9 SDK](https://dotnet.microsoft.com/download/dotnet/9.0). Release mode is recommended — it feels noticeably snappier.

Tip: add it to your startup applications so your notes, links, and launchers are there every time you log in. See [docs/install-startup.md](docs/install-startup.md) for Windows (Task Scheduler / Startup folder + an opt-in `-Update` flag) and Linux (XDG autostart) walkthroughs.

Want to export your data? Render any CSV as a GitHub-flavored Markdown table:

```bash
dotnet run --project ExcelConsole.csproj -- mydata.csv --export-md mydata.md
```

Use `--export-md -` to write to stdout — handy in csvkit / Miller / qsv pipelines. See [docs/csvkit-comparison.md](docs/csvkit-comparison.md) for the drop-in pipeline notes.

New here? The 60-second tour is in [docs/tour.md](docs/tour.md). Or jump straight to [docs/recipes.md](docs/recipes.md) for ready-to-paste dashboard layouts. All keyboard shortcuts are in [docs/keyboard-shortcuts.md](docs/keyboard-shortcuts.md). Common questions: [docs/faq.md](docs/faq.md).

Using Claude Code, Copilot CLI, or Aider? `dotnet run --project ExcelConsole.csproj -- examples/ai-workflow.csv --desktop` opens a pre-seeded panel of the top AI-CLI launchers (`claude --continue`, `gh copilot suggest`, `aider --ask`), a `/cost` token sparkline, and the copilot-in-a-cell extension hook.

Audience-specific guides: [for homelabbers](docs/for-homelab.md) (Plex / Pi-hole / *arr / Home Assistant) · [for traders](docs/for-traders.md) (P/L, watchlist, FX, news strip) · [for SREs & DevOps](docs/for-sre.md) (service health, k8s pods, TLS expiry) · [for students](docs/for-students.md) (coursework, budget, developer tools).

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

### Section headers in a cell
Prefix a cell with `# ` (H1) or `## ` (H2) to create a visually distinct section header.
- `# Work` → rendered in **gold** — ideal for top-level dashboard sections
- `## Servers` → rendered in **cyan** — ideal for sub-sections

The prefix is hidden during display; only the label is shown. Combine with color-coded cells and links to build organized, scannable desktop dashboards.

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

The [quicksheet-copilot-ext](https://github.com/Deskworks/quicksheet-copilot-ext) extension brings AI directly into your spreadsheet grid. Ask questions, generate structured data, or summarize cell ranges — all without leaving your desktop.

```
ext: github:Deskworks/quicksheet-copilot-ext
copilot: test, 1, 1
```

![Copilot extension running on the desktop](copilot-ext-example.png)

The `copilot:` prefix takes a natural-language prompt followed by the number of columns and rows you want written to the grid:

```
copilot: <your prompt>, <columns>, <rows>
```

**Use cases:**

| Use case | Example cell |
|---|---|
| Generate test data | `copilot: 10 user records with name email city, 4, 10` |
| Compare technologies | `copilot: compare React Vue Angular on performance learning-curve, 3, 4` |
| Create quiz questions | `copilot: 5 Python beginner questions with answers, 2, 5` |
| Summarize your notes | `copilot: summarize key points from {A1::A20}, 1, 5` |
| Sprint planning agenda | `copilot: 8 sprint planning agenda items as task and duration, 2, 8` |
| List best practices | `copilot: 6 REST API best practices as rule and reason, 2, 6` |
| Debug cheat sheet | `copilot: common Python errors with cause and fix, 2, 8` |
| Salary benchmarks | `copilot: software engineer salaries by level junior mid senior, 2, 5` |

Cell references like `{A1::C10}` are expanded before being sent to Copilot, so it receives the actual cell contents as context — useful for summarization, analysis, or transformation of existing data.

A ready-made multi-section dashboard is available in [`examples/copilot-dashboard.csv`](examples/copilot-dashboard.csv) — open it with `quicksheet copilot-dashboard.csv` to see data generation, research, learning, planning, and finance sections all in one grid.

### Example: Weather forecast widget

The [quicksheet-weather](https://github.com/Deskworks/quicksheet-weather) extension turns a cell into a live 7-day weather forecast:

```
ext: github:Deskworks/quicksheet-weather
wthr: Seattle, 2, 7
```

![Weather extension showing a 7-day forecast on the desktop](weather-ext-example.png)

### Example: TLS certificate checker

The [quicksheet-tls-ext](https://github.com/Deskworks/quicksheet-tls-ext) extension shows TLS cert expiry and issuer — useful as an ambient SRE dashboard:

```
ext: github:Deskworks/quicksheet-tls-ext
tls: github.com, 1, 4
```

### More extensions

| Prefix | What it does | Install |
|--------|-------------|---------|
| `price:` | Crypto prices (CoinGecko) | `ext: github:Deskworks/quicksheet-price-ext` |
| `def:` | Dictionary lookups | `ext: github:Deskworks/quicksheet-define-ext` |
| `mort:` | Mortgage calculator | `ext: github:Deskworks/quicksheet-mortgage-ext` |
| `stock:` | Stock quotes (Stooq) | `ext: github:Deskworks/quicksheet-stock-ext` |
| `cal:` | Calendar events (.ics) | `ext: github:Deskworks/quicksheet-cal` |
| `todo:` | Task management | `ext: github:Deskworks/quicksheet-todo` |
| `ping:` | HTTP status & latency | `ext: github:Deskworks/quicksheet-ping-ext` |
| `fx:` | Currency conversion | `ext: github:Deskworks/quicksheet-fx` |
| `gha:` | GitHub Actions status | `ext: github:Deskworks/quicksheet-gha` |
| `ghpr:` | GitHub PR dashboard | `ext: github:Deskworks/quicksheet-ghpr` |
| `docker:` | Container health | `ext: github:Deskworks/quicksheet-docker` |
| `gitst:` | Git repo status | `ext: github:Deskworks/quicksheet-gitst` |
| `portck:` | TCP port checker | `ext: github:Deskworks/quicksheet-portck` |
| `k8s:` | Kubernetes pod status | `ext: github:Deskworks/quicksheet-k8s` |
| `cntdn:` | Countdown timers | `ext: github:Deskworks/quicksheet-cntdn` |
| `1099:` | US self-employment tax | `ext: github:Deskworks/quicksheet-1099-ext` |
| `mileage:` | IRS mileage deduction | `ext: github:Deskworks/quicksheet-mileage-ext` |
| `budget:` | Budget envelopes | `ext: github:Deskworks/quicksheet-budget` |
| `worldtm:` | World clock / timezones | `ext: github:Deskworks/quicksheet-worldtm` |
| `margin:` | Break-even & margin calc | `ext: github:Deskworks/quicksheet-margin-ext` |
| `depr:` | Depreciation schedules | `ext: github:Deskworks/quicksheet-depr-ext` |
| `jwtdec:` | JWT token decoder | `ext: github:Deskworks/quicksheet-jwtdec` |
| `cronck:` | Cron expression parser | `ext: github:Deskworks/quicksheet-cronck` |
| `guid:` | GUID/UUID generator | `ext: github:Deskworks/quicksheet-guid` |
| `regex:` | Regex pattern explainer | `ext: github:Deskworks/quicksheet-regex` |
| `urlenc:` | URL encode/decode | `ext: github:Deskworks/quicksheet-urlenc` |
| `curl:` | HTTP client (Postman-in-a-cell) | `ext: github:Deskworks/quicksheet-curl` |
| `arxiv:` | arXiv paper lookup | `ext: github:Deskworks/quicksheet-arxiv` |
| `pihole:` | Pi-hole DNS stats | `ext: github:Deskworks/quicksheet-pihole` |
| `health:` | HTTP endpoint health checker | `ext: github:Deskworks/quicksheet-health` |
| `env:` | Env var inspector (lookup/filter/PATH, auto-masks secrets) | `ext: github:Deskworks/quicksheet-envck` |
| `roll:` | Dice roller — 2d6+3, d20, 4d6kh3, Fudge, critical detection | `ext: github:Deskworks/quicksheet-dice` |
| `lc:` | LeetCode — daily challenge, problem lookup, user stats | `ext: github:Deskworks/quicksheet-leetcode` |
| `ghst:` | GitHub contribution streak — current streak, longest, sparkline | `ext: github:Deskworks/quicksheet-ghstreak` |
| `iss:` | Live ISS position + people in space — lat/lon, region, crew by craft | `ext: github:Deskworks/quicksheet-iss` |
| `npm:` | npm package info — version, weekly downloads, license, author; multi-package compare | `ext: github:Deskworks/quicksheet-npm` |
| `pypi:` | PyPI package info — version, license, author, Python requirement, release date | `ext: github:Deskworks/quicksheet-pypi` |
| `crates:` | Rust crate lookup — `crates: tokio` (version, downloads, description, keywords), `crates: search async` (top 5). Via crates.io, no API key. | `ext: github:Deskworks/quicksheet-crates` |
| `ghtrend:` | GitHub trending repos — stars, forks, language, description by language or overall | `ext: github:Deskworks/quicksheet-gh-trends` |
| `co2:` | Live atmospheric CO₂ from NOAA Mauna Loa — ppm, baseline, YoY, 30-day trend | `ext: github:Deskworks/quicksheet-co2` |
| `weather:` | Live weather via Open-Meteo — temperature, conditions, wind, humidity, pressure (free, no API key) | `ext: github:Deskworks/quicksheet-openmeteo` |
| `ip:` | IP geolocation — ISP, org, ASN, timezone, hostname, coords for any IP or your own | `ext: github:Deskworks/quicksheet-ipinfo` |
| `dns:` | DNS record lookup — A, AAAA, MX, TXT, NS, CNAME, PTR. `dns: MX gmail.com` · `dns: PTR 8.8.8.8` | `ext: github:Deskworks/quicksheet-dns` |
| `ssl:` | SSL certificate expiry checker — status 🟢/🟡/🔴, days left, issuer. `ssl: github.com` · `ssl: api.internal:8443` | `ext: github:Deskworks/quicksheet-ssl` |
| `whois:` | WHOIS domain/IP lookup — registrar, status, expiry urgency 🟡🔴, nameservers. `whois: github.com` · `whois: 8.8.8.8` | `ext: github:Deskworks/quicksheet-whois` |
| `tracert:` | Network traceroute — each hop with RTT and reverse-DNS hostname. `tracert: github.com` · `tracert: 8.8.8.8` | `ext: github:cemheren/quicksheet-tracert` |
| `mtr:` | Route tracer with per-hop RTT stats and packet loss. `mtr: google.com` (full route) · `mtr: ping 8.8.8.8` (ping stats: min/avg/max, loss%). | `ext: github:Deskworks/quicksheet-mtr` |
| `speed:` | Network speed test — download Mbps from CDN endpoints + HTTP latency to Cloudflare/Google/GitHub. `speed:` · `speed: download` · `speed: latency` | `ext: github:Deskworks/quicksheet-speedtest` |
| `epoch:` | Unix timestamp converter — `epoch: now` · `epoch: 1700000000` · `epoch: 2025-01-01` · `epoch: diff t1 t2` · `epoch: add <ts> 30d`. Zero network. | `ext: github:Deskworks/quicksheet-epoch` |
| `color:` | Color converter + palette — `color: #ff6600` · `color: rgb(255,102,0)` · `color: orange` · `color: palette #3b82f6`. HEX↔RGB↔HSL↔CMYK, WCAG contrast. Zero network. | `ext: github:Deskworks/quicksheet-color` |
| `caniuse:` | Browser compat lookup via caniuse-db — `caniuse: css-grid` · `caniuse: flexbox` · `caniuse: webassembly`. Shows support across Chrome/Firefox/Safari/Edge + usage %. No API key. | `ext: github:Deskworks/quicksheet-caniuse` |
| `gem:` | Ruby gem lookup via rubygems.org — version, total downloads, authors, license, homepage, description. `gem: rails` (detail) · `gem: search json parser` (top 5). No API key. | `ext: github:Deskworks/quicksheet-rubygems` |

See [docs/extensions.md](docs/extensions.md) for the full directory with descriptions and protocol docs.

### Build your own

Extensions are regular .NET (or any language) programs that read/write JSON lines on stdin/stdout. The protocol is intentionally minimal:

1. QuickSheet sends `{"type":"init"}` → extension replies with `{"type":"register", "prefix":"xyz", ...}`
2. When a user activates a cell matching the prefix, QuickSheet sends `{"type":"activate", ...}` → extension replies with `{"type":"write", "cells":[...]}` to fill the grid.

Ship a `quicksheet-extension.json` manifest in your repo and you're done. See the **[full protocol spec](docs/extension-protocol.md)** for message formats, coordinate system, rules, and working examples in C# and Python. Or check the [weather extension](https://github.com/Deskworks/quicksheet-weather) for a minimal working example.

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
| Ctrl+T | Cycle theme (Dark → Light → Nord → Solarized → SolarizedLight → Matrix → Dracula → Synthwave → Gruvbox → Monokai → HotdogStand) |

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
