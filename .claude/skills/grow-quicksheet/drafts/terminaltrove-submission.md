# Terminal Trove submission — draft (2026-06-07)

> User submits manually at https://terminaltrove.com/post/. Do not post anywhere.

## Submission fields

**Name:** QuickSheet

**URL:** github.com/cemheren/QuickSheet

**Tagline (≤100 chars):**

A transparent spreadsheet grid that replaces your desktop wallpaper.

**Description (250–300 chars):**

QuickSheet embeds an interactive CSV-backed grid behind your windows using platform-native tricks (WorkerW on Windows, X11 _NET_WM_WINDOW_TYPE_DESKTOP on Linux). Cells run shell commands, stream live subprocess output, render sparklines, and query 80+ community extensions — all from your desktop background.

**2–3 standout features (150–300 chars each):**

1. **Live subprocess cells** — prefix a cell with `i: htop` or `i: docker stats` and get continuously-updating output rendered inline, directly on your wallpaper.

2. **80+ extensions via JSON-lines protocol** — weather, stocks, RSS feeds, Docker status, TLS cert checks, GitHub Actions, timezones, and more. Install with one line: `ext: github:user/repo`.

3. **Zero dependencies** — no NuGet packages, no Electron, no browser. Clone the repo, `dotnet build`, run. The entire supply chain is the .NET 9 SDK.

**Additional notable features:**

- Plain CSV persistence — autosaves every 5 seconds to a file you can version-control
- Cross-platform: Windows (embedded behind desktop icons) + Linux (X11 desktop window type)
- Cell types: shell commands (`r:`), inline processes (`i:`), sparklines (`s:`), hyperlinks, cell-range references (`{A1::C10}`)
- Keyboard-driven: vim-inspired navigation, search, command palette
- Headless export: `--export-md`, `--export-json` for CI pipelines

**Who is this for / when to use it (150–250 chars):**

Developers and power users who want persistent, glanceable information on their desktop without opening windows. Perfect for homelab dashboards, quick-reference notes, SRE status boards, or daily task tracking.

**Primary language:** C#

**License:** MIT

**Categories (select on form):** Desktop, Productivity, DevOps (closest matches)

**Preview image:** ⚠️ HUMAN ACTION REQUIRED — Terminal Trove requires a PNG, GIF, or MP4 preview. Use an existing screenshot from the repo or capture a new one showing the grid on a desktop with live extension output.

**Install instructions:**

```
# Prerequisites: .NET 9 SDK

# Clone and build
git clone https://github.com/cemheren/QuickSheet.git
cd QuickSheet
dotnet build -c Release ExcelConsole.csproj

# Run (Linux)
dotnet run -c Release --project ExcelConsole.csproj

# Run with a CSV
dotnet run -c Release --project ExcelConsole.csproj -- data.csv
```

**Are you the author?** Yes (check box — user is the author)

**Read posting criteria?** Yes (check box)

## Notes for the submitter

- Terminal Trove is specifically curated for terminal/CLI/TUI tools. QuickSheet fits because it's developer-focused, keyboard-driven, and built with raw platform APIs (no GUI framework).
- The preview image/GIF is **mandatory** — submission will be rejected without one. A 5-second GIF showing navigation + a live `i:` cell updating is ideal.
- Terminal Trove features a "Tool of the Week" spotlight that drives significant traffic. Being a .NET project makes it novel in their largely Go/Rust catalog.
- Their audience skews heavily toward Linux CLI users and r/unixporn enthusiasts — emphasize the X11 integration and keyboard-driven workflow.
- Cross-link with the Console.dev submission for compounding visibility.

## Positioning vs. similar tools on Terminal Trove

- **VisiData** — read-only data exploration; QuickSheet is read-write with live cells
- **sc-im** — terminal-only spreadsheet; QuickSheet is wallpaper-embedded
- **Conky** — static widget renderer; QuickSheet has editable cells + extension protocol

## After submission

- If accepted: permanent listing on terminaltrove.com with backlink to GitHub repo
- Potential "Tool of the Week" feature → 500–2000 visitors from newsletter
- Terminal Trove also has a YouTube channel where they demo tools — GIF/video increases chances
- Realistic outcome: 20–80 stars if featured as Tool of the Week; 5–15 from catalog listing alone
