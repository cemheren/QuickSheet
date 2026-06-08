# TerminalTrove submission — draft (2026-06-08)

> User submits manually at https://terminaltrove.com/post/ or emails hello@terminaltrove.com.
> Do not post anywhere.

## Submission form fields

**Tool name:** QuickSheet

**URL:** https://github.com/cemheren/QuickSheet

**Categories (select all that apply):**
- UI & Display
- Data & Text
- General

**Short description (1 line):**

A zero-dependency .NET spreadsheet that replaces your desktop wallpaper with an interactive, transparent grid.

## Email pitch (alternative to form)

**Subject:** New tool submission — QuickSheet (desktop wallpaper spreadsheet)

**Body:**

Hi Terminal Trove team,

I'd like to submit QuickSheet for listing.

**What it is:** QuickSheet replaces your desktop wallpaper with a transparent, interactive spreadsheet grid. It runs on Windows (WinForms/Win32 interop) and Linux (raw X11 P/Invoke) with zero external dependencies — the entire supply chain is the .NET 9 SDK.

**Why terminal/TUI folks would care:**

- Cells run shell commands (`r: code .`, `r: htop`), making it an always-visible app launcher.
- `i: <cmd>` prefix streams live subprocess output into cells (think `i: tail -f /var/log/syslog` or `i: sensors`).
- 69+ extensions via a JSON-lines protocol — weather, stocks, Docker stats, system monitoring, Git streak, RSS feeds, and more.
- Vim-style navigation, keyboard-driven, CSV persistence, autosaves every 5s.
- Sparkline rendering, column sums (Σ), row products (Π), cell-range references.

**Links:**
- GitHub: https://github.com/cemheren/QuickSheet
- Extension ecosystem: 45+ separate repos at https://github.com/cemheren?tab=repositories&q=quicksheet
- License: MIT

It's actively maintained (commits this week), cross-platform, and aims to turn dead desktop space into a useful always-on dashboard.

Thanks for considering it!

## Notes for user

- TerminalTrove has no documented quality bar — single-purpose TUIs get listed.
- They feature a "Tool of the Week" and "Newly Added" carousel.
- Tagging @terminaltrove on social after submission may increase visibility.
- Best submitted *after* a demo GIF exists in the README (visual sells).
