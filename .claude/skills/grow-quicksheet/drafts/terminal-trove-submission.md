# Terminal Trove Submission Draft

> Submit at: https://terminaltrove.com/post/

## Fields

**Name:** QuickSheet

**URL:** github.com/cemheren/QuickSheet

**Tagline (≤100 chars):**
A zero-dependency .NET spreadsheet that embeds as your Linux/Windows desktop wallpaper.

**Description (250-300 chars):**
QuickSheet turns your desktop background into a live, editable spreadsheet. Track tasks, run inline shell commands, monitor services—all visible behind your icons with zero alt-tab. Ships as a single binary with no NuGet dependencies. Supports CSV, formulas, themes, and 85+ extensions.

**Feature 1 (150-300 chars):**
Desktop wallpaper mode — the grid renders behind your desktop icons using X11 (Linux) or WorkerW embedding (Windows). Always visible, never in the way. Your data lives where you look all day.

**Feature 2 (150-300 chars):**
Inline process output — prefix any cell with `i:` to run a shell command and stream its stdout directly into the cell. Monitor system resources, git status, or API responses without leaving the spreadsheet.

**Feature 3 (150-300 chars):**
Extension ecosystem — 85+ community extensions (weather, stocks, Docker status, DNS lookups, GitHub streak) via a simple JSON-lines protocol. Install with one line: `ext: github:user/repo`.

**Primary Language:** C#

**License:** MIT

**Categories:** TUI, Data, Productivity, System

**Are you the author?** Yes

**Preview Image:** [PLACEHOLDER — user must capture a wallpaper screenshot showing the grid behind desktop icons, ideally with 2-3 inline process cells active]

---

## Console.dev Email Draft

**To:** team@console.dev
**Subject:** Tool submission: QuickSheet — desktop-wallpaper spreadsheet (Linux/Windows)

Hi Console team,

I'd like to submit QuickSheet for your newsletter consideration.

**One-line pitch:** A zero-dependency .NET spreadsheet that replaces your desktop wallpaper — your data lives behind your icons, always visible, zero alt-tab.

**Why it's interesting for your audience:**
- Novel UX: embeds as the actual desktop background via X11 `_NET_WM_WINDOW_TYPE_DESKTOP` (Linux) or Win32 WorkerW (Windows). Not a floating widget — it IS the wallpaper.
- Inline shell execution: `i: docker ps`, `i: curl -s api.example.com/health` — cells that stream live process output.
- 85+ extensions via JSON-lines protocol, installable with `ext: github:user/repo`.
- Zero external dependencies. Single binary. MIT licensed.

**Links:**
- GitHub: https://github.com/cemheren/QuickSheet
- [Screenshot/GIF: PLACEHOLDER]

Thanks for considering it!

---

**Note to user:** Both submissions require a wallpaper screenshot or GIF showing
the grid behind desktop icons. Do not submit until that asset exists.
