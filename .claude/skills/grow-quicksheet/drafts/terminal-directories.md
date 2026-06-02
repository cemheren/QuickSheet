# Terminal tool directory submissions — drafts (2026-06-01)

> User submits manually via each site's form. Do not post on behalf of the user.

---

## 1. TerminalTrove (terminaltrove.com/post/)

### Form fields

| Field | Value |
|-------|-------|
| **name** | QuickSheet |
| **url** | github.com/cemheren/QuickSheet |
| **tagline** | Interactive CSV spreadsheet that runs as your desktop wallpaper — cells run commands, stream output, render sparklines. |

**Description (250–300 chars):**

> QuickSheet is an open-source .NET 9 spreadsheet that embeds behind your desktop icons. Your wallpaper becomes a live, interactive grid. Type notes, paste URLs (auto-linked), run shell commands from cells, or install extensions with one cell. CSV is the file format — everything round-trips through standard tools.

**2–3 standout features (150–300 chars):**

> 1) `r: cmd` cells execute shell commands on Enter. 2) `i: cmd` cells stream live subprocess output (htop, tail -f, sensors). 3) 40+ one-cell installable extensions via `ext: github:user/repo` (stocks, weather, dice, citations, unit conversion).

**Other notable features (150–300 chars):**

> Sparkline-in-cell rendering (`s: 3,7,2,9`). Inline cell-range references (`{A1::C10}` embeds another region). Column sums (Σ) and row products (Π). CSV autosave every 5 seconds. Headless export: `--export-md`, `--export-tsv`. Zero NuGet dependencies — all platform interop is hand-written P/Invoke.

**Who is this for / when to use it (150–250 chars):**

> Power users who want ambient data on their desktop without a widget engine. Homelabbers monitoring services, students tracking coursework, SREs watching endpoints, traders with a live watchlist — anyone who wants a scratchpad that's always visible.

**Technical details:**
- Primary language: **c#**
- License: **mit**

**Categories (select all that apply):**
- `cli`, `tui`, `productivity`, `cross-platform`, `utilities`, `text-processing`

**Operating Systems:**
- `linux`, `windows`

**Install instructions:**

```
# .NET 9 SDK required
git clone https://github.com/cemheren/QuickSheet.git
cd QuickSheet
dotnet run -c Release --project ExcelConsole.csproj
```

(When a GitHub Release with binaries lands, update to: download from Releases page.)

**Preview image:** User needs to provide a PNG screenshot of QuickSheet running on desktop + optionally a GIF showing cell editing.

**Are you the author?** yes

### Notes for the submitter

- TerminalTrove has >1,000 tools listed. Getting accepted brings long-tail discovery via their search + category pages + weekly newsletter mentions.
- They've listed .NET tools before (e.g., `spectre.console`). Being niche isn't disqualifying.
- The "post criteria" link on the form should be read before submitting — skim for any disqualifying rules.
- If they ask for a Homebrew/apt/scoop install command, say "binary releases coming soon" or omit until the release workflow lands.

---

## 2. Console.dev (console.dev/submit/)

### Submission form fields

| Field | Value |
|-------|-------|
| **Name** | QuickSheet |
| **URL** | https://github.com/cemheren/QuickSheet |
| **Short description** | An interactive CSV spreadsheet that replaces your desktop wallpaper. Cells can run shell commands, stream live subprocess output, render sparklines, and install extensions from GitHub — all with zero dependencies. .NET 9, MIT, cross-platform (Windows + Linux X11). |
| **Contact email** | cemheren@gmail.com |
| **Demo link** | *(leave blank or link to README screencast if one exists)* |
| **GitHub** | https://github.com/cemheren/QuickSheet |

### Pitch angle for console.dev's editorial voice

Console.dev writes short "What we like / What we don't like" reviews. Anticipate:

**What they'd like:**
- Novel concept: wallpaper as interactive grid, not just display widgets.
- Zero dependencies — refreshing in 2026's dep-heavy ecosystem.
- Extension system that installs from a GitHub repo in one cell (`ext: github:user/repo`).
- CSV as the universal interchange format — data isn't locked in.
- Cross-platform with the same codebase (.NET conditional compilation, no Electron).

**What they'd flag as "don't like":**
- No prebuilt binaries yet (requires .NET SDK to build from source). ← mitigated once release workflow merges.
- Linux requires X11 (no Wayland support).
- Small community / low star count at time of submission.
- No macOS support.

### Notes for the submitter

- Console.dev features ~6 tools per weekly newsletter (30k+ subscribers). Getting picked = significant spike.
- They lean toward tools with polished landing pages and easy install. Submitting after the GitHub Release workflow merges (prebuilt binaries) significantly increases acceptance odds.
- Recommended timing: submit 1–2 weeks after the first tagged release with binaries is live.
- Tone: developer-first, no hype. Lead with the "wallpaper as spreadsheet" hook — it's visually distinct from anything else they've covered.

---

## Submission timing strategy

1. **TerminalTrove** — submit now (or after first binary release). Lower bar, longer tail, broader audience.
2. **Console.dev** — submit after `v1.0.0` release with binaries. Higher bar, but 30k-subscriber newsletter = large single-day spike potential.

Both are independent of each other and of awesome-list submissions. Can do all in parallel.

## Expected impact

- TerminalTrove: 5–20 stars in long-tail traffic over months (search, category browsing, RSS).
- Console.dev: 50–200 stars if featured in main newsletter (based on comparable .NET/TUI tool features).
- Combined: establishes presence in two key discovery channels for terminal tools.
