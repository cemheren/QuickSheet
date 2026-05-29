# Adjacent-Project Teardown (2026-05-29)

## Summary

Compared QuickSheet's README/positioning to three successful terminal tools:
**VisiData** (9k★, data explorer), **Harlequin** (7k★, SQL IDE), **Lazygit** (78k★, git UI).
Key patterns: animated GIF above the fold, one-line install, a "star popup" social nudge,
and HN launch timing. QuickSheet's README is already strong; the gap is social proof and
a single viral moment.

---

## VisiData (9k★) — Terminal tabular data explorer

**Repo:** https://github.com/saulpw/visidata

### README structure
- Title + version number + badge row (CI, Discord, Mastodon)
- **One-sentence pitch:** "A terminal interface for exploring and arranging tabular data."
- **Animated GIF immediately** — shows a frequency table being rearranged in real-time.
- Format support list (tsv, csv, sqlite, json, xlsx, hdf5…)
- Install: `pip3 install visidata` (one command)
- Usage: `vd <input>` — two characters
- Links to docs, tutorial, quick reference

### What works
1. **GIF is the hero.** The README has almost no prose before the GIF. Visitors "get it" in 3 seconds.
2. **Ultra-short install.** `pip3 install visidata` — no clone, no build step.
3. **Format breadth sells.** Leading with "supports csv, xlsx, json, sqlite, hdf5" signals versatility.
4. **Community signals:** Discord badge + Mastodon badge + Patreon link = "this is alive."

### What doesn't apply to QuickSheet
- VisiData is a data analysis tool; QuickSheet is a productivity/wallpaper tool. Different pitch axis.
- VisiData's install is trivial (pip); QuickSheet requires .NET SDK + clone. Can't compete on install friction without binary releases.

---

## Harlequin (7k★) — Terminal SQL IDE

**Repo:** https://github.com/tconbeer/harlequin

### README structure
- Title + PyPI badge + platform badge ("Runs on Linux | MacOS | Windows")
- **One-line pitch:** "The SQL IDE for Your Terminal."
- **SVG hero image** (not a GIF, a beautifully-rendered static screenshot in SVG format)
- Tip callout box linking to full docs
- Multi-method install (uv, pip, pipx, brew) — `uv tool install harlequin`
- Adapter/plugin system explanation
- Usage examples per database (DuckDB, SQLite, Postgres)
- Links to docs, discussions, issues, sponsorship

### What works
1. **Tagline is a genre label.** "The SQL IDE for Your Terminal" — 7 words, instantly understood. Establishes category.
2. **SVG screenshot** is pixel-perfect at any resolution. Looks professional in any viewer.
3. **Adapter architecture** makes the project feel bigger — "dozens of databases" via plugins.
4. **HN launch was strategic.** v1.0 announcement hit #2 on HN front page. The author waited until the product was polished before launching publicly.
5. **Docs site (harlequin.sh)** gives legitimacy beyond "just a README." The README is short and directs to docs.

### What applies to QuickSheet
- QuickSheet's extension system is analogous to Harlequin's adapters — "69+ extensions" is already in the badge. Good.
- The "tagline as genre label" pattern: QuickSheet's "Your desktop is a spreadsheet" is strong. ✓
- A docs site (even GitHub Pages) would increase perceived polish.
- Harlequin waited for polish before HN. QuickSheet should too — the current 0-star state means one shot matters.

---

## Lazygit (78k★) — Terminal Git UI

**Repo:** https://github.com/jesseduffield/lazygit

### README structure
- Title + massive animated GIF (full terminal interaction demo)
- Badges (CI, Go report, releases, Homebrew)
- Features list (staging lines, rebase, cherry-pick, merge conflicts)
- Install: multiple package managers (brew, choco, scoop, apt, go install)
- Usage: just `lazygit`
- Keybinding docs link

### What worked for growth
1. **Star popup on first run.** Lazygit shows a gentle nudge asking users to star the repo. Author credits this as a major growth driver.
2. **"Show HN" launch.** Creator posted when the tool was genuinely useful, not half-baked. The ensuing CLI-vs-GUI debate drove comments (engagement = visibility).
3. **GIF is spectacular.** Shows a complex rebase happening in seconds with keyboard-only input. Immediately conveys "this saves time."
4. **Cross-platform from day one.** Linux + macOS + Windows = maximum addressable audience.
5. **Simple mental model.** "It's git, but visual, in your terminal." No explanation needed.

### What applies to QuickSheet
- **Star popup:** QuickSheet could add a subtle first-run message ("If you find this useful, please ⭐ on GitHub"). Low-cost, high-upside. The author of lazygit credits this as the single biggest growth hack.
- **GIF quality matters enormously.** QuickSheet's current README has a static PNG. An animated GIF showing: type a note → launch an app → see live stock data → sparkline updating — would be transformative.
- **"Show HN" timing.** Don't launch on HN until: (a) demo GIF exists, (b) binary releases exist, (c) install is ≤2 steps. Currently QuickSheet requires clone + .NET SDK. A GitHub Release with pre-built binaries would lower friction to "download + run."

---

## Comparative positioning matrix

| Dimension | VisiData | Harlequin | Lazygit | **QuickSheet** |
|-----------|----------|-----------|---------|----------------|
| One-line pitch | "Explore tabular data" | "SQL IDE for terminal" | "Simple terminal UI for git" | "Desktop is a spreadsheet" |
| Hero visual | Animated GIF | SVG screenshot | Animated GIF | Static PNG ⚠️ |
| Install friction | `pip install` | `uv tool install` | `brew install` | Clone + dotnet build ⚠️ |
| Time-to-value | 5 sec | 10 sec | 2 sec | 30 sec ⚠️ |
| Plugin/extension story | Format loaders | DB adapters (dozens) | None | 69+ extensions ✓ |
| Social proof | Patreon + Discord | HN #2 + sponsor | Star popup + 78k★ | 0★ ⚠️ |
| Platform | Linux/macOS/WSL | All | All | Windows + Linux ✓ |
| License | GPLv3 | MIT | MIT | MIT ✓ |

---

## Implications for QuickSheet — Top 3 actions to queue

### 1. Create pre-built binary releases (HIGH PRIORITY)
Every successful tool above has a one-command install. QuickSheet's "clone + dotnet build"
is a real barrier. Use `dotnet publish -r linux-x64 --self-contained` and
`dotnet publish -r win-x64 --self-contained` to produce single-file binaries, then
attach them to GitHub Releases via `gh release create`. This transforms install from
"install .NET SDK + clone + build" to "download + chmod +x + run."

**Expected impact:** Unlocks awesome-list acceptance (maintainers test tools), enables HN launch,
lowers GitHub visitor → user conversion friction.

### 2. Animated demo GIF (needs human capture)
All three comparables lead with a GIF or high-quality visual showing the tool in action.
QuickSheet has a static PNG. An animated GIF showing:
- Typing a note on the desktop
- Launching an app via `r: code .`
- Live sparkline or extension updating
- Cell navigation and search

Would dramatically improve first-impression conversion. *This requires a human with a running
desktop to capture.* Queue as a request to the maintainer, not an autonomous action.

### 3. First-run star nudge (Bucket E, small)
Lazygit's author credits the "please star this repo" popup as the single biggest growth
contributor. QuickSheet could add a one-time message on first launch (when no autosave file
exists): "If QuickSheet is useful, please ⭐ github.com/cemheren/QuickSheet — it helps others
find it." Show once, never again. This is ~15 lines of code, zero dependencies, high ROI.

**Expected impact:** Converts a fraction of every new user into a stargazer. Compounds with
every other growth action.

---

## Secondary observations

- **Docs site:** Both VisiData and Harlequin have dedicated docs sites. QuickSheet's `docs/` folder is good but a GitHub Pages deployment would add legitimacy.
- **Community signal:** Discord/Matrix badge or Discussions tab signals "alive project." QuickSheet has none currently.
- **Version number in title:** VisiData shows "v3.3" right in the README header. This signals maturity. QuickSheet is at v0.36 — consider renaming to v1.0 when binary releases ship.
- **HN launch timing:** Do NOT launch on HN until binary releases + GIF exist. One shot. Make it count.
