# Changelog

All notable changes to QuickSheet. Format roughly follows [Keep a Changelog](https://keepachangelog.com/) and [SemVer](https://semver.org/).

## 0.28.0 — 2026-05-22

### Added
- **quicksheet-whois extension** (`whois:`) — WHOIS domain/IP lookup with 25+ TLD server mappings, referral following, expiry urgency indicators (🟢🟡🔴), 1-hour cache. `whois: example.com`.

### Fixed
- Extension output no longer persisted to CSV (issue #156). Extension cell values are written to an ephemeral overlay dictionary; `SaveToCsv` reads from `_data` directly so extension text never appears in the saved file.
- Extension prefix cells with a trailing colon in the manifest (e.g. `roll:`, `lc:`, `env:`, `pihole:`) now register correctly. `TrimEnd(':')` normalises the prefix on registration so `roll: 2d6` activates the dice extension as expected. Closes #154.

## 0.36.0 — 2026-05-24

### Added
- **Ctrl+K row duplication** — press Ctrl+K in desktop mode to insert an exact copy of the current row below it; cursor moves to the duplicate. Fully undoable with Ctrl+Z.
- **`b:` bold cell prefix** — render any cell in bold. Type `b: text` to display text in bold weight (FontStyle.Bold on Windows, Xft bold on Linux). Bright white foreground for high contrast.
- **Auto-fit column widths** — desktop mode now sizes each column to its longest content (5–30 chars) with proportional scale-down when total exceeds screen width.
- **quicksheet-hackage extension** (`hackage:`) — Haskell package lookup from Hackage. Version, synopsis, author, category, license, homepage. `hackage: search <query>` for top results.
- **quicksheet-rubygems extension** (`gem:`) — Ruby gem lookup via rubygems.org. Version, download counts, authors, license, description. `gem: search <query>` for results.
- **quicksheet-crates extension** (`crates:`) — Rust crate lookup via crates.io. Version, download stats, description, links. `crates: search <query>` for results.
- **quicksheet-mtr extension** (`mtr:`) — network route tracer with per-hop RTT stats and packet loss. `mtr: google.com` (traceroute), `mtr: ping 8.8.8.8` (ping stats).
- **quicksheet-dict extension** (`dict:`) — English dictionary definitions, synonyms, and antonyms via free dictionaryapi.dev. `dict: syn <word>` / `dict: ant <word>` for targeted lookups.
- **quicksheet-stocks extension** (`stocks:`) — live stock and crypto ticker via Yahoo Finance. Single-ticker detail view or multi-ticker comparison table.

## Unreleased

(empty — bump here before the next tag)

## 0.24.0 — 2026-05-18

### Added
- **`--export-json`** CLI flag — exports CSV as a JSON array of objects. First row becomes object keys; numeric cells become JSON numbers. Supports stdout (`-`) for piping into `jq` and other tools.
- **quicksheet-co2 extension** (`co2:`) — live atmospheric CO₂ from NOAA Mauna Loa Observatory. Shows current ppm, pre-industrial baseline comparison, year-over-year change, 7-day sparkline. `co2: trend` (30-day view), `co2: stats` (annual averages). Free NOAA CSV, no API key.
- **quicksheet-gh-trends extension** (`ghtrend:`) — GitHub trending repos by language. Name, stars, forks, description grid. `ghtrend: python`, `ghtrend: rust`, `ghtrend: all`, etc.
- **`examples/ai-workflow.csv`** — pre-built Claude/Copilot/Aider launcher panel template.

### Docs
- `docs/export-formats.md` updated with `--export-json` section including `jq` piping examples and updated format comparison table.
- `docs/tour.md` updated with `config:` cell prefix documentation.

## 0.23.0 — 2026-05-18

### Added
- **`config:` cell prefix** — persist configuration (e.g. active theme) inside the CSV itself. Type `config: theme=Nord` in any cell; theme restores on next load. Ctrl+T writes back the new theme. No sidecar files — CSV is the only persistence. Closes #129.
- **quicksheet-npm extension** — npm package info (version, weekly downloads, license, author, last publish date). Single-package detail view or multi-package comparison table. No API key required.
- **quicksheet-pypi extension** — PyPI package info (version, license, author, Python requirement, release date, homepage). Single-package detail or comparison table. Free pypi.org JSON API.

## 0.22.0 — 2026-05-17

### Added
- **`--export-html`** CLI flag — exports CSV as a self-contained dark-themed HTML table with auto-linked URLs, right-aligned numbers, and a QuickSheet backlink. Supports stdout piping for scripts.
- **quicksheet-iss extension** — live International Space Station tracker. Shows lat/lon, altitude, speed, and ground region, plus all people currently in space grouped by spacecraft.
- **Export formats guide** — `docs/export-formats.md` covering CSV, Markdown (`--export-md`), and HTML (`--export-html`) with usage examples and piping patterns.

## 0.21.0 — 2026-05-17

### Added
- **quicksheet-leetcode extension** — LeetCode daily challenge, problem lookup by number/slug, user solve stats (Easy/Medium/Hard counts, global rank).
- **quicksheet-ghstreak extension** — GitHub contribution streak tracker. Current streak 🔥, longest streak, total contributions, 14-day unicode sparkline. No auth needed.
- **Startup launcher scripts** — `scripts/quicksheet-startup.ps1` (Windows) and [docs/install-startup.md](docs/install-startup.md) walkthrough for Task Scheduler, Startup folder, and XDG autostart. Opt-in `-Update` flag for auto-pull-and-build.
- **Copilot use-case table** in README + `examples/copilot-dashboard.csv` template (8 AI workflow scenarios).


## 0.20.0 — 2026-05-17

### Added
- **quicksheet-dice extension** — dice roller with standard notation (`roll: 2d6+3`, `roll: d20`, `roll: 4d6kh3` keep-highest, Fudge/FATE `dF`, percentile `d%`). Critical hit/fumble detection on d20.

### Fixed
- `Program.cs` version constant updated to match v0.19.0 tag.

## 0.19.0 — 2026-05-17

### Added
- **5 new extensions**: quicksheet-pihole (Pi-hole DNS stats), quicksheet-health (HTTP endpoint health checker), quicksheet-envck (env var inspector with secret masking), quicksheet-gha (GitHub Actions workflow status), quicksheet-arxiv (arXiv paper lookup).
- **Audience landing pages**: [for-traders](docs/for-traders.md), [for-sre](docs/for-sre.md), [for-students](docs/for-students.md) — with starter CSV dashboards.
- **csvkit/Miller/qsv comparison page** ([docs/csvkit-comparison.md](docs/csvkit-comparison.md)) — SEO-targeted.
- **README hero tightened** — highlights CSV format, runnable cells, cross-platform.

### Fixed
- `--help` output now lists `c:color:` cell prefix.

## 0.18.0 — 2026-05-16

### Added
- **quicksheet-urlenc extension** — URL encode/decode with auto-detect, UTF-8 aware.
- **quicksheet-curl extension** — cURL-style HTTP client (GET/POST/PUT/DELETE) with JSON pretty-print.
- **c?: conditional color prefix** — color only applied when cell value is non-empty.
- **for-homelab landing page** ([docs/for-homelab.md](docs/for-homelab.md)) with starter CSV.

## 0.17.0 — 2026-05-16

### Added
- **quicksheet-guid extension** — generates GUIDs/UUIDs with format options (standard/no-dash/braced/uppercase), batch up to 20. Closes #91.
- **quicksheet-regex extension** — regex pattern explainer. Tokenizes anchors, character classes, quantifiers, groups. Closes #90.
- **quicksheet-b64 extension** — base64 encode/decode with auto-detect.
- **5 additional themes** (Dracula, Gruvbox, Tokyo Night, Catppuccin, One Dark) added to theme cycling.
- **Extension protocol spec** ([docs/extension-protocol.md](docs/extension-protocol.md)) — full lifecycle, message schemas, coordinate system, common mistakes. Closes #89.

### Fixed
- Themes now render correctly in Linux desktop mode (#94) — `ConsoleColorToRgb()` drives all color decisions.



### Added
- **quicksheet-news extension** — RSS/Atom feed reader with 15+ built-in aliases (HN, Reddit, Lobsters, dev.to, BBC, TechCrunch). Supports any feed URL.
- **CONTRIBUTING.md refresh** — architecture diagram, extension authoring guide with protocol details, scannable ground rules table.
- **gh-pages update** — 9 new extension cards (cntdn, worldtm, margin, mileage, depr, jwtdec, cronck, k8s, news), version bump to 0.15.0, comparison table updated to "40+ extensions".

## 0.15.0 — 2026-05-15

### Added
- **Sparkline rendering in Windows desktop mode** (#9) — `s: 1,2,3,4,5` now renders as unicode block-bar glyphs in wallpaper mode, matching the Linux implementation.
- **README badges** — GitHub release version (dynamic), zero dependencies, 30+ extensions count.

### Fixed (extensions)
- Fixed `quicksheet-cronck` activate handler to read from `params` array instead of `cells`.

## 0.14.0 — 2026-05-15

### Fixed
- **Theme cycling in Windows desktop mode** (#77) — `ConsoleColorToRgb()` full 16-color mapper replaces hardcoded colors. All theme presets (Dark/Light/Nord/Solarized/Matrix) now render correctly in desktop wallpaper mode.
- **5 missing desktop shortcuts restored** (#66) — Ctrl+Z (undo), Ctrl+Y (redo), Ctrl+T (theme), Ctrl+G (goto cell), Ctrl+H (help overlay) now work in Windows desktop mode.

### Fixed (extensions)
- Fixed protocol bugs in 4 extensions: `quicksheet-cronck`, `quicksheet-jwtdec`, `quicksheet-rate`, `quicksheet-gitlog`.

## 0.13.0 — 2026-05-15

### Added
- **quicksheet-cronck extension** — cron expression parser. Ranges, lists, steps, named days/months.
- **quicksheet-gitlog extension** — recent git commits in cells for project dashboards.

### Fixed
- Fixed duplicate `colorParsed` variable (#75) — build error on some configurations.

## 0.12.0 — 2026-05-15

### Fixed
- **c:color: prefix in Windows desktop mode** (#65) — `CellPrefix.ParseColor()` now rendered in DesktopForm with ConsoleColor→RGB mapping matching the Linux implementation.

### Added
- **Ctrl+Z/Y/T/G shortcuts in Linux desktop mode** (#66, partial) — Undo, redo, new tab, and goto-cell shortcuts now work in X11 desktop host.
- `quicksheet-jwtdec` extension — decode JWTs entirely offline. Privacy-first: tokens never leave your machine.

### Fixed (extensions)
- Fixed `quicksheet-depr-ext` protocol types (`register`/`invoke` → `init`/`activate`, cell format `r`/`c`/`v`).
- Fixed `quicksheet-mileage-ext` manifest (missing prefix field).
- Fixed `quicksheet-worldtm` manifest key (`entrypoint` → `entry`).

## 0.11.0 — 2026-05-15

### Added
- **Cell color prefix `c:COLOR:`** (#61) — 9 named colors (red, green, blue, yellow, cyan, magenta, white, gray/grey) as cell backgrounds. Works in console + Linux desktop.
- **FAQ page** on website — common questions about desktop mode, extensions, CSV format.
- **Color prefix documentation** — keyboard-shortcuts.md color table + tour.md cheatsheet entry.

### Extensions (new repos)
- `quicksheet-depr-ext` — Straight-line & MACRS depreciation schedules.

## 0.10.0 — 2026-05-15

### Added
- **Ctrl+R Find & Replace in desktop mode** (#50) — full find-and-replace state machine in both Windows (WinForms) and Linux (X11) desktop hosts. Phase-based: Find → Replace → Confirm.
- **Sparkline rendering on Linux desktop** (#9, partial) — `s:` prefix cells now render unicode block-bar charts in Linux desktop mode.

### Extensions (new repos)
- `quicksheet-worldtm` — Multi-timezone world clock with 40+ aliases, business-hours indicators (🟢🟡🔴).
- `quicksheet-k8s` — Live Kubernetes pod status from kubeconfig. Color-coded status icons.
- `quicksheet-mileage-ext` — IRS standard-mileage deduction calculator (business/medical/charity rates).
- `quicksheet-margin-ext` — Break-even & contribution margin calculator with health indicators.

### Fixed (extensions)
- Fixed `quicksheet-worldtm` manifest key (`entrypoint` → `entry`).
- Fixed `quicksheet-mileage-ext` manifest (missing prefix field, enriched metadata).
- Fixed `quicksheet-portck` manifest key (`entrypoint` → `entry`).
- Fixed `quicksheet-hntop` manifest key (`entrypoint` → `entry`).
- Fixed `quicksheet-ghpr` search fields + params mismatch.
- Fixed `quicksheet-gitst` params field (`arguments` → `params` array).

### Changed
- **README curated** — trimmed from 26 verbose extension examples to 3 hero showcases + compact install table.
- **Repo topics refreshed** — added 11 new GitHub topics for discoverability.

## 0.8.0 — 2026-05-15

### Added
- **Find & Replace (Ctrl+R)** — search and replace text across cells with per-match confirmation dialog. Works in both console and desktop mode.
- **Keyboard shortcuts documentation** — complete reference at `docs/keyboard-shortcuts.md` covering all keys, cell prefixes, built-in math, and desktop mode notes.
- **Shortcuts page on website** — SEO-targeted page at [/shortcuts/](https://cemheren.github.io/QuickSheet/shortcuts/).

### Fixed
- **Extension protocol: int version accepted** (#43) — extensions sending `version: 1` (int) instead of `"1.0.0"` (string) in register messages now work via `FlexVersionConverter`. Fixes quicksheet-gitst, quicksheet-ghpr, quicksheet-docker.

### Extensions (separate repos)
- `quicksheet-cntdn` — countdown to dates (deadlines, holidays, launches).

### Fixed (extensions)
- Fixed `quicksheet-cntdn` manifest key (`entryPoint` → `entry`).
- Fixed `quicksheet-gitst` params field (`arguments` → `params` array).
- Fixed `quicksheet-docker` Windows named pipe support, params field, error visibility.

## 0.7.0 — 2026-05-15

### Fixed
- **Define-ext crash** (#17) — extensions sending array-of-arrays cell format no longer crash the host. Added `CellWriteArrayConverter` handling object, grid, and flat formats with try-catch safety.
- **Ctrl+B sort in desktop mode** (#35) — column sorting was only wired in console mode. Now works in Windows (WinForms) and Linux (X11) desktop wallpaper mode.

### Extensions (separate repos)
- `quicksheet-ghpr` — GitHub PR review dashboard with status icons.
- `quicksheet-portck` — TCP port/service health checker with color-coded UP/DOWN.
- `quicksheet-docker` — Docker container status dashboard via Engine API.
- `quicksheet-gitst` — Git repo status (branch, changes, stashes, last commit).
- `quicksheet-rate` — Freelance hourly rate calculator with tax/benefit modeling.

## 0.6.0 — 2026-05-15

### Fixed
- **Extension install corruption** (#26) — install failures no longer append status suffixes to cell text. Tracked in memory only.

### Added
- **Ctrl+B column sorting** — sort any column ascending/descending with numeric awareness. Empty cells sort last. Fully undoable.

### Extensions (separate repos)
- `quicksheet-hntop` — Hacker News top stories on your wallpaper (Firebase API, 5-min cache).
- `quicksheet-apistatus` — Service health monitor for GitHub, Cloudflare, npm, Discord + 14 more.
- `quicksheet-budget` — Budget envelope visualizer with progress bars.
- `quicksheet-qtr` — IRS quarterly tax deadline countdown.
- `quicksheet-fx` — Live currency conversion (170+ currencies, ECB rates).

### Fixed (extensions)
- Fixed cell format bug in 10 extensions (stock, price, ping, mortgage, tls, mxck, grav, 1099, thes, cite) — all now emit correct `{r,c,v}` format.
- Fixed `quicksheet-price-ext` CoinGecko 403 — added CoinCap API fallback.

## 0.5.0 — 2026-05-14

### Added
- **Go-to-cell navigation** — press Ctrl+G, type a cell reference (e.g. `C5`), and jump directly to it. Invalid refs show inline error feedback.

### Extensions (separate repos)
- `quicksheet-todo` — task management with priorities, due dates, completion tracking, and progress bar.

### Docs
- Renamed generic screenshot filenames (`image.png`, `image-1.png`, etc.) to meaningful SEO-friendly names for better discoverability.

## 0.4.0 — 2026-05-14

### Added
- **Enhanced status bar** — shows filename, modified indicator (●), and non-empty cell count alongside Σ/Π.
- **Extension deactivation** — extensions receive a deactivate message when their prefix cell is deleted, enabling clean shutdown.
- **Issue templates** — structured bug report, feature request, and extension idea templates for better community contributions.
- **GitHub Discussions** enabled for community Q&A and feature brainstorming.

### Docs
- Complete keyboard shortcuts table in README (Ctrl+T themes, Ctrl+Z/Y undo/redo, Ctrl+H help).
- Fixed all image alt text descriptions to be meaningful and accessible.

## 0.3.0 — 2026-05-14

### Added
- **Undo / redo** with Ctrl+Z / Ctrl+Y. Tracks cell edits, row insertions, row deletions, and bulk row clears. Multi-step grouping for compound operations.
- `--list-extensions` flag prints installed extensions from `~/.quicksheet/extensions/` with prefix, repo dir, and version from each manifest.
- `--export-md -` writes the Markdown table to stdout instead of a file, for piping.
- Extension lifecycle debug logging — easier to diagnose extension load/exit issues.

## 0.2.0 — 2026-05-14

### Added
- Theme presets — cycle TUI color palette with Ctrl+T.
- `--version` / `-v` flag.
- `s: A1::A10` range form for sparkline cells (closes #1).
- `docs/recipes.md` — eight ready-to-paste wallpaper dashboards (ops, portfolio, stock watchlist, command center, writer, pomodoro+tasks, freelancer, academic, AI scratchpad).
- `docs/extensions.md` — live directory of installable extensions.
- `docs/tour.md` — 60-second feature tour.
- `CONTRIBUTING.md`, `SECURITY.md`, `CHANGELOG.md` — repo-health docs.

### Fixed
- Column auto-width over-counted raw cell length for `s:` sparkline cells (closes #2).

### Extensions (separate repos, installable via `ext: github:cemheren/<repo>`)
- `quicksheet-tls-ext`
- `quicksheet-price-ext`
- `quicksheet-define-ext`
- `quicksheet-mortgage-ext`
- `quicksheet-mxck-ext`
- `quicksheet-ping-ext`
- `quicksheet-cite-ext`
- `quicksheet-thes-ext`
- `quicksheet-stock-ext`
- `quicksheet-1099-ext`
- `quicksheet-grav-ext`

## 0.1.0 — 2026-05-14

First tagged release. The project had been usable for some time; this snapshot tags a coherent point with documented features and a small extension ecosystem.

### Core
- Interactive TUI spreadsheet with CSV persistence and 5-second autosave.
- Desktop wallpaper mode (`--desktop`) — Win32 WorkerW on Windows, `_NET_WM_WINDOW_TYPE_DESKTOP` on Linux/X11.
- Cell prefixes: `r:` runnable, `i:` inline subprocess, `s:` sparkline, `L:` interval loop, `ext: github:user/repo` extension install. URLs auto-detected as hyperlinks.
- `{A1::C10}` range references inside any text.
- Σ (column sum) and Π (row product) in the status bar.
- Headless `--export-md` for CSV → GitHub-flavored Markdown.
- `--help` with full prefix cheatsheet.

### Constraints kept
- Zero NuGet dependencies. All native interop hand-written P/Invoke.
- CSV as the only persistence format.

### Known limits
- No Wayland desktop-embedding (issue #3).
- No prebuilt binaries in this release.

Tag: https://github.com/cemheren/QuickSheet/releases/tag/v0.1.0
