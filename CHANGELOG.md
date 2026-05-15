# Changelog

All notable changes to QuickSheet. Format roughly follows [Keep a Changelog](https://keepachangelog.com/) and [SemVer](https://semver.org/).

## Unreleased

(empty — bump here before the next tag)

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
