# Changelog

All notable changes to QuickSheet. Format roughly follows [Keep a Changelog](https://keepachangelog.com/) and [SemVer](https://semver.org/).

## Unreleased

(empty — bump here before the next tag)

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
