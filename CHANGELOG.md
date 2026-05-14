# Changelog

All notable changes to QuickSheet. Format roughly follows [Keep a Changelog](https://keepachangelog.com/) and [SemVer](https://semver.org/).

## Unreleased

### Added
- `--version` / `-v` flag.
- `s: A1::A10` range form for sparkline cells (issue #1).
- `docs/recipes.md` — eight ready-to-paste wallpaper dashboards (ops, portfolio, stock watchlist, command center, writer, pomodoro+tasks, freelancer, academic, AI scratchpad).
- `docs/extensions.md` — live directory of installable extensions.
- `docs/tour.md` — 60-second feature tour.
- `CONTRIBUTING.md`, `SECURITY.md` — repo-health docs.

### Fixed
- Column auto-width over-counted raw cell length for `s:` sparkline cells (issue #2).

### Extensions
New extensions published as separate repos, installable via `ext: github:cemheren/<repo>`:
- `quicksheet-tls-ext` — TLS cert expiry + issuer.
- `quicksheet-price-ext` — crypto price quotes (CoinGecko).
- `quicksheet-define-ext` — inline dictionary (dictionaryapi.dev).
- `quicksheet-mortgage-ext` — fixed-rate amortization calculator.
- `quicksheet-mxck-ext` — MX record lookup (DNS-over-HTTPS).
- `quicksheet-ping-ext` — HTTP probe (status + latency).
- `quicksheet-cite-ext` — DOI → citation (Crossref).
- `quicksheet-thes-ext` — thesaurus (Datamuse).
- `quicksheet-stock-ext` — stock quotes (Stooq).
- `quicksheet-1099-ext` — US self-employment tax estimator.

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
