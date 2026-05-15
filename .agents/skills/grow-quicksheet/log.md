# grow-quicksheet log (Copilot CLI)

Persistent memory across runs. Append-only (except the Queued section at bottom).

## 2026-05-13

- Stars: 0 (baseline)
- Action: Skill created. Studying competitor log to plan first high-impact action.
- Bucket: meta
- Outcome: Skill scaffolded at `.agents/skills/grow-quicksheet/`.
- Competitor last did: Set GitHub repo topics (Bucket B), README audit fix (Bucket A).
- Follow-up: First real run should ship a small feature (Bucket E) or create compelling content that the competitor hasn't covered yet.

## 2026-05-13 22:35

- Stars: 0 (Δ +0 since baseline)
- Action: Drafted Show HN post with title options, first comment, timing notes, and pre-flight checklist.
- Bucket: C
- Outcome: draft saved at `.agents/skills/grow-quicksheet/drafts/showhn.md`
- Competitor last did: Bucket B (repo topics) and Bucket A (README audit).
- Follow-up: Next run should ship a feature (Bucket E) — theme presets or status bar. Then use the feature as screenshot material for the HN post.

## 2026-05-13 23:35

- Stars: 0 (Δ +0)
- Action: Created quicksheet-pomodoro extension as separate repo + added to main README.
- Bucket: F
- Outcome: Repo live at https://github.com/cemheren/quicksheet-pomodoro (commit 25298ee). README updated on main (commit a4cafc0).
- Competitor last did: Bucket B (topics), Bucket A (README).
- Follow-up: Next run — create quicksheet-sysmon (CPU/RAM monitor) or build the GitHub Pages site.

## 2026-05-14 07:00

- Stars: 0 (Δ +0)
- Action: Bucket C — created GitHub Pages landing page. Plain HTML/CSS, dark theme, SEO-optimized. Features: hero with install command, feature grid (6 cards), extension directory (12 extensions), comparison table (vs sc-im, VisiData), 4-step getting started, CTAs. SEO: Open Graph + Twitter Card meta, Schema.org SoftwareApplication structured data, sitemap.xml, robots.txt, 404 page. Keywords targeted: terminal spreadsheet, desktop wallpaper spreadsheet, tui spreadsheet linux, dotnet spreadsheet cli. Set repo homepage URL.
- Bucket: C
- Outcome: Pushed to `gh-pages` branch (commit 45ca5a6). Pages building at https://cemheren.github.io/QuickSheet/. Repo homepage set.
- Competitor last did: 34+ runs — extensive extensions (F), features (E), docs (A), releases (B), social drafts (C). No website yet.
- Follow-up: Verify site is live after build. Add subpages for long-tail keywords (features, extensions). Add OG image. Consider Bucket E next for variety.

## 2026-05-14 07:20

- Stars: 0 (Δ +0)
- Action: Added theme presets for TUI mode — 5 themes (Dark, Light, Nord, Solarized, Matrix) with Ctrl+T cycling.
- Bucket: E
- Outcome: Shipped commit 4b96bb6 on main. Build green. Additive-only (new Theme.cs + minimal SpreadsheetApp.cs wiring).
- Competitor last did: 11 local runs — extensions (F), features (sparkline, --help, --version, --export-md), docs (A), release (B).
- Follow-up: Update README keyboard shortcuts table to document Ctrl+T. Screenshot the Matrix theme for README eye candy.

## 2026-05-14 08:00

- Stars: 0 (Δ +0)
- Action: Bucket C — added /features/ and /extensions/ subpages to GitHub Pages site. Features page targets long-tail keywords (terminal spreadsheet, desktop wallpaper spreadsheet, tui spreadsheet linux). Extensions page is a full directory with install commands and protocol docs. Added nav bar across all pages. Updated sitemap.xml.
- Bucket: C
- Outcome: Pushed to gh-pages branch (commit a7392d6). Pages building.
- Competitor last did: 11 runs — extensions (F), features (E sparkline/export-md), docs/tour.md (A), issues (B), drafts (C).
- Follow-up: Next run — Bucket E (status bar) or Bucket F (quicksheet-sysmon) for variety. OG image still needed.

## 2026-05-14 09:00

- Stars: 0 (Δ +0)
- Action: Created quicksheet-sysmon extension — live CPU/RAM/disk/uptime monitor with visual progress bars and color-coded indicators.
- Bucket: F
- Outcome: Repo live at https://github.com/cemheren/quicksheet-sysmon. README + docs/extensions.md updated on main (commit 43edee4).
- Competitor last did: 17 runs — extensive extensions (F), features (E), docs (A), issues (B), drafts (C). Column auto-width fix for sparklines.
- Follow-up: Next run — Bucket E (small feature) or Bucket A (screenshot rename polish) for variety. OG image still needed for website.

## 2026-05-14 10:00

- Stars: 0 (Δ +0)
- Action: Bucket E — added undo/redo stack. New UndoManager.cs with action-grouped history (200 steps). All GridManager mutations (SetCellValue, ClearRow, DeleteRow, ShiftRowsDown/Up) record undo history. Ctrl+Z / Ctrl+Y wired in SpreadsheetApp. Help overlay updated.
- Bucket: E
- Outcome: Shipped commit a22b973 on main. Build green. Pure additive — no existing behavior changed.
- Competitor last did: 17+ runs — extensions (F), features (E: --export-md stdout, --list-extensions, sparkline column-width fix), docs (A: extensions.md, tour.md, --help), releases (B: v0.1.0, v0.2.0), research (R: Show HN patterns, niche communities).
- Follow-up: Update CHANGELOG.md Unreleased section. Update docs/tour.md with undo/redo mention. Cut v0.3.0 release next.

## 2026-05-14 11:00

- Stars: 0 (Δ +0)
- Action: Bucket B — created 3 structured issue templates (bug, feature request, extension idea) + config with links. Enabled GitHub Discussions. Created welcome discussion post.
- Bucket: B
- Outcome: Shipped commit 6b0afa1 on main. Discussion live at https://github.com/cemheren/QuickSheet/discussions/4.
- Competitor last did: no-op run — noted "supply saturated, publication is bottleneck."
- Follow-up: OG image for social sharing. Website extensions page update with sysmon.

## 2026-05-14 12:00

- Stars: 0 (Δ +0)
- Action: Bucket A — completed README keyboard shortcuts table (added Ctrl+T themes, Ctrl+Z/Y undo/redo, Ctrl+H help as new sections). Fixed all generic "alt text" image descriptions with meaningful captions. Added Ctrl+T to in-app help overlay.
- Bucket: A
- Outcome: Shipped commit b369cfe on main. Build green.
- Competitor last did: 4 consecutive no-op runs — noted "supply saturated, publication is bottleneck."
- Follow-up: OG image for website social sharing. Update gh-pages extensions page with sysmon.

## 2026-05-14 13:00

- Stars: 0 (Δ +0)
- Action: Bucket C — updated all 3 gh-pages site pages with features/extensions shipped since last site update. Homepage: added undo/redo + theme preset feature cards, sysmon extension card, 2 new comparison table rows (undo/redo, themes), bumped structured data to v0.3.0. Features page: added undo/redo section, added Ctrl+Z/Y/H to keyboard shortcuts table. Extensions page: added sysmon and mortgage calculator cards.
- Bucket: C
- Outcome: Pushed to gh-pages (commit 38f3c0d). Pages rebuilding.
- Competitor last did: 4+ consecutive no-op runs — "supply saturated, publication is bottleneck."
- Follow-up: OG image for social sharing. Bucket E (column auto-resize) for variety next.

## 2026-05-14 14:00

- Stars: 0 (Δ +0)
- Action: Bucket E — enhanced TUI status bar with filename, modified indicator (●), and non-empty cell count. Added dirty-state tracking across all edit operations.
- Bucket: E
- Outcome: Shipped commit 2705ee8 on main. Build green. Pure additive — one file changed (SpreadsheetApp.cs), 26 insertions.
- Competitor last did: 4+ consecutive no-op runs — "supply saturated, publication is bottleneck."
- Follow-up: Cut v0.4.0 release bundling recent improvements. Update gh-pages site with status bar feature.

## 2026-05-14 15:00

- Stars: 0 (Δ +0)
- Action: Bucket D — cut v0.4.0 release bundling status bar enhancement, extension deactivation, community issue templates, and docs polish.
- Bucket: D
- Outcome: Release live at https://github.com/cemheren/QuickSheet/releases/tag/v0.4.0 (tag 88fab87). Appears in follower feeds.
- Competitor last did: 4+ consecutive no-op runs — "supply saturated, publication is bottleneck."
- Follow-up: Update gh-pages site structured data to v0.4.0. OG image for social sharing.

## 2026-05-14 16:00

- Stars: 0 (Δ +0)
- Action: Bucket C — created OG image (1200×630, dark theme, mini spreadsheet grid with sample data, feature badges) and deployed to gh-pages. Added og:image + twitter:image meta tags to all 3 site pages. Updated structured data softwareVersion to 0.4.0.
- Bucket: C
- Outcome: Pushed to gh-pages (commit 98a6a4d). OG image live at https://cemheren.github.io/QuickSheet/og-image.png. Social sharing previews now render on Twitter, Discord, Slack, etc.
- Competitor last did: 4+ consecutive no-op runs — "supply saturated, publication is bottleneck."
- Follow-up: Column auto-resize (Bucket E) or screenshot rename (Bucket A) for variety.

## 2026-05-14 17:00

- Stars: 0 (Δ +0)
- Action: Created quicksheet-todo extension — task management with priorities (!low/normal/high/critical), due dates (@YYYY-MM-DD), completion tracking, persistent CSV storage, progress stats with visual bar.
- Bucket: F
- Outcome: Repo live at https://github.com/cemheren/quicksheet-todo. README + docs/extensions.md updated on main (commit 703c0dd).
- Competitor last did: 4+ consecutive no-op runs — "supply saturated, publication is bottleneck."
- Follow-up: Column auto-resize (Bucket E) or quicksheet-cal (Bucket F) for variety next.

## 2026-05-14 18:00

- Stars: 0 (Δ +0)
- Action: Bucket A — renamed 4 generic screenshot files (image.png, image-1.png, image-2.png, image-4.png) to SEO-friendly descriptive names (desktop-wallpaper-commands, hyperlink-dashboard, data-tracking-autosum, desktop-files-grid). Updated all 5 README image references.
- Bucket: A
- Outcome: Shipped commit c8a9ca0 on main. No code changes — safe rename + README update.
- Competitor last did: 4+ consecutive no-op runs — "supply saturated, publication is bottleneck."
- Follow-up: Column auto-resize keybinding (Bucket E) or quicksheet-cal (Bucket F) for variety.

## 2026-05-14 19:00

- Stars: 0 (Δ +0)
- Action: Bucket E — added Ctrl+G go-to-cell navigation. Prompts for cell reference (e.g. A1, C5), jumps cursor. Reuses existing CellPrefix.ParseCellRef. Error feedback for invalid refs. Help overlay updated.
- Bucket: E
- Outcome: Shipped commit e5ac8ca on main. Build green. Pure additive — one file changed (SpreadsheetApp.cs), 65 insertions.
- Competitor last did: 4+ consecutive no-op runs — "supply saturated, publication is bottleneck."
- Follow-up: Update gh-pages site with go-to-cell feature. Cut v0.5.0 release next.

## 2026-05-14 20:00

- Stars: 0 (Δ +0)
- Action: Bucket D — cut v0.5.0 release bundling go-to-cell navigation, quicksheet-todo extension, and SEO screenshot renames.
- Bucket: D
- Outcome: Release live at https://github.com/cemheren/QuickSheet/releases/tag/v0.5.0 (tag 9687dd6). Appears in follower feeds.
- Competitor last did: 4+ consecutive no-op runs — "supply saturated, publication is bottleneck."
- Follow-up: Update gh-pages site with go-to-cell feature + v0.5.0 structured data. quicksheet-cal extension for variety.

## 2026-05-14 21:00

- Stars: 0 (Δ +0)
- Action: Bucket C — updated all 3 gh-pages site pages with v0.5.0 features. Homepage: added go-to-cell and status bar feature cards, quicksheet-todo extension card, 2 new comparison table rows. Features page: added go-to-cell section, status bar section, Ctrl+G to keyboard shortcuts. Extensions page: added quicksheet-todo card. Bumped structured data to v0.5.0.
- Bucket: C
- Outcome: Pushed to gh-pages (commit f9295cb). Pages rebuilding.
- Competitor last did: 4+ consecutive no-op runs — "supply saturated, publication is bottleneck."
- Follow-up: Column auto-resize (Bucket E) or quicksheet-cal (Bucket F) for variety next.

## 2026-05-14 22:00

- Stars: 0 (Δ +0)
- Action: Created quicksheet-cal extension — reads .ics calendar files (RFC 5545), shows upcoming events grouped by date with time/summary/location/duration. Auto-scans common calendar dirs (Evolution, Thunderbird, KDE, Calcurse, ~/Calendars). Supports "cal: today", "cal: week", "cal: month", "cal: path/to/file.ics".
- Bucket: F
- Outcome: Repo live at https://github.com/cemheren/quicksheet-cal. README + docs/extensions.md updated on main (commit 71a2e27).
- Competitor last did: 4+ consecutive no-op runs — "supply saturated, publication is bottleneck."
- Follow-up: Update gh-pages with cal extension card. quicksheet-news (RSS) for variety next.

## 2026-05-14 23:09

- Stars: 0 (Δ +0)
- Action: Fixed issue #10 — added all 7 missing extensions to README (stock-ext, ping-ext, 1099-ext, grav-ext, thes-ext, cite-ext, mxck-ext).
- Bucket: A (issue fix / polish)
- Outcome: PR #11 opened (commit de6677a on grow/readme-all-extensions). Closes #10.
- Competitor last did: 4+ consecutive no-op runs — "supply saturated, publication is bottleneck."
- Follow-up: Update gh-pages extensions page with all 7 new extensions.

## 2026-05-15 00:00

- Stars: 0 (Δ +0)
- Action: Fixed issue #17 — define-ext crashes the process. Added CellWriteArrayConverter to handle both `{r,c,v}` object format and `string[][]` grid format in extension protocol. Added try-catch around message processing to prevent any malformed extension message from crashing the host.
- Bucket: E (bug fix)
- Outcome: PR #18 opened (commit 7886d46 on grow/fix-define-ext-crash). Closes #17, also fixes quicksheet-define-ext#4 and #1.
- Competitor last did: 4+ consecutive no-op runs — "supply saturated, publication is bottleneck."
- Follow-up: Update gh-pages extensions page. Vim keybindings or quicksheet-news for variety.

## 2026-05-15 00:27

- Stars: 0 (Δ +0)
- Action: Deep market research for issue #15 (accounting extensions). Studied hledger/ledger/beancount ecosystem (~15k stars combined), HN plain text accounting threads, IRS quarterly dates, Frankfurter currency API (free, no key). Rated 8 extension ideas on usefulness × virality × feasibility.
- Bucket: R (research)
- Outcome: Design doc saved at `.agents/skills/grow-quicksheet/drafts/designs/accounting-suite.md`. Top 3: budget envelope visualizer (5/5/5), quarterly tax countdown (5/4/5), currency conversion (4/4/5).
- Competitor last did: 4+ consecutive no-op runs.
- Follow-up: Build `quicksheet-budget` extension next run (Tier 1 #1 — highest virality, pure math, ~80 LOC). Then `qtr:` and `fx:` in subsequent runs.

## 2026-05-15 00:49

- Stars: 0 (Δ +0)
- Action: Created quicksheet-budget extension — budget envelope visualizer with visual progress bars, color-coded spending indicators (🟢🟡🟠🔴), remaining balance tracking. Uses correct {r,c,v} cell format.
- Bucket: F
- Outcome: Repo live at https://github.com/cemheren/quicksheet-budget (commit 4783662). PR #20 adds to README/docs.
- Competitor last did: 4+ consecutive no-op runs — "supply saturated, publication is bottleneck."
- Follow-up: Build quicksheet-qtr (quarterly tax countdown) next. Then quicksheet-fx (currency).

## 2026-05-15 00:57

- Stars: 0 (Δ +0)
- Action: Fixed issue #19 — ext: cells processed during typing, causing premature install failures. Added editingRow/editingCol params to ScanGrid() so extension system skips the cell being typed into. Fixed on both Windows and Linux.
- Bucket: E (bug fix)
- Outcome: PR #21 opened (commit e12ccab on grow/fix-ext-typing-race). Closes #19.
- Competitor last did: Created quicksheet-budget extension (PR #20).
- Follow-up: Build quicksheet-qtr (quarterly tax countdown) next run.

## 2026-05-15 01:00

- Stars: 0 (Δ +0)
- Action: Created quicksheet-qtr extension — quarterly IRS estimated tax deadline countdown with urgency indicators (🔴🟠🟡🟢), progress bar, and full-year view. Auto-detects next deadline, adjusts for weekends. Pairs with 1099-ext and budget.
- Bucket: F
- Outcome: Repo live at https://github.com/cemheren/quicksheet-qtr. PR #22 adds to README/docs.
- Competitor last did: Fixed issue #19 ext typing race (PR #21).
- Follow-up: Build quicksheet-fx (currency conversion) next. Then update gh-pages site.

## 2026-05-15 01:45

- Stars: 0 (Δ +0)
- Action: Built and shipped quicksheet-fx extension — live currency conversion via ECB/Frankfurter API. 200+ currencies, no API key, 1-hour rate cache, multi-target conversion. Tested with real API calls. Created repo, added to README + docs/extensions.md.
- Bucket: F
- Outcome: Repo live at https://github.com/cemheren/quicksheet-fx. PR #24 adds to README/docs.
- Competitor last did: Created quicksheet-qtr (PR #22).
- Follow-up: Update gh-pages site with recent extension cards. Then Tier 2 extensions.

## 2026-05-15 01:55

- Stars: 0 (Δ +0)
- Action: Updated gh-pages site with 4 new extension cards (budget, qtr, fx, cal) on both homepage and extensions directory. Updated sitemap dates.
- Bucket: C (website)
- Outcome: Pushed to gh-pages branch (commit c3660bc). Site live at cemheren.github.io/QuickSheet/.
- Competitor last did: Created quicksheet-fx (PR #24).
- Follow-up: docs/tour.md polish or Tier 2 extensions next.

## 2026-05-15 01:57

- Stars: 0 (Δ +0)
- Action: Bucket C — resolved merge conflicts on gh-pages with competitor's parallel update. Merged best descriptions. Fixed build artifacts that leaked into gh-pages. Added .gitignore.
- Bucket: C (merge fix)
- Outcome: Pushed to gh-pages (commit e7da80b). Used `ext: github:` install format.
- Competitor last did: Also updated gh-pages with budget/qtr/fx/cal cards (concurrent work).
- Follow-up: Market research for next extension vertical. Or Tier 2 accounting extensions.

## 2026-05-15 02:27

- Stars: 0 (Δ +0)
- Action: Bucket A — updated docs/tour.md with 3 missing extensions (budget, qtr, fx), new section 5½ for headless --export-md mode, and freelancer finance dashboard use case.
- Bucket: A
- Outcome: PR #25 opened (commit 43012e7 on grow/update-tour-docs). Docs-only, zero risk.
- Competitor last did: Still stalled — last real action was ~May 14 local run #12 (markdown export).
- Follow-up: Market research for devops/data-science extension vertical next run.

## 2026-05-15 02:49

- Stars: 0 (Δ +0)
- Action: Merged all 10 extension cell-format fix PRs (stock, price, ping, mortgage, tls, mxck, grav, 1099, thes, cite). Merged 4 main repo docs PRs (#20 budget, #22 qtr, #24 fx, #25 tour). Closed duplicate PR #23. Rebased #24 to resolve conflicts.
- Bucket: D (merge & maintenance)
- Outcome: 10 extension bugs fixed (all now emit correct {r,c,v} format). 4 docs PRs merged. 1 duplicate closed. 14 PRs resolved in one run.
- Competitor last did: Updated docs/tour.md (PR #25).
- Follow-up: Tier 2 extensions (rate, deduct, pl) or Bucket E features.

## 2026-05-15 02:57

- Stars: 0 (Δ +0)
- Action: Fixed issue #26 — ext: install failures were corrupting cell values by appending status suffixes. Moved failure tracking to in-memory _failedSources dict. Added backward-compat suffix stripping in ParseExtensionSource. Re-edit and F3 rebuild now allow retry.
- Bucket: E (bug fix)
- Outcome: PR #27 opened (commit a2320b1 on grow/fix-ext-install-corruption). Closes #26. Build green.
- Competitor last did: Merged 14 PRs in one run — fixed 10 extension {r,c,v} bugs, 4 doc PRs.
- Follow-up: Market research for new extension vertical (devops/data science).

## 2026-05-15 03:00

- Stars: 0 (Δ +0)
- Action: Bucket E — added Ctrl+B column sorting. Numeric-aware (numbers sort numerically, text lexicographically). Empty cells sort last. Toggle asc/desc on repeated press. Fully undoable via Ctrl+Z. Help overlay updated.
- Bucket: E
- Outcome: PR #28 opened (commit a889552 on grow/sort-by-column). Build green.
- Competitor last did: 4+ consecutive no-op runs — "supply saturated, publication is bottleneck."
- Follow-up: Update gh-pages site with sort feature. Update README keyboard shortcuts table. quicksheet-rate extension for variety.

## 2026-05-15 03:27

- Stars: 0 (Δ +0)
- Action: Deep market research for devops/SRE/developer-productivity extensions. Studied trending TUI tools (k9s 33k★, lazygit 54k★, gh-dash 8k★, wtfutil 16k★), identified 6 viral patterns, designed 15+ extension concepts across 3 tiers. Key insight: "replace a browser tab" is the strongest viral pattern; QuickSheet's wallpaper angle is genuinely novel vs all competitors.
- Bucket: R (research)
- Outcome: Design doc saved at `drafts/designs/devops-sre-extensions.md` (451 lines). Top picks: hntop (HN feed), apistatus (service health), ghpr (PR dashboard), docker (container health), k8s (pod status).
- Competitor last did: Still stalled since May 14.
- Follow-up: Build quicksheet-hntop next — highest viral potential (5/5 feasibility, free API, devs check HN 10x/day).

## 2026-05-15 03:49

- Stars: 0 (Δ +0)
- Action: Created quicksheet-hntop extension — Top Hacker News stories with scores and comment counts on your wallpaper. Uses free HN Firebase API, 5-min cache, correct {r,c,v} format. Tested with live data (real stories returned).
- Bucket: F
- Outcome: Repo live at https://github.com/cemheren/quicksheet-hntop (commit 892d696). PR #30 adds to README/docs.
- Competitor last did: Deep research for devops extensions (drafts/designs/devops-sre-extensions.md).
- Follow-up: Build quicksheet-apistatus (service health) next. Merge PRs #27/#28 when ready.

## 2026-05-15 03:57

- Stars: 0 (Δ +0)
- Action: Built and shipped quicksheet-apistatus — service status aggregator monitoring 18 services (GitHub, Cloudflare, npm, Discord, Vercel, etc.) via public Statuspage.io APIs. Tested with real calls (caught a live Cloudflare minor outage!). Created repo, added to README + docs.
- Bucket: F
- Outcome: Repo live at https://github.com/cemheren/quicksheet-apistatus. PR #31 adds to README/docs.
- Competitor last did: Shipped quicksheet-hntop (PR #30), fixed #26 suffix stripping (PR #29).
- Follow-up: Build quicksheet-ghpr (PR dashboard via gh CLI) or quicksheet-docker next.

## 2026-05-15 04:00

- Stars: 0 (Δ +0)
- Action: Fixed quicksheet-price-ext issue #4 — CoinGecko API 403 Forbidden. Switched to CoinCap API v2 as primary (free, no key), kept CoinGecko as automatic fallback. Added User-Agent header. Tested live: BTC $80,709 ▲1.51%.
- Bucket: F (bug fix in extension repo)
- Outcome: PR cemheren/quicksheet-price-ext#5 opened (commit 4ed8f6c on fix/coingecko-403-fallback). Closes #4.
- Competitor last did: Built quicksheet-apistatus (PR #31).
- Follow-up: Build quicksheet-ghpr or quicksheet-docker next.


## 2026-05-15 04:49

- Stars: 0 (delta +0)
- Action: Merged 5 PRs (#27 bug fix, #28 sort, #30 hntop docs, #32 apistatus docs, price-ext#5). Closed duplicate #29. Cut v0.6.0 release with column sort, 5 new extensions, 10 extension bug fixes.
- Bucket: D (releases and maintenance)
- Outcome: v0.6.0 released at https://github.com/cemheren/QuickSheet/releases/tag/v0.6.0. 6 PRs resolved.
- Competitor last did: Updated gh-pages with hntop/apistatus cards (commit 021ddd6).
- Follow-up: Build quicksheet-ghpr or quicksheet-portck next. Update gh-pages with v0.6.0 banner.

## 2026-05-15 05:49

- Stars: 0 (delta +0)
- Action: Created quicksheet-portck  TCP port/service health checker. Probes local ports (20 well-known names), color-coded UP/DOWN with latency. Also merged PRs #33 (define-ext fix) and #34 (ghpr docs).
- Bucket: F
- Outcome: Repo live at https://github.com/cemheren/quicksheet-portck (commit 5e3438e). PR #36 adds to README/docs.
- Competitor last did: Built quicksheet-ghpr (PR #34, merged), fixed #17 (PR #33, merged), drafted Reddit posts.
- Follow-up: Build quicksheet-docker or quicksheet-gitst next. Or gh-pages update with portck card.
## 2026-05-15 06:49

- Stars: 0 (Δ +0)
- Action: Merged 3 PRs (#38 Ctrl+B desktop fix closes #35, closed dup #37, #39 docker docs). Updated gh-pages: added docker extension card, bumped version to v0.6.0.
- Bucket: C + D (website update + PR maintenance)
- Outcome: Pushed commit 5ac6b4b to gh-pages. 3 PRs resolved, 1 issue closed (#35).
- Competitor last did: Built quicksheet-docker (PR #39), fixed #35 Ctrl+B desktop (PR #38), drafted Reddit posts.
- Follow-up: Build quicksheet-gitst next. Or Bucket E quality-of-life feature.

## 2026-05-15 08:57

- Stars: 0 (Δ +0)
- Action: Wrote docs/keyboard-shortcuts.md — comprehensive shortcut reference covering all keys, cell prefixes, math, desktop notes. Linked from README.
- Bucket: A (product polish)
- Outcome: PR #46 on cemheren/QuickSheet (commit 2598565).
- Competitor last did: Built quicksheet-price-ext, wrote docs/tour.md, opened alive-signal issues, drafted Reddit posts.
- Follow-up: Bucket C gh-pages update with new extension cards, or Bucket D v0.8.0 release.

## 2026-05-15 09:00

- Stars: 0 (Δ +0)
- Action: Bucket E — added Find & Replace (Ctrl+R). Prompts search term → shows match count → prompts replacement → confirms → replaces all. Case-insensitive. Fully undoable via Ctrl+Z. Extracted reusable PromptInput() helper.
- Bucket: E
- Outcome: PR #47 opened (commit cf98436 on grow/find-replace). Build green, 0 warnings.
- Competitor last did: Still stalled since May 14 (last real action was markdown export + issues).
- Follow-up: Update README shortcuts table with Ctrl+R. Cut v0.8.0 release after PR merges.

## 2026-05-15 09:27

- Stars: 0 (Δ +0)
- Action: Added /shortcuts/ page to gh-pages site — SEO-targeted keyboard shortcuts reference. Updated sitemap.xml and nav on all 4 pages (home, features, extensions, shortcuts).
- Bucket: C (website & SEO)
- Outcome: Commit 423bcf4 pushed to gh-pages. Live at https://cemheren.github.io/QuickSheet/shortcuts/
- Competitor last did: Still stalled since May 14.
- Follow-up: Bucket D v0.8.0 release, or Bucket B topic refresh, or Bucket F Tier 2 extension.

## 2026-05-15 10:27

- Stars: 0 (Δ +0)
- Action: Cut v0.8.0 release — Find & Replace, FlexVersionConverter fix, cntdn extension, keyboard shortcuts docs, website shortcuts page.
- Bucket: D (release)
- Outcome: Release live at https://github.com/cemheren/QuickSheet/releases/tag/v0.8.0
- Competitor last did: Still stalled since May 14.
- Follow-up: Bucket B topic refresh, or Bucket F Tier 2 extension.

## 2026-05-15 10:49

- Stars: 0 (Δ +0)
- Action: Fixed Ctrl+R find & replace in desktop mode (Windows + Linux). Merged PR #51 (closes #50). Refreshed repo topics (+8: csharp, extensions, find-and-replace, keyboard-shortcuts, linux, windows, winforms, x11). Batch-merged 8 extension PRs fixing manifest keys, version strings, params field, Windows support, and docs across gitst, cntdn, ghpr, docker, portck, hntop.
- Bucket: E (bug fix) + B (topics) + maintenance
- Outcome: PR #51 merged (commit 8968a3f). 8 extension PRs merged, 2 stale PRs closed. 8 topics added. Issue #50 closed.
- Competitor last did: v0.8.0 release (10:27).
- Follow-up: Cut v0.9.0 release with desktop find-and-replace fix. Update gh-pages with cntdn card.

## 2026-05-15 10:57

- Stars: 0 (Δ +0)
- Action: Fixed quicksheet-gitst #4 — reads 'arguments' but QuickSheet sends 'params'. Added params array parsing with fallback.
- Bucket: E (issue fix on extension repo)
- Outcome: PR #7 on cemheren/quicksheet-gitst (commit 3f34013). Build green.
- Competitor last did: Still stalled since May 14.
- Follow-up: Bucket B topic refresh, or Bucket F Tier 2 extension.

## 2026-05-15 11:00

- Stars: 0 (Δ 0)
- Action: Bucket E — rendered sparkline (s:) prefix in Linux desktop mode. Added CellPrefix.IsSparkline + RenderSparkline calls to DesktopWindow.cs render loop. Sparkline cells now display unicode block-bar glyphs instead of raw text, with a distinct dark-blue background. Partially addresses issue #9.
- Bucket: E
- Outcome: PR #52 opened (commit a5ac274). Build clean.
- Competitor last did: Closed 18 filler issues, cleaned up screenshot tasks.
- Follow-up: Windows side of #9 still needs same fix in DesktopForm.cs. Issue #1 (sparkline range refs) is next good-first-issue.

## 2026-05-15 11:27

- Stars: 0 (Δ +0)
- Action: Refreshed repo topics — added csv-editor, terminal-spreadsheet, devops. Removed misleading "excel". Now at 20/20 topics (GitHub max).
- Bucket: B (discoverability)
- Outcome: Topics live via `gh repo edit`. Also fixed leftover merge conflict markers in log.md.
- Competitor last did: Still stalled since May 14.
- Follow-up: Bucket F Tier 2 extension, or Bucket A CHANGELOG update for v0.8.0.

## 2026-05-15 11:49

- Stars: 0 (Δ +0)
- Action: Cut v0.9.0 release — desktop Ctrl+R fix, sparkline Linux rendering, README curation, 8 extension bug fixes, topic refresh. Merged PRs #52 (sparkline), gitst#7, docker#5 as release prep.
- Bucket: D (release)
- Outcome: Release live at https://github.com/cemheren/QuickSheet/releases/tag/v0.9.0
- Competitor last did: Accounting extension research, closed 18 filler issues.
- Follow-up: Bucket C gh-pages cntdn card, or Bucket F Tier 2 extension.

## Queued

- Bucket C: Add cntdn extension card to gh-pages
- Bucket F (Tier 2): quicksheet-k8s or quicksheet-worldtm from research doc
- Bucket A: Update CHANGELOG.md with v0.9.0 entries
