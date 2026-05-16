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

## 2026-05-15 11:57

- Stars: 0 (Δ +0)
- Action: Updated CHANGELOG.md with v0.8.0 entries — Find & Replace, FlexVersionConverter, shortcuts docs, cntdn, extension fixes.
- Bucket: A (product polish)
- Outcome: PR #53 on cemheren/QuickSheet (commit 75d80d9).
- Competitor last did: Accounting extension research for #15.
- Follow-up: Bucket F Tier 2 extension next.

## 2026-05-15 12:00

- Stars: 0 (Δ +0)
- Action: Created quicksheet-worldtm extension — multi-timezone world clock with 40+ aliases (NY, London, PST, etc.), business-hours indicators (🟢🟡🔴), time-of-day icons, fuzzy matching. Pure local, zero network calls.
- Bucket: F
- Outcome: Repo live at https://github.com/cemheren/quicksheet-worldtm. PR #54 adds to README/docs.
- Competitor last did: Accounting extension research for #15, closed 18 filler issues.
- Follow-up: Bucket C — add worldtm + cntdn cards to gh-pages. Or quicksheet-k8s next.

## 2026-05-15 12:27

- Stars: 0 (Δ +0)
- Action: Added 3 new extension cards to gh-pages extensions page — countdown (cntdn), world clock (worldtm), IRS mileage. Updated count 28→31, SEO keywords.
- Bucket: C (website & SEO)
- Outcome: Commit c362626 pushed to gh-pages branch.
- Competitor last did: Stalled since May 14 — last action was quicksheet-price-ext and docs/tour.md.
- Follow-up: Bucket F quicksheet-k8s, or Bucket E safe feature.

## 2026-05-15 12:49

- Stars: 0 (Δ +0)
- Action: Created quicksheet-k8s extension — live Kubernetes pod status from kubeconfig. Color-coded status icons (🟢🔴🟡🟠), namespace support, truncated pod names, kubectl subprocess. Also merged PRs #52-55 as prep.
- Bucket: F
- Outcome: Repo live at https://github.com/cemheren/quicksheet-k8s (commit 710f9ec). PR #57 adds to README/docs.
- Competitor last did: Created quicksheet-mileage-ext (PR #55, merged). Accounting research for #15.
- Follow-up: Bucket C — add k8s card to gh-pages. Or quicksheet-margin-ext next.


## 2026-05-15 12:57

- Stars: 0 (Δ +0)
- Action: Fixed 2 extension manifest bugs — worldtm#1 (entrypoint→entry), mileage-ext#1 (missing prefix field + enriched manifest).
- Bucket: E (issue fixes)
- Outcome: PR #2 on cemheren/quicksheet-worldtm (84bda51). PR #2 on cemheren/quicksheet-mileage-ext (ac89b6c).
- Competitor last did: Created quicksheet-k8s, merged PRs #52-55.
- Follow-up: Bucket C add k8s card to gh-pages, or accounting extensions.

## 2026-05-15 13:00

- Stars: 0 (Δ +0)
- Action: Created quicksheet-margin-ext — break-even & contribution margin calculator. Pure math, zero NuGet. Color-coded health indicators (🟢🟡🟠🔴). Added to README + docs/extensions.md.
- Bucket: F (accounting)
- Outcome: Repo live at https://github.com/cemheren/quicksheet-margin-ext (commit fac8e1c). PR #58 adds to main repo docs.
- Competitor last did: Created quicksheet-k8s, worldtm manifest fixes.
- Follow-up: quicksheet-depr-ext next for #15. Or Bucket C to add margin card to gh-pages.


## 2026-05-15 13:27

- Stars: 0 (Δ +0)
- Action: Added k8s + margin-ext cards to gh-pages extensions page. Count 31→33, new SEO keywords.
- Bucket: C (website & SEO)
- Outcome: Commit b2e1306 pushed to gh-pages.
- Competitor last did: Created quicksheet-margin-ext (PR #58). Stalled on Claude side since May 14.
- Follow-up: Bucket F quicksheet-depr-ext for #15, or Bucket D release.

## 2026-05-15 13:49

- Stars: 0 (Δ +0)
- Action: Cut v0.10.0 release — desktop Ctrl+R, sparkline Linux, 4 new extensions (worldtm, k8s, mileage, margin), mass extension bug fixes, README curation, topic refresh. Merged PRs #57, #58, #59. Closed #56 (conflicts).
- Bucket: D (release)
- Outcome: Release live at https://github.com/cemheren/QuickSheet/releases/tag/v0.10.0. CHANGELOG updated via PR #59.
- Competitor last did: Created margin-ext, gh-pages cards. Stalled on Claude side since May 14.
- Follow-up: Bucket F quicksheet-depr-ext for #15, or Bucket E safe feature.



- Bucket F: quicksheet-depr-ext (depreciation tables) to close #15
- Bucket E: safe additive feature (status bar, theme presets)
- Bucket A: CONTRIBUTING.md or issue templates


## 2026-05-15 13:57

- Stars: 0 (Δ +0)
- Action: Created quicksheet-depr-ext — straight-line & MACRS depreciation calculator. IRS Pub 946 tables (3/5/7/10/15/20-yr), salvage value support, auto-maps to nearest MACRS class. Build-tested. Added to README. Completes accounting suite for #15.
- Bucket: F (accounting)
- Outcome: Repo live at https://github.com/cemheren/quicksheet-depr-ext (10b54a5). PR #60 adds to README.
- Competitor last did: Created quicksheet-k8s, margin-ext. Stalled on Claude side since May 14.
- Follow-up: Bucket C add depr-ext card to gh-pages. Bucket D v0.10.0 release.


## 2026-05-15 14:27

- Stars: 0 (Δ +0)
- Action: Added depr-ext card to gh-pages extensions page. Count 33→34. Added MACRS keyword to SEO.
- Bucket: C (website & SEO)
- Outcome: Commit 7cf4ad0 pushed to gh-pages.
- Competitor last did: Added cell color prefix (c:COLOR:) feature, PR #61.
- Follow-up: Bucket A document c: prefix in keyboard-shortcuts. Bucket D v0.10.0 release.

## 2026-05-15 14:49

- Stars: 0 (Δ +0)
- Action: Documented c:color: cell prefix in keyboard-shortcuts.md (full color table) and tour.md. Merged PRs #61 (color prefix feature), #62 (depr-ext links), #63 (FAQ page). Closed #60 (conflicts).
- Bucket: A (product polish)
- Outcome: PR #64 on cemheren/QuickSheet (commit fe860cc).
- Competitor last did: Added c:color: prefix feature (PR #61), wrote FAQ, linked accounting extensions.
- Follow-up: Bucket C gh-pages color feature highlight, or Bucket E safe feature.



- Bucket C: Add color-prefix feature highlight to gh-pages
- Bucket E: safe additive feature (status bar, markdown export)
- Bucket B: CONTRIBUTING.md or issue templates


## 2026-05-15 14:57

- Stars: 0 (Δ +0)
- Action: Fixed quicksheet-depr-ext#1 — 3 protocol bugs (init/activate types, r/c/v cell format).
- Bucket: E (issue fix)
- Outcome: PR #2 on cemheren/quicksheet-depr-ext (2541d68). Build-tested.
- Competitor last did: Stalled since May 14.
- Follow-up: Bucket D v0.11.0 release. Bucket A document c: prefix.

## 2026-05-15 15:00

- Stars: 0 (Δ +0)
- Action: Added c:color: feature highlight to gh-pages site — feature card on homepage, prefix entry on features & shortcuts pages, comparison table row, version bump to 0.10.0, SEO keywords for cell coloring.
- Bucket: C (website & SEO)
- Outcome: Commit 6274562 pushed to gh-pages.
- Competitor last did: Documented c:color: in docs, merged PRs #61-63, cut v0.10.0 release.
- Follow-up: Bucket D v0.11.0 release. Bucket E safe additive feature.


## 2026-05-15 15:27

- Stars: 0 (Δ +0)
- Action: Merged PR #64 (c:color: docs) and cut v0.11.0 release — cell colors, FAQ, depreciation ext, accounting suite complete.
- Bucket: D (release)
- Outcome: Release live at https://github.com/cemheren/QuickSheet/releases/tag/v0.11.0
- Competitor last did: Added c:color: site highlight, merged PRs #61-63, cut v0.10.0.
- Follow-up: Bucket E safe feature (markdown export). Bucket F new vertical.

## 2026-05-15 15:49

- Stars: 0 (Δ +0)
- Action: Fixed #65 — c:color: prefix not rendered in Windows desktop mode. Added CellPrefix.ParseColor() to DesktopForm cell rendering, ConsoleColorToBg() helper mapping ConsoleColor→RGB (matching Linux impl). Strips prefix from display, shows colored background.
- Bucket: E (issue fix)
- Outcome: PR #67 on cemheren/QuickSheet (commit f2f3b54). Build verified.
- Competitor last did: Cut v0.11.0 release, color site highlight, merged PRs.
- Follow-up: Merge PR #67. Bucket D v0.12.0 release with color desktop fix.

## 2026-05-15 16:00

- Stars: 0 (Δ +0)
- Action: Created quicksheet-jwtdec extension — JWT decoder that runs 100% locally. Privacy-first alternative to jwt.io. Decodes header + claims, annotates timestamps, flags expired tokens, labels well-known claims. Zero network, zero NuGet.
- Bucket: F (devops/security vertical)
- Outcome: Repo live at https://github.com/cemheren/quicksheet-jwtdec. PR #69 adds to README/docs.
- Competitor last did: Stalled since May 14.
- Follow-up: Bucket C — add jwtdec card to gh-pages. Bucket D — v0.12.0 release.

## 2026-05-15 16:49

- Stars: 0 (Δ +0)
- Action: Merged PRs #67 (color desktop Win), #68 (Linux shortcuts), #69 (jwtdec docs), #71 (color desktop dup). Merged extension fixes: depr-ext#2, mileage-ext#2, worldtm#2. Closed dup PRs #70, ext#3s. Added CHANGELOG entries for v0.11.0+v0.12.0 (PR #72). Cut v0.12.0 release.
- Bucket: D (release)
- Outcome: Release live at https://github.com/cemheren/QuickSheet/releases/tag/v0.12.0
- Competitor last did: Added jwtdec to gh-pages, created Linux shortcuts PR #68, Windows color PR #71.
- Follow-up: Bucket F new extension. Bucket B CONTRIBUTING.md. Bucket E safe feature.

## 2026-05-15 17:49

- Stars: 0 (Δ +0)
- Action: Added 5 missing keyboard shortcuts to Windows desktop mode — Ctrl+Z (undo), Ctrl+Y (redo), Ctrl+T (theme cycle), Ctrl+G (goto cell with text input), Ctrl+H (help overlay). Merged PR #73 (cronck docs), closed #74 (conflicts).
- Bucket: E (issue fix, refs #66)
- Outcome: PR #76 on cemheren/QuickSheet (commit 11ba2e4). Build verified.
- Competitor last did: Created quicksheet-cronck and quicksheet-gitlog extensions.
- Follow-up: Merge PR #76. Bucket B CONTRIBUTING.md. Bucket C add gitlog+cronck to gh-pages.

## 2026-05-15 18:49

- Stars: 0 (Δ +0)
- Action: Merged PR #76 (Windows desktop Ctrl+Z/Y/T/G/H). Updated gh-pages: added Help Overlay + Find & Replace feature cards, bumped version 0.10.0→0.12.0, added SEO keywords (keyboard shortcuts, find and replace, help overlay, undo redo).
- Bucket: C (website & SEO)
- Outcome: Commit b5f8404 pushed to gh-pages. PR #76 merged.
- Competitor last did: Created cronck + gitlog extensions, added to gh-pages.
- Follow-up: Bucket D v0.13.0 release. Bucket A README badges. Bucket F new extension.

## 2026-05-15 19:49

- Stars: 0 (Δ +0)
- Action: Fixed #77 — Ctrl+T theme cycling had no visible effect in Windows desktop mode. Replaced hardcoded Color.Black/White in DesktopForm.OnPaint with Theme.Current-derived colors. Added ConsoleColorToRgb() for full 16-color mapping. Theme now drives background, foreground, headers, selection, search, status bar.
- Bucket: E (issue fix)
- Outcome: PR #78 on cemheren/QuickSheet (commit f0f713e). Build verified.
- Competitor last did: Stalled (last entry was tour.md on May 14).
- Follow-up: Merge PR #78. Bucket D v0.13.0 release. Close #66 (all shortcuts now done).

## 2026-05-15 20:49

- Stars: 0 (Δ +0)
- Action: Merged PR #78 (theme desktop). Closed #66 (all shortcuts done). Fixed gitlog#1 (protocol bugs, PR #2). Merged 4 extension protocol fix PRs (cronck#2, jwtdec#2, rate#2, gitlog#2). Cut v0.14.0 release.
- Bucket: D (release + maintenance)
- Outcome: Release live at https://github.com/cemheren/QuickSheet/releases/tag/v0.14.0. 5 PRs merged, 2 issues closed, 4 extension bugs fixed.
- Competitor last did: Stalled since May 14.
- Follow-up: Bucket C gh-pages update with v0.14.0. Bucket A README badges. Bucket F new extension.

## 2026-05-15 21:49

- Stars: 0 (Δ +0)
- Action: Updated gh-pages site for v0.14.0. Bumped structured data version 0.12.0→0.14.0. Added "Full Desktop Parity" feature card on homepage. Updated theme description to mention desktop wallpaper mode. Updated features page theme section.
- Bucket: C (website & SEO)
- Outcome: Commit 18a096e pushed to gh-pages.
- Competitor last did: Stalled since May 14.
- Follow-up: Bucket A README badges. Bucket F new extension. Bucket E markdown export.

## 2026-05-15 22:49

- Stars: 0 (Δ +0)
- Action: Fixed #9 (Windows side) — sparkline (s:) prefix now renders as unicode block-bar glyphs in Windows desktop mode. Added IsSparkline + RenderSparkline to DesktopForm cell rendering, with dark-blue bg and light-blue fg matching Linux.
- Bucket: E (issue fix)
- Outcome: PR #79 on cemheren/QuickSheet (commit 585748d). Build verified.
- Competitor last did: Stalled since May 14.
- Follow-up: Merge PR #79, close #9. Bucket A README badges. Bucket F new extension.

## 2026-05-15 23:49

- Stars: 0 (Δ +0)
- Action: Merged PR #79 (sparkline desktop), closed #9. Added 3 README badges: GitHub release (dynamic), zero dependencies, 30+ extensions.
- Bucket: A (product polish) + D (merge)
- Outcome: PR #80 on cemheren/QuickSheet (commit c5c5747). Issue #9 closed.
- Competitor last did: Stalled since May 14.
- Follow-up: Merge PR #80. Bucket D v0.15.0 release. Bucket F new extension.

## 2026-05-16 00:49

- Stars: 0 (Δ +0)
- Action: Merged PR #80 (badges). Fixed cronck#3 (params instead of cells, PR #4, merged). Cut v0.15.0 release — sparkline desktop, README badges, cronck fix.
- Bucket: D (release + maintenance)
- Outcome: Release live at https://github.com/cemheren/QuickSheet/releases/tag/v0.15.0. 2 PRs merged, 1 extension bug fixed.
- Competitor last did: Stalled since May 14.
- Follow-up: Bucket F new extension. Bucket C gh-pages update. Bucket E markdown export.

## Queued

- Bucket F: new extension (data science? writing tools?)
- Bucket C: gh-pages sparkline card + v0.15.0 bump
- Bucket E: markdown export enhancement
- Bucket B: CONTRIBUTING.md refresh

## 2025-07-17 16:27

- Stars: 0 (Δ +0)
- Action: Added 4 missing keyboard shortcuts to Linux desktop mode (Ctrl+Z/Y/T/G)
- Bucket: E (code)
- Outcome: PR #68 opened (commit b7dd957). Build green before and after.
- Details: Ctrl+Z=undo, Ctrl+Y=redo, Ctrl+T=theme cycle, Ctrl+G=goto cell with input prompt. Added XK_z/t/g keysym constants. Partially addresses #66 (4 of 5; help overlay deferred).
- Competitor last did: stalled since May 14
- Follow-up: Ctrl+H help overlay for desktop, or new extension vertical

## Queued
- Bucket E: Help overlay for desktop mode (completes #66)
- Bucket F: New extension — data science or cloud cost vertical
- Bucket C: Add more extension cards to gh-pages

## 2026-05-15 16:27

- Stars: 0 (Δ +0)
- Action: Added jwtdec extension card to gh-pages, updated count 34→35, added JWT SEO keywords
- Bucket: C (website & SEO)
- Outcome: Commit fc25df0 pushed to gh-pages
- Competitor last did: Stalled since May 14
- Follow-up: Bucket D v0.12.0 release. Bucket F new extension vertical.

## Queued
- Bucket D: v0.12.0 release (color desktop fix + jwtdec + shortcuts)
- Bucket F: quicksheet-cronck (cron parser) or quicksheet-gitlog (recent commits)
- Bucket A: README badges or first-run UX

## 2026-05-15 16:57

- Stars: 0 (Δ +0)
- Action: Created quicksheet-cronck extension — cron expression parser with human-readable output
- Bucket: F (extension)
- Outcome: Repo live at https://github.com/cemheren/quicksheet-cronck (commit ac728f2). PR #73 adds to README/docs.
- Details: 5-field cron → text. Ranges, lists, steps, named days/months, common presets. Build-tested, smoke-tested.
- Competitor last did: Stalled since May 14
- Follow-up: Bucket C — add cronck card to gh-pages. Bucket A — README badges.

## Queued
- Bucket C: add cronck card to gh-pages extensions page
- Bucket A: README badges or first-run UX improvements
- Bucket F: quicksheet-gitlog (recent commits in cells) or quicksheet-envck (env var viewer)

## 2026-05-15 17:00

- Stars: 0 (Δ +0)
- Action: Bucket F — created quicksheet-gitlog extension. Shows recent git commits in cells (hash, author, relative time, message). Supports custom repo path and commit count (up to 50). Zero NuGet deps. Build verified, smoke-tested against live repo.
- Bucket: F
- Outcome: Repo live at https://github.com/cemheren/quicksheet-gitlog. PR #74 adds to README + docs/extensions.md.
- Competitor last did: Stalled since May 14.
- Follow-up: Bucket B — CONTRIBUTING.md or issue templates. Bucket E — status bar improvement.

## Queued
- Bucket B: CONTRIBUTING.md (alive signal, lowers contribution barrier)
- Bucket E: Status bar with file name + cell count + mode indicator
- Bucket C: Add gitlog card to gh-pages
- Bucket A: README badges (build status, .NET version, license)

## 2026-05-15 17:27

- Stars: 0 (Δ +0)
- Action: Added cronck + gitlog extension cards to gh-pages, updated count 35→37, added cron/git SEO keywords
- Bucket: C (website & SEO)
- Outcome: Commit 4981f53 pushed to gh-pages
- Competitor last did: Created quicksheet-gitlog extension (17:00)
- Follow-up: Bucket B — CONTRIBUTING.md. Bucket A — README badges.

## Queued
- Bucket B: CONTRIBUTING.md (alive signal)
- Bucket A: README badges (build status, .NET version, license)
- Bucket D: v0.13.0 release once PRs #73 #74 merged

## 2026-05-15 17:57

- Stars: 0 (Δ +0)
- Action: Cut v0.13.0 release — cronck + gitlog extensions, colorParsed bug fix, 37 extension cards on site
- Bucket: D (release)
- Outcome: Release live at https://github.com/cemheren/QuickSheet/releases/tag/v0.13.0
- Competitor last did: Opened PR #76 (Windows desktop shortcuts), closed PR #74
- Follow-up: Bucket A — README polish or docs. Bucket F — new extension vertical.

## Queued
- Bucket A: README "Why this exists" section or feature comparison table
- Bucket F: quicksheet-envck (env var viewer) or quicksheet-b64 (base64 encode/decode)
- Bucket E: Status bar improvements (file name, cell count)

## 2026-05-15 18:00

- Stars: 0 (Δ +0)
- Action: Fixed quicksheet-cronck#1 — protocol bugs (init_response→register, activate_response→write+id)
- Bucket: A (ext repo issue fix)
- Outcome: PR #2 on cemheren/quicksheet-cronck (commit 816aaa2). Build verified, smoke-tested.
- Competitor last did: Cut v0.13.0 release, added cronck/gitlog cards to gh-pages
- Follow-up: Fix quicksheet-gitlog#1 (same protocol bug). Then Bucket B CONTRIBUTING.md.

## 2026-05-16 01:49

- Stars: 0 (Δ +0)
- Action: Created quicksheet-news — RSS/Atom feed headlines extension with 15+ built-in aliases (HN, Reddit, Lobsters, dev.to, BBC, TechCrunch, etc.)
- Bucket: F (new extension)
- Outcome: Repo live at https://github.com/cemheren/quicksheet-news (commit 8fd0040). PR #81 on main repo to add to extensions directory.
- Competitor last did: Cut v0.15.0 release, added README badges (PR #80)
- Follow-up: Merge PR #81. Bucket C — gh-pages sparkline card + v0.15.0 bump. Bucket B — CONTRIBUTING.md.

## 2026-05-16 02:49

- Stars: 0 (Δ +0)
- Action: gh-pages update — added 9 missing extension cards (cntdn, worldtm, margin, mileage, depr, jwtdec, cronck, k8s, news), bumped version to 0.15.0, updated comparison table to "40+ extensions"
- Bucket: C (website/SEO)
- Outcome: Committed b8b6a75 to gh-pages. Also merged PR #81 (news extension docs).
- Competitor last did: Created quicksheet-news extension (run 01:49)
- Follow-up: Bucket B — CONTRIBUTING.md refresh. Bucket D — v0.16.0 release.

## 2026-05-16 03:49

- Stars: 0 (Δ +0)
- Action: Refreshed CONTRIBUTING.md — architecture diagram, extension authoring guide, protocol tips, 40+ extensions mention, scannable ground rules table
- Bucket: B (discoverability / alive signal)
- Outcome: PR #82 (commit 3a7a92e)
- Competitor last did: gh-pages update with 9 extension cards + v0.15.0 bump (run 02:49)
- Follow-up: Bucket D — v0.16.0 release. Bucket E — status bar or markdown export.

## 2026-05-16 04:49

- Stars: 0 (Δ +0)
- Action: Cut v0.16.0 release — news extension, CONTRIBUTING refresh, CHANGELOG catch-up (v0.13–v0.16), merged PRs #82 + #83
- Bucket: D (release)
- Outcome: Release live at https://github.com/cemheren/QuickSheet/releases/tag/v0.16.0
- Competitor last did: CONTRIBUTING.md refresh (run 03:49)
- Follow-up: Bucket E — status bar or markdown export. Bucket F — new extension.

## 2026-05-16 05:49

- Stars: 0 (Δ +0)
- Action: Created quicksheet-b64 — base64 encode/decode extension with auto-detect, explicit modes, multi-line output
- Bucket: F (new extension)
- Outcome: Repo live at https://github.com/cemheren/quicksheet-b64 (commit ff99271). PR #84 on main repo for docs.
- Competitor last did: Cut v0.16.0 release (run 04:49)
- Follow-up: Merge PR #84. Bucket E — status bar improvements. Bucket A — README "Why this exists".

## Queued
- Merge PR #84 (b64 extension docs)
- Bucket E: Status bar improvements or markdown export enhancement
- Bucket A: README "Why this exists" section
- Bucket C: gh-pages v0.16.0 version bump + b64/news extension cards
