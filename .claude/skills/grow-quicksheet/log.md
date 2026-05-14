# grow-quicksheet log

Persistent memory across runs. Append-only (except the Queued section at bottom).

## 2026-05-13

- Stars: 0 (baseline)
- Action: Skill created. No promotion action yet — awaiting first scheduled run.
- Bucket: meta
- Outcome: Skill scaffolded at `.claude/skills/grow-quicksheet/`.
- Follow-up: First real run should audit README first-impression and pick top fix.

## 2026-05-14 (remote run)

- Stars: 0 (Δ +0 since baseline)
- Action: README first-impression audit + top fix. Broken Quick Start command (`dotnet run -c Release --desktop` missing `--project ExcelConsole.csproj --`) corrected; replaced self-deprecating "Disclaimers" block ("can't attest for the code quality") with a tighter "A note on the code" that owns the AI assistance while framing the zero-NuGet rule as a strength; sharpened hero paragraph grammar; moved Quick Start up to immediately follow the hero/badges, added TUI-mode hint and startup-app tip.
- Bucket: A
- Outcome: shipped (commit 0aaaac4 on main).
- Follow-up: Set GitHub repo topics + description (Bucket B). Demo GIF still queued (needs human capture).

## 2026-05-14 (local run)

- Stars: 0 (Δ 0 since baseline)
- Action: Set 10 GitHub repo topics (cli, console, csv, desktop-wallpaper, dotnet, excel, productivity, spreadsheet, terminal, tui).
- Bucket: B
- Outcome: Topics live on repo. `gh repo view` confirms.
- Follow-up: Set social preview image (openGraphImage). Set homepage URL once website/docs page exists.

## 2026-05-14 (local run #2)

- Stars: 0 (Δ 0)
- Action: Drafted Show HN post (title options + first-comment pitch + reuse notes for Lobsters/Reddit). Wallpaper-as-spreadsheet framed as hook; zero-NuGet rule + extensions as differentiators.
- Bucket: C
- Outcome: draft saved at `.claude/skills/grow-quicksheet/drafts/showhn.md`. User posts manually.
- Follow-up: Capture a 10s GIF of wallpaper mode to attach in HN replies. Watch HN clock — best window is Tue–Thu, 7–10am Pacific.

## 2026-05-14 (local run #3)

- Stars: 0 (Δ 0)
- Action: Shipped `s: 1,2,3,...` sparkline cell prefix. Parses comma-separated numbers and renders as 8-level unicode block bars (▁▂▃▄▅▆▇█). Added `IsSparkline` + `RenderSparkline` to CellPrefix; SpreadsheetApp render path checks before painting. README and CLAUDE.md updated.
- Bucket: E
- Outcome: build clean (dotnet 9.0.115). Committed and pushed.
- Follow-up: Capture screenshot showing sparkline cells next to data column — high-shareability for HN/social. Could also auto-detect numeric range references later (`s: A1::A10`).

## 2026-05-14 (local run #4)

- Stars: 0 (Δ 0)
- Action: Drafted awesome-list + directory submissions (awesome-tuis, awesome-dotnet, awesome-cli-apps, terminaltrove, console.dev). Each with one-line entry, PR title, PR body, and recommended section. Suggested submission order to avoid burst pattern.
- Bucket: D
- Outcome: draft saved at `.claude/skills/grow-quicksheet/drafts/awesome-lists.md`. User forks and PRs manually.
- Follow-up: After Show HN posts, drafts open for r/commandline + r/dotnet (Bucket C). After list PRs land, link them in README "as seen in" footer.

## 2026-05-14 (local run #5)

- Stars: 0 (Δ 0)
- Action: Added Bucket F (vertical extensions, scaffold + auto-create repo + push). Drafted, created repo, and pushed: `quicksheet-tls-ext` — TLS cert expiry + issuer checker. Targets SRE/sysadmin vertical, zero NuGet (BCL `System.Net.Security`). Builds clean on .NET 9. Linked from main README's Extensions section.
- Bucket: F
- Outcome: Repo live at https://github.com/cemheren/quicksheet-tls-ext. SKILL.md updated to allow autonomous repo create/push (no user interrupt). Main README updated with link.
- Follow-up: Pick next vertical (finance — yfinance JSON, real estate — Zillow API, writing — dictionary). Smoke-test the TLS ext end-to-end inside QuickSheet on next interactive opportunity.

## 2026-05-14 (local run #6)

- Stars: 0 (Δ 0)
- Action: Added top-level CONTRIBUTING.md — quick-start, ground rules (zero NuGet, CSV persistence, cross-platform conditionals, don't break keybindings/CSV/flags), "what's helpful" priority list, PR style. Alive-signal doc.
- Bucket: B
- Outcome: build clean. Committed and pushed.
- Follow-up: GitHub auto-detects CONTRIBUTING.md and surfaces it in the Issues/PR sidebar. Consider opening a starter issue or two labeled "good first issue" next run.

## 2026-05-14 (local run #32)

- Stars: 0 (Δ 0)
- Action: Bucket A — added stock-watchlist recipe (2b) to `docs/recipes.md` using new `stock:` ext. 30min loop cadence matched to Stooq's data refresh rate.
- Bucket: A
- Outcome: Committed and pushed.
- Follow-up: All 12 extensions are now usable from documentation. Future Bucket A: open-source-maintainer dashboard (ping + tls + ping for own repo pages?).

## 2026-05-14 (local run #31)

- Stars: 0 (Δ 0)
- Action: Bucket F — scaffolded, built, created+pushed `quicksheet-stock-ext`. Stock quotes via Stooq's free CSV endpoint. Default `.us` suffix; supports any Stooq exchange suffix. 5min cache. Zero NuGet. Linked from docs.
- Bucket: F
- Outcome: Repo live at https://github.com/cemheren/quicksheet-stock-ext. 12 extensions in directory now.
- Follow-up: Could ship a "watchlist" recipe combining stock + sparkline range (manual close-history) in docs/recipes.md.

## 2026-05-14 (local run #30)

- Stars: 0 (Δ 0)
- Action: Bucket A — added `SECURITY.md` with explicit threat model (in/out of scope), reporting email, dependency-policy note, disclosure terms. GitHub auto-surfaces this in the Issues sidebar and on the security tab.
- Bucket: A
- Outcome: build clean. Committed and pushed.
- Follow-up: Health check signal complete: README, CONTRIBUTING, SECURITY, LICENSE all present.

## 2026-05-14 (local run #29)

- Stars: 0 (Δ 0)
- Action: Bucket B — refined repo description and set homepage URL. New description leads with the differentiator (wallpaper) + .NET 9 + zero NuGet + the extension hook. Homepage points to `docs/tour.md` so first-touch visitors get the 60-second tour.
- Bucket: B
- Outcome: Live via `gh repo edit`. GitHub shows the new description on the repo card and on search results.
- Follow-up: When more verticals ship, swap homepage to a real GitHub Pages site if/when one exists.

## 2026-05-14 (local run #28)

- Stars: 0 (Δ 0)
- Action: Bucket E — added `--version` / `-v` flag. Prints `QuickSheet 0.1.0 (linux|windows, .NET <runtime>)` and exits. Version constant in Program.cs so it bumps in one place. Updated --help to mention it.
- Bucket: E
- Outcome: build clean. Smoke-test passes.
- Follow-up: Bump constant + tag for next release. Could later wire build-time GitVersion if needed.

## 2026-05-14 (local run #27)

- Stars: 0 (Δ 0)
- Action: Bucket F — scaffolded, built, created+pushed `quicksheet-thes-ext`. Thesaurus via free Datamuse API (`rel_syn`). Pairs naturally with `quicksheet-define-ext`. Cached in-memory; zero NuGet. Added to extensions directory and tour.
- Bucket: F
- Outcome: Repo live at https://github.com/cemheren/quicksheet-thes-ext. 11 extensions in directory now.
- Follow-up: Could pair-extend the writer recipe with `thes:` column. Remaining verticals: gravatar, tax/1099, legal (Caselaw). Prebuilt binaries for v0.2.0 still open.

## 2026-05-14 (local run #26)

- Stars: 0 (Δ 0)
- Action: Bucket A — added "Academic writing reference" recipe to `docs/recipes.md` combining `cite:` + `def:` extensions (DOI column auto-fills citation, term column auto-fills definition). Renumbered AI scratchpad to recipe #7.
- Bucket: A
- Outcome: Committed and pushed.
- Follow-up: Recipe page now spans ops, finance, productivity, writing, focus, academic, AI. Could add a "creative" or "media" recipe next.

## 2026-05-14 (local run #25)

- Stars: 0 (Δ 0)
- Action: Bucket F — scaffolded, built, created+pushed `quicksheet-cite-ext`. DOI → citation lookup via Crossref's free API. Normalizes `https://doi.org/`, `doi:`, raw forms. Author list truncated to 3 + "et al.". Cached in-memory. Polite User-Agent header per Crossref etiquette. Zero NuGet. Added to extensions directory and tour.
- Bucket: F
- Outcome: Repo live at https://github.com/cemheren/quicksheet-cite-ext. 10 extensions in directory now.
- Follow-up: Remaining verticals: gravatar, tax/1099, legal (Caselaw), thesaurus. Or pause Bucket F and broaden — recipes for academic-writing dashboard combining cite + define.

## 2026-05-14 (local run #24)

- Stars: 0 (Δ 0)
- Action: Bucket C — drafted dev.to / Medium long-form blog post (~1500 words). Structure: itch → wallpaper trick (Windows WorkerW + Linux X11) → zero NuGet rationale → extensions as git URLs → what it's good for → feedback asks. Includes code snippets (P/Invoke samples). Publishing notes for dev.to / Medium / personal blog + cross-post timing rules.
- Bucket: C
- Outcome: draft saved at `.claude/skills/grow-quicksheet/drafts/devto-blog.md`.
- Follow-up: All major content drafts are now in place (HN, Reddit-cl, Reddit-dotnet, Lobsters, Twitter, Mastodon/Bluesky, dev.to). After user posts a few, observe deltas to inform which channel deserves a second wave.

## 2026-05-14 (local run #23)

- Stars: 0 (Δ 0)
- Action: Bucket B — cut v0.1.0 release. Tagged + pushed v0.1.0, created GitHub release with structured notes (core features, constraints, docs, full extension list, install command, known limits). No prebuilt binaries this release (cross-platform publish is a separate step). GitHub will now surface "Latest release" badge on the repo page.
- Bucket: B
- Outcome: release live at https://github.com/cemheren/QuickSheet/releases/tag/v0.1.0.
- Follow-up: For v0.2.0 — prebuilt self-contained binaries via `dotnet publish -r {linux,win}-x64 --self-contained` attached as release artifacts.

## 2026-05-14 (local run #22)

- Stars: 0 (Δ 0)
- Action: Bucket A — wrote `docs/recipes.md`. Six concrete wallpaper-dashboard recipes with copy-paste CSV blocks: ops on-call (ping + tls + mxck + L: loops), portfolio glance (price), command center (launchers + sparkline range), writer's reference (define), pomodoro+tasks, AI scratchpad (copilot). Tips section on multi-select / L: cadence / Σ Π / CSV portability. Linked from README hero. Concrete content asset for HN/Reddit replies showing the network of extensions in use.
- Bucket: A
- Outcome: build clean. Committed and pushed.
- Follow-up: Add a "did you ship a recipe?" PR-friendly footer to invite community recipes. After dev.to long-form, this page becomes the natural deep-link.

## 2026-05-14 (local run #21)

- Stars: 0 (Δ 0)
- Action: Bucket F — scaffolded, built, created+pushed `quicksheet-ping-ext`. HTTP HEAD (falls back to GET) returns status code + latency. Indicator glyph (✓/⚠/✗). No cache — designed to be polled with `L:`. Zero NuGet. Added to extensions directory and tour.
- Bucket: F
- Outcome: Repo live at https://github.com/cemheren/quicksheet-ping-ext. 9 extensions in directory now.
- Follow-up: gravatar, cite (DOI), tax, legal verticals remain. Could also write a `docs/wallpaper-dashboards.md` showing combined recipes (ping + tls + mxck for an ops dashboard).

## 2026-05-14 (local run #20)

- Stars: 0 (Δ 0)
- Action: Bucket C — drafted Mastodon (500-char single post, hashtag tips, instance recommendations: hachyderm/fosstodon) and Bluesky (300-char 3-post thread) drafts with timing notes, etiquette (alt-text mandate), and cross-promotion sequencing.
- Bucket: C
- Outcome: draft saved at `.claude/skills/grow-quicksheet/drafts/mastodon-bluesky.md`.
- Follow-up: dev.to / Medium long-form blog post (Bucket C). Then re-circulate — wait for user to post some of these and observe star deltas before drafting more.

## 2026-05-14 (local run #19)

- Stars: 0 (Δ 0)
- Action: Bucket E — closed issue #1. Sparkline cell now accepts a cell range: `s: A1::A10` pulls numeric values from the referenced range via GridManager. Non-numeric cells skipped. Literal `s: 1,2,3,...` path unchanged. Renamed local `range` to `sparkRange` to avoid scope collision. SpreadsheetApp callers pass `_grid` through. README, docs/tour.md, --help all updated.
- Bucket: E
- Outcome: build clean. Commit closes #1 via keyword.
- Follow-up: Issue #3 (Wayland) is the remaining open enhancement — needs human with Wayland desktop.

## 2026-05-14 (local run #18)

- Stars: 0 (Δ 0)
- Action: Bucket F — scaffolded, built, created+pushed `quicksheet-mxck-ext`. MX record lookup via Google's DNS-over-HTTPS resolver. 1h cache. Zero NuGet. Added to docs/extensions.md and docs/tour.md.
- Bucket: F
- Outcome: Repo live at https://github.com/cemheren/quicksheet-mxck-ext.
- Follow-up: 8 extensions in directory now. Next verticals: gravatar (email contacts), tax/1099, legal (Caselaw API), cite (DOI formatter), ping (latency check).

## 2026-05-14 (local run #17)

- Stars: 0 (Δ 0)
- Action: Bucket A — wrote `docs/extensions.md`. Live directory page with prefix/name/what-it-does/repo table for all 7 current extensions (copilot, weather, tls, price, define, mortgage, pomodoro). Install command, protocol spec, manifest template, conventions for new ext authors, submission instructions. README's Extensions section now points readers there.
- Bucket: A
- Outcome: build clean. Committed and pushed.
- Follow-up: As new extensions ship, append a row to the table in the same commit. Could later add screenshot thumbnails column.

## 2026-05-14 (local run #16)

- Stars: 0 (Δ 0)
- Action: Bucket E — closed issue #2. Column auto-width in `SpreadsheetApp.GetColumnWidths` now uses rendered length for `s:` sparkline cells instead of raw `s: 1,2,3,...` string. Pure additive — falls back to raw length for everything else and for unparseable sparklines.
- Bucket: E
- Outcome: build clean. Committed, pushed, and closed issue #2 via commit keyword.
- Follow-up: Issue #1 (sparkline range refs) is the next good-first-issue fix; #3 Wayland is the open hard problem.

## 2026-05-14 (local run #15)

- Stars: 0 (Δ 0)
- Action: Bucket C — drafted Lobsters submission and Twitter/X thread (5 tweets + single-tweet variant). Lobsters lead leans hard on technique (WorkerW + X11 + zero NuGet + cell-prefix concept) with a first-comment that solicits pushback on the cell-prefix design. Twitter thread sequences hook → technique → killer feature → ecosystem → CTA, with platform-specific tips and pitfalls noted.
- Bucket: C
- Outcome: drafts saved at `.claude/skills/grow-quicksheet/drafts/lobsters.md` and `twitter-thread.md`. User publishes.
- Follow-up: Mastodon/Bluesky drafts; dev.to / Medium long-form post; ordering reminder — HN first, then Reddit Tue+1, Lobsters Wed+2, Twitter same day as HN or +1.

## 2026-05-14 (local run #14)

- Stars: 0 (Δ 0)
- Action: Bucket F — scaffolded, built, created+pushed `quicksheet-mortgage-ext`. Fixed-rate amortization calculator (monthly payment, total interest, total cost). Pure math, no network, no cache, no state. Zero NuGet. Linked from README + tour.
- Bucket: F
- Outcome: Repo live at https://github.com/cemheren/quicksheet-mortgage-ext.
- Follow-up: Next vertical — email/MX, legal (Caselaw API), tax calc. Smoke-test running extensions inside QuickSheet on user's next interactive session.

## 2026-05-14 (local run #13)

- Stars: 0 (Δ 0)
- Action: Bucket A — added `--help` / `-h` flag. Prints usage (TUI / --desktop / --export-md / --help), cell prefix cheatsheet (r:/i:/s:/L:/ext:/URL), `{A1::C10}` range hint, links to tour and issues. Pure additive — no changes to default behavior.
- Bucket: A
- Outcome: build clean, smoke-test passes. Committed and pushed.
- Follow-up: When new prefixes ship (e.g. issue #1 sparkline ranges), update the help block in same commit.

## 2026-05-14 (local run #12)

- Stars: 0 (Δ 0)
- Action: Bucket E — added headless CSV→Markdown export. `GridManager.SaveToMarkdown(path)` trims trailing empty rows/cols, escapes `|`, treats row 0 as header. New CLI flag `--export-md <out.md>` in Program.cs uses a headless GridManager, no UI launch. Smoke-tested on a sample CSV (quoted fields, empty cells) — output is valid GitHub-flavored markdown. README + CLAUDE.md updated.
- Bucket: E
- Outcome: build clean (0/0). Smoke-test passes. Committed and pushed.
- Follow-up: Could add `--export-md` as a cell-prefix or in-app keybinding later. Could support range scoping. Both deferred.

## 2026-05-14 (local run #11)

- Stars: 0 (Δ 0)
- Action: Bucket F — scaffolded, built, created+pushed `quicksheet-define-ext`. Inline dictionary lookups via free dictionaryapi.dev (no key). Returns one definition per part-of-speech. 24h cache. Zero NuGet. Linked from README + tour.
- Bucket: F
- Outcome: Repo live at https://github.com/cemheren/quicksheet-define-ext.
- Follow-up: Next vertical candidates — email/gravatar/MX, legal/case lookup, mortgage calc. Or smoke-test the running extensions inside QuickSheet on user's next interactive session.

## 2026-05-14 (local run #10)

- Stars: 0 (Δ 0)
- Action: Opened 3 GitHub issues with real specs (no manufactured activity — these are queued TODOs from prior runs). #1 sparkline range refs (good first issue, enhancement). #2 column auto-width over-counts raw chars on `s:`/`i:` cells (good first issue, enhancement). #3 Wayland support investigation (help wanted).
- Bucket: B
- Outcome: Issues live at https://github.com/cemheren/QuickSheet/issues/1, /2, /3. CONTRIBUTING.md now has concrete entry points; alive-signal complete.
- Follow-up: None per-issue. Next bucket: Bucket E (markdown export, theme presets), or Bucket F next vertical (writing/email/legal).

## 2026-05-14 (local run #9)

- Stars: 0 (Δ 0)
- Action: Drafted Reddit posts for r/commandline and r/dotnet. r/commandline lead with UX hook (wallpaper that does something), screenshot strategy, posting timing. r/dotnet lead with technical hook (zero NuGet, WorkerW, X11 P/Invoke, ConPTY) — leans into BCL purity. Each draft includes title alternatives, body, posting tips, and pitfalls to avoid.
- Bucket: C
- Outcome: drafts saved at `.claude/skills/grow-quicksheet/drafts/reddit-commandline.md` and `reddit-dotnet.md`. User posts manually.
- Follow-up: Stagger posts (Show HN first, then r/commandline 1–2 days after, then r/dotnet 3–5 days after). Lobsters draft + Twitter thread next Bucket C run.

## 2026-05-14 (local run #8)

- Stars: 0 (Δ 0)
- Action: Bucket F — scaffolded, built, created repo, and pushed `quicksheet-price-ext`. Live CoinGecko crypto price quotes with 24h change arrow. Built-in ticker alias map (btc/eth/sol/etc) plus raw CoinGecko id pass-through. 60s response cache. Zero NuGet deps. Linked from main README + docs/tour.md.
- Bucket: F
- Outcome: Repo live at https://github.com/cemheren/quicksheet-price-ext. README + tour updated.
- Follow-up: Next Bucket F — writing (define/thesaurus) or email (gravatar/MX). Smoke-test price ext in actual QuickSheet on user's next interactive session.

## 2026-05-14 (local run #7)

- Stars: 0 (Δ 0)
- Action: Wrote `docs/tour.md` — 60-second guided tour. Prefix cheatsheet table (`r:`, `i:`, `s:`, `L:`, `ext:`, URLs), Σ/Π math, `{A1::C10}` references, full extension lineup (copilot, weather, tls, pomodoro), hard rules, "good for / isn't" framing. Linked from README hero section. Doubles as a content asset for HN/Reddit replies and Twitter threads.
- Bucket: A
- Outcome: build clean. Committed and pushed.
- Follow-up: Add docs/tour.md link to the awesome-list submission drafts (one-line entries already done; could mention tour in PR body).

## Queued

- Capture sparkline screenshot for README/social (needs human or `--desktop` smoke test).
- More Bucket E small wins: theme presets, `w: url` live web-fetch prefix, markdown export (each <50 LOC, additive).
- More Bucket F verticals: finance (yfinance JSON), real estate (Zillow), writing (dictionary), email (gravatar/MX).
- Demo GIF of desktop wallpaper mode (needs human capture — deferred).
- Audit screenshot filenames (`image.png`, `image-1.png`, etc.) — give meaningful names and update README refs.
- Set social preview image (openGraphImage) — needs custom upload via web UI or API.
- Draft r/commandline + r/dotnet posts (Bucket C) once Show HN result is known.
- Sparkline could accept range refs (`s: A1::A10`) — open enhancement.
- Open 1–2 "good first issue" stubs (alive signal, after CONTRIBUTING is live).
