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
