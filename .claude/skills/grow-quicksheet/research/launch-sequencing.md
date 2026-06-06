# Launch Sequencing & GitHub Trending Mechanics

> Research brief: how to time the QuickSheet launch for maximum trending signal.

## TL;DR

- GitHub Trending uses **star velocity** (spike vs. baseline), not absolute count.
- For C# language filter: as few as **10–20 stars/day** can trend if baseline is ~0.
- For "All Languages" daily: need **100–500 stars in 24h**.
- Best strategy: **coordinate all platforms within a 6-hour window** on one day.
- Best HN slot: **Tuesday/Wednesday 14:00–15:00 UTC** (9–10am EST).
- Sunday 12:00–14:00 UTC has highest "breakout" rate but smaller professional audience.

## GitHub Trending Algorithm — Key Signals

1. **Star velocity** — spike relative to repo's own baseline is the primary signal.
2. **Fork rate** — secondary but contributes.
3. **Recent activity** — commits, issues, PRs signal "alive" project.
4. **Unique contributors** — multiple contributors > solo dev (perception).
5. **Repo age** — newer repos or dormant-then-spiking repos favored.
6. **Anti-abuse** — coordinated mass-starring detected and filtered.
7. **Language scoping** — C# has far less competition than JS/Python; threshold is lower.

## Quantitative Thresholds (2024–2025 data)

| Target | Stars needed in 24h |
|--------|---------------------|
| C# language daily | ~10–30 (low competition) |
| All Languages daily | 100–500 |
| All Languages #1 | 500+ |
| Sustained weekly | 1000+ in 72h for lasting presence |

Source: community analysis of GitHub trending, arxiv paper on HN→GitHub star diffusion, AFFiNE 33k-star case study.

## Optimal Launch Sequence for QuickSheet

### Pre-flight (before launch day)
1. **Merge all 13 pending PRs** — gives a burst of commit activity + "alive" signal.
2. **Cut first GitHub Release** (v1.0.0 tag → release workflow fires → binaries available). People star at download time.
3. **Ensure README has**: demo GIF/screenshot, install one-liner, clear value prop in first 2 lines.
4. **Seed 5–10 organic stars** from friends/colleagues in the 48h before launch (sets a non-zero baseline so the spike looks organic, not bot-like).

### Launch Day (T+0) — Recommended: Tuesday or Wednesday

| Time (UTC) | Action |
|------------|--------|
| 13:00 | Final README check, ensure GH release is live |
| 13:30 | Tweet/X thread (from personal account) |
| 14:00 | **Show HN** post — title: "Show HN: QuickSheet – A spreadsheet that *is* your desktop wallpaper (.NET, zero deps)" |
| 14:30 | Reddit r/commandline post |
| 15:00 | Reddit r/dotnet post |
| 15:30 | Reddit r/linux post (if X11 screenshot available) |
| 16:00 | Reddit r/selfhosted post (homelab angle) |
| 16:30 | Mastodon/Bluesky post |
| 17:00 | Lobsters submission |
| Evening | Dev.to blog post goes live |

### Post-Launch (T+1 to T+7)
- **Respond to every HN comment** within 1h for first 4h (critical for ranking).
- **r/unixporn rice post** on day 2–3 (different audience, extends the tail).
- **Submit to awesome-lists** on day 3–4 (awesome-windows, awesome-linux-software).
- **Newsletter pitches** (Console Weekly, TerminalTrove) on day 2.
- **Product Hunt** on day 5 if still gaining momentum (optional, less dev-focused).

## Why This Works for QuickSheet Specifically

- **C# language filter has low competition** — 10–30 stars/day likely lands on language-specific trending. Most C# trending repos are well-known (Roslyn, .NET runtime). A fresh project spiking even modestly will stand out.
- **"Desktop wallpaper" is visually novel** — screenshots/GIFs will carry the HN and Reddit posts. The "spreadsheet on the wallpaper" concept is inherently screenshot-worthy.
- **Extension ecosystem is a discussion magnet** — people will ask "can it do X?" and the answer is usually "yes, there's an extension." This drives HN comment engagement, which drives ranking.
- **Zero NuGet deps is a talking point** — HN loves supply-chain-aware design decisions.

## Risks & Mitigations

| Risk | Mitigation |
|------|-----------|
| No demo GIF available | User must capture before launch. Blocker. |
| 0 existing stars looks "dead" | Seed 5–10 from inner circle 48h before |
| HN flagged as self-promotion | Use honest "Show HN" format; don't vote-ring |
| Reddit shadowban from multi-posting | Space posts 30–60 min apart, different angles each sub |
| Release workflow untested | Merge CI PR, test tag push on a throw-away tag first |

## Implications for QuickSheet — Concrete Next Actions

1. **Merge the 13 PRs first.** Every feature adds to the "alive" signal and makes the README richer for launch day. Batch-merge over 1–2 days.
2. **Capture a demo GIF/screenshot.** This is THE blocker for launch. No visual = no virality. Needs human with X11 or Windows desktop.
3. **Plan a single "launch Tuesday."** Pick a date, do all the platform posts in one 4-hour window. Don't drip — spike is the goal.
4. **C# trending is achievable at current scale.** Even 15 stars in a day from organic HN/Reddit traffic should land QuickSheet on the C# daily trending page. That provides a flywheel of organic discovery.
5. **Queue "milestone" tweet/post for 100 stars** — social proof compounds. Have the draft ready.

## Sources

- GitHub Community Discussion: "Curious about trending repos calculations" (https://github.com/orgs/community/discussions/163970)
- Myriade: "Best Time to Post on Show HN" (https://www.myriade.ai/blogs/when-is-it-the-best-time-to-post-on-show-hn)
- HN Best Posting Times analysis (https://github.com/JoseMarquezAlberti/hn-best-posting-times)
- FlowJam: "How to Get on the Front Page of HN in 2025" (https://www.flowjam.com/blog/how-to-get-on-the-front-page-of-hacker-news-in-2025-the-complete-up-to-date-playbook)
- arxiv: "Launch-Day Diffusion: Tracking HN Impact on GitHub Stars" (https://arxiv.org/html/2511.04453v1)
- AFFiNE growth case study (dev.to: "How We Grew from 0 to 60K")
