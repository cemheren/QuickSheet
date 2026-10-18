# Terminal-Tool Influencers & Signal-Boosters

**TL;DR:** 10 accounts/hubs that regularly amplify novel terminal tools. QuickSheet's
"spreadsheet as wallpaper" hook is unusual enough to catch attention if surfaced to the
right 2–3 of these. Priority targets: Terminal Trove (submission form), rothgar/awesome-tuis
(PR), and Charm's Twitter for visibility in the Go/TUI crowd—even though QuickSheet is
.NET, the audience overlap is high.

---

## Tier 1 — Direct submission channels (highest ROI, actionable)

| # | Who / What | Platform | Why they matter | Action |
|---|-----------|----------|-----------------|--------|
| 1 | **Terminal Trove** (terminaltrove.com) | Website + Twitter @TerminalTrove | Curated "tool of the week" for CLI/TUI. Active newsletter. Discovery engine for the exact audience. | Submit via https://terminaltrove.com/post/ — needs name, tagline (100 char), description (250-300 char), 2-3 features, preview image (PNG/GIF), categories, language, license. |
| 2 | **awesome-tuis** (github.com/rothgar/awesome-tuis) | GitHub | 4k+ stars list, first stop for people searching "what TUI tools exist." | Fork → add entry alphabetically: `- [QuickSheet](https://github.com/cemheren/QuickSheet) - Zero-dep .NET spreadsheet that runs as your desktop wallpaper.` → PR. Format: `[Name](url) - Short description.` |
| 3 | **Console.dev** newsletter | Email / Twitter @consoledotdev | Weekly curated 4-6 tools for devs. Editorial picks from submissions. | Email team@console.dev — pitch: zero-dep desktop-wallpaper spreadsheet, unique "lives behind your icons" angle. Include GIF/screenshot. |

## Tier 2 — Individual amplifiers (follow + be visible to)

| # | Who | Handle(s) | Known for | Engagement style |
|---|-----|-----------|-----------|-----------------|
| 4 | **Julia Evans** (@b0rk) | Twitter, Mastodon @b0rk@social.jvns.ca, jvns.ca | Zines, terminal education, "cool tools" posts. 300k+ followers combined. | Doesn't typically boost unsolicited DMs. Best path: appear organically in her feed via HN or awesome-list. If QuickSheet lands on HN front page, she'll notice. |
| 5 | **Simon Willison** (@simonw) | Twitter, Mastodon, simonwillison.net | CLI tools, datasette, LLM CLI. Regularly shares novel terminal tools. | Actively tracks "interesting CLI" tag. May notice QuickSheet if framed as "data tool in terminal" rather than "spreadsheet." Blog tag: simonwillison.net/tags/cli/ |
| 6 | **Will McGugan** (@willmcgugan) | Mastodon @willmcgugan@mastodon.social, Twitter | Rich/Textual creator. 50k+ combined. Boosts TUI innovations. | Engages with novel TUI demos. Even though QuickSheet isn't Python, the "wallpaper embedding" trick is visually interesting enough to catch his eye on Mastodon. |
| 7 | **Christian Rocha** (Charm/BubbleTea) | Twitter @charmcli, charm.sh | Go TUI framework. Company Twitter amplifies cool TUI projects in any language. | Tag @charmcli when posting — they retweet novel TUIs regardless of language if the demo looks good. |

## Tier 3 — Community hubs (passive visibility)

| # | Who / What | Platform | Notes |
|---|-----------|----------|-------|
| 8 | **Jesse Duffield** (@jesseduffield) | GitHub, Twitter | lazygit/lazydocker creator. Engages with other TUI devs. Less active signal-boosting of others, but high credibility if he stars/comments. |
| 9 | **Saul Pwanson** (@saulpw) | Twitter, GitHub | VisiData creator. Exact adjacent space (terminal spreadsheet). Might see QuickSheet as complementary (wallpaper angle ≠ data-wrangling angle). |
| 10 | **r/commandline** moderators + power users | Reddit | 500k+ subscribers. Frequent "what cool tools do you use" threads. Self-promotion allowed if genuinely useful. Best as follow-up 1-2 days after a Show HN lands. |

---

## Competitive landscape note (HN June 2025)

Recent HN threads show active demand for TUI spreadsheets:
- "Sheets: Terminal based spreadsheet tool" (HN item 47636456, June 2025) — strong praise.
- "TUI spreadsheet app in Rust based on IronCalc engine" (HN item 42282190, Dec 2024).
- Comment: "Oh man, a TUI spreadsheet application that can edit ODF or XLSX format would be absolutely killer" (HN item 47456938).

QuickSheet's differentiator vs these: **wallpaper mode.** None of the competitors embed
as the desktop background. This is the hook. Every submission/post should lead with
"spreadsheet that IS your wallpaper" not "another TUI spreadsheet."

---

## Implications for QuickSheet

1. **Submit to Terminal Trove immediately** (Bucket D draft). This is the highest-ROI
   single action: free, editorial-curated, exact audience, low effort. Requires a
   preview image (PNG/GIF). Queue: draft the submission text now; user submits once a
   wallpaper screenshot exists.

2. **PR to awesome-tuis** (Bucket D draft). One-line addition, low friction. Draft the
   fork-PR body. User submits after screenshot lands.

3. **Frame ALL outreach around "wallpaper" not "spreadsheet."** The TUI spreadsheet
   space is getting crowded (Rust/IronCalc, VisiData, sc-im). QuickSheet wins on the
   novelty of embedding behind desktop icons, not on formula power. Every pitch
   should lead: "A spreadsheet that replaces your desktop wallpaper."

4. **Console.dev email pitch is a 5-minute draft** — save it alongside the Terminal
   Trove draft for user to send when screenshot is ready.
