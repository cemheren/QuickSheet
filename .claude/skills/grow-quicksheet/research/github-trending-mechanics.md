# GitHub Trending Mechanics & Optimization

## TL;DR

- **Star velocity** (stars/day) is the primary signal, not total stars.
- C# is a **niche language on GitHub** — only ~10–50 stars/day needed to trend in the C# filter (vs 100–1000+ for JS/Python).
- Daily trending resets at **midnight UTC**.
- **Repo metadata** (name, description, topics) directly affects search ranking and discoverability on the trending page.
- A coordinated HN + Reddit + Twitter launch on a **Tuesday–Thursday morning US-Eastern** can spike 20–50+ stars in a day — enough to trend for C#.

## Detailed Findings

### How the Algorithm Works

1. **Star velocity over time window** — GitHub measures how many new stars a repo receives within the daily/weekly/monthly window. A sudden spike outperforms slow accumulation.
2. **Language normalization** — Trending is filtered per-language. C# has far lower competition than JS/Python. QuickSheet can trend for C# with ~10–50 stars/day.
3. **Anti-gaming filters** — Stars from suspicious/new/bot accounts are filtered. Must be organic.
4. **Recent activity signal** — Commits, PRs merged, issues closed, releases cut — all signal "alive" to the algorithm.
5. **Daily reset at midnight UTC** — the "today" counter resets. Stars should be concentrated within a single UTC day for maximum daily-trending impact.

### Thresholds by Language (approximate, 2024-2025 data)

| Language     | Stars/day to trend (daily) | Competition |
|-------------|---------------------------|-------------|
| JavaScript  | 100–1000+                 | Very High   |
| Python      | 50–300+                   | High        |
| **C# / .NET** | **10–50+**             | **Moderate** |
| F#          | 1–10+                     | Very Low    |

### Metadata SEO Factors (ranked by impact)

1. **Repository name** — keyword-rich names rank higher in search.
2. **About/description** — concise, keyword-rich. Surfaces in trending cards.
3. **Topics/tags** — directly influence search results and topic-page inclusion. Changes take effect near real-time.
4. **README** — doesn't directly affect trending rank but is the conversion funnel (visitor → starrer). Quality README = higher star conversion rate.

### Optimal Launch Timing

- **Day of week:** Tuesday–Thursday (highest HN/Reddit traffic).
- **Time:** 6:30–9:30 AM US-Eastern (catches US morning + European afternoon).
- **Relation to UTC reset:** Stars that arrive after midnight UTC "today" count for the new day. A US-morning launch (5:00–14:00 UTC) puts most activity into a single UTC day.
- **Concentration:** All promotional posts (HN, Reddit, Twitter, newsletters) should go out within a 2–3 hour window to maximize same-day velocity.

### Case Studies

- **Dub.sh:** Show HN → 100 to 4,000+ stars in 12 hours. Clean README, instant demo link, responsive maintainer in comments.
- **Turborepo:** Show HN on Tuesday AM ET → 7,000+ stars in 24 hours. Solved trending pain point, clear title, active comment engagement.
- **Pattern:** 90%+ of launch stars arrive in first 24–48 hours. Retention is 70–90%.

### Anti-Patterns (What Doesn't Work)

- Total star count alone (slow accumulation).
- Fake/bot stars (filtered, can hurt ranking).
- Irrelevant keyword stuffing in topics.
- Automated/spammy commits.
- Cross-posting to too many platforms simultaneously (HN community dislikes this).

## Implications for QuickSheet

### 1. C# language filter is the golden path
QuickSheet only needs ~10–50 stars on launch day to trend in the C# daily filter. This is achievable with a single coordinated HN + r/dotnet + r/commandline push. **Action:** Ensure the repo primary language is detected as C# (it should be, given the .csproj).

### 2. Metadata optimization before launch day
**Action (Bucket B, pre-launch):** Ensure description and topics are keyword-optimized:
- Description: "A zero-dependency .NET spreadsheet that runs as your desktop wallpaper — Linux & Windows"
- Topics should include: `csharp`, `dotnet`, `spreadsheet`, `desktop-wallpaper`, `terminal`, `csv`, `linux`, `windows`, `tui`, `extensions`
- Verify these are set: `gh repo view cemheren/QuickSheet --json repositoryTopics,description`

### 3. Launch-day star concentration
**Action:** All promotional posts must go out within the same 2–3 hour window on a Tuesday/Wednesday/Thursday at 7–9 AM ET. The `drafts/launch-day-checklist.md` should explicitly sequence: HN first (7 AM ET), Reddit 30 min later, Twitter/Mastodon/Bluesky 1 hour later. This concentrates star velocity into a single UTC day.

### 4. "Alive" signals before launch
**Action:** Merge a batch of PRs and cut a v1.0.0 release 1–2 days before launch. This ensures the repo shows recent activity (merged PRs, new release) when trending-page visitors arrive. A stale-looking repo with 34 open PRs and no merges looks abandoned.

### 5. README as conversion funnel
Once a visitor lands on the repo from the trending page, they decide to star in ~5 seconds. **Critical:** Demo GIF at the top, one-line value prop, install one-liner. The README is the star-conversion page, not the trending-ranking factor. This is already a known HUMAN BLOCKER (demo GIF needed).
