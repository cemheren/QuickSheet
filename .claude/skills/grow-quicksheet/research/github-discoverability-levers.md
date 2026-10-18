# GitHub Discoverability Levers for QuickSheet

**TL;DR:** Star velocity (not total stars) drives Trending. For C#, ~10-20 new stars/day gets you on the daily page; 40+ puts you near the top. The highest-leverage untapped channel is submitting to GitHub Explore's `productivity-tools` curated collection — it surfaces repos to millions of visitors with zero ongoing effort. A coordinated launch day (social posts + community shares → 50-100 stars in 24h) is the ignition event everything else depends on.

---

## 1. GitHub Trending Algorithm

- **Primary signal:** star velocity — new stars gained in a rolling window (daily/weekly/monthly).
- **Secondary signals:** forks, commits, issues, PRs. Believed to be minor vs. stars.
- **Recency weighting:** recent activity matters far more than historical.
- **Rotation:** GitHub rotates repos off Trending to surface new projects, so a spike matters more than sustained trickle.
- **De-duplication:** same repo doesn't stay on Trending indefinitely.
- **Language filter matters:** C# Trending is less competitive than "all languages." Threshold: ~10-20 stars/day for daily C# Trending; ~40+ for top spots.
- **Sources:** [GitHub Blog](https://github.blog/news-insights/company-news/explore-what-is-trending-on-github/), [OSSInsight analysis](https://ossinsight.io/blog/introducing-trending-page), [Trending C#](https://github.com/trending/csharp)

### Implication for QuickSheet
A coordinated launch day pushing 50-100 stars in 24 hours would almost certainly land QuickSheet on C# Daily Trending. The social-post drafts already exist — the missing piece is the human pressing "publish" on HN/Reddit/Twitter simultaneously.

---

## 2. GitHub Explore Curated Collections (untapped)

GitHub's Explore page features hand-curated collections visible to millions. Anyone can submit a PR to [github/explore](https://github.com/github/explore) to add a repo.

### Collections where QuickSheet fits

| Collection | Why it fits | Current repos |
|---|---|---|
| **`productivity-tools`** | QuickSheet is a desktop productivity tool (notes, launcher, data tracking). Same category as Terminal, ripgrep, bat, zoxide, ShareX. | ~46 repos |
| **`text-editors`** | Stretch — QuickSheet is a spreadsheet/grid editor, not a text editor. But some entries are loose (Overleaf, Notepads). | ~41 repos |

**`productivity-tools` is the best fit.** The collection description is "Build software faster with fewer headaches, using these tools and tricks." QuickSheet's desktop scratchpad + launcher + data dashboard fits squarely.

### Submission process
1. Fork `github/explore`
2. Edit `collections/productivity-tools/index.md` — add `cemheren/QuickSheet` to `items` list
3. Open PR with description explaining what QuickSheet is and why it belongs
4. GitHub maintainers review and merge

### Implication for QuickSheet
**Draft a PR description for the user to submit.** This is a one-time action with permanent visibility payoff. No ongoing effort. Saved as `drafts/github-explore-productivity.md`.

---

## 3. Social Preview Image

- Recommended size: **1280×640px**
- Set via repo Settings → Social preview
- Appears on Twitter/X cards, LinkedIn, Slack, Discord, etc.
- QuickSheet currently has **no custom social preview** — GitHub auto-generates one from the README, which is generic.

### Implication for QuickSheet
A custom social preview showing the desktop wallpaper screenshot with the tagline "Your desktop is a spreadsheet" would dramatically improve click-through from every shared link. **Requires human to create the image** (skill has no display), but the spec is: 1280×640px PNG, dark background, QuickSheet grid with cells visible, tagline overlaid.

---

## 4. GitHub Topics (already optimized)

QuickSheet has 20 topics set — this is well-covered. Current topics include: cli, console, csv, desktop-wallpaper, dotnet, productivity, spreadsheet, terminal, tui, csharp, extensions, linux, windows, x11, csv-editor, terminal-spreadsheet, devops.

No action needed here.

---

## 5. Launch Day Playbook (synthesis)

To maximize the chance of hitting C# Trending:

1. **T-minus 1 day:** Set custom social preview image (human creates).
2. **T=0 (Tuesday or Wednesday, 9am EST):**
   - Post Show HN (draft exists in `drafts/`)
   - Post to r/commandline, r/dotnet, r/linux, r/programming (drafts exist)
   - Tweet thread (draft exists)
   - Post to Lobsters (draft exists)
3. **T+1 hour:** Submit PR to `github/explore` `productivity-tools` collection.
4. **T+1 day:** If HN hits front page, post dev.to article (draft exists).
5. **T+2 days:** Submit to awesome-lists (drafts exist for 5+ lists).

**Tuesday/Wednesday launch** is optimal — weekday developer traffic peaks, and the Trending window resets daily.

---

## Queued Actions

1. **Draft `github/explore` productivity-tools submission** → `drafts/github-explore-productivity.md` (this run)
2. **Social preview image spec** → needs human to create 1280×640px PNG (logged as follow-up)
3. **Launch day coordination checklist** → append to social drafts (future run)
