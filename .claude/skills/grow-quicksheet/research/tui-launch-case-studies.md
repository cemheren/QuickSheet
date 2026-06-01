# TUI Launch Case Studies: 0 → 1000+ Stars

Research run 2026-05-31. Goal: distill what specific projects did to cross star
milestones, with numbers, so QuickSheet can replicate.

## Summary (5 lines)

- A single viral HN post yields **121 stars/24h, 189/48h, 289/week** on average (arXiv study of 138 repos).
- Optimal post time: **Tue/Wed 14:00–15:00 UTC** or **Sun 0:00–2:00 UTC** (highest breakout rates at 12–15.7%).
- "Show HN" tag gives NO statistical edge after controlling for timing — it's the hour, not the prefix.
- Lazygit's key hack: **in-app "star this repo" popup on first launch** — directly converts users to stargazers.
- From 0 to initial traction: **seed 30–100 stars from personal network** before the public launch to trigger GitHub trending algorithms.

## Case Study 1: lazygit (78k+ stars)

**Launch**: Aug 2018. "Show HN: Lazygit – simple terminal UI for git commands"
- Initial HN post: ~192 points, 77 comments.
- Growth hack (per Jesse Duffield's "Lazygit Turns 5" retrospective):
  - On first launch, displays a popup encouraging users to star the repo.
  - Keeps everything (docs, info) inside the GitHub repo — no separate site.
  - Users are always "one click away from starring."
- Lesson: **The repo IS the landing page.** Direct engagement where users already are.
- Luck acknowledged: the HN post gave critical early visibility. Without it, would have stalled.

Source: https://jesseduffield.com/Lazygit-5-Years-On/

## Case Study 2: btop (resource monitor)

**Launch**: "Btop: A better modern alternative of htop with a gamified interface"
- HN post: 190 points, 118 comments.
- First comment was about title editorialization (author apologized) — yet still performed.
- Key factor: **stunning visual demo** in README. The screenshots sell the tool instantly.
- Lesson: **Looks matter more than perfect HN etiquette.** If it's beautiful, people star it.

## Case Study 3: Preevy (0 → 1,500 stars in 12 weeks)

Not a TUI, but the campaign plan is gold:
- 3 launch blog posts targeting different dev audiences.
- 2 Product Hunt launches (alpha + key version).
- 5–7 relevant subreddits (r/selfhosted, r/coolgithubprojects, etc.).
- Influencer engagement on Twitter for retweets.
- Weekly spikes of 100–200 stars after each push.

Source: https://livecycle.io/blogs/opensource-github-stars/

## Case Study 4: Open Interface (cross-platform LLM desktop app)

- A single "mediocre" r/OpenAI Reddit post unexpectedly hit **143K views + 500 shares**.
- Drove months of SEO traffic and sustained star growth.
- Lesson: **Reddit can outperform HN** if you hit the right niche sub at the right moment.

Source: https://www.indiehackers.com/post/0-to-1000-github-stars-for-your-open-source-projects-db2efb62f1

## Case Study 5: cell (terminal spreadsheet — direct competitor)

- garritfra/cell — "A fast terminal spreadsheet editor with Vim keybindings" (Rust).
- ~298 stars. Posted to HN (item 40060335) but didn't break out virally.
- Missing ingredient: **no animated demo GIF in the HN thread or README hero.**
- Lesson: QuickSheet's closest competitor couldn't cross 500 even with HN exposure.
  This means the **category is underexploited** — a better launch could own it.

## Quantitative Data (arXiv: Launch-Day Diffusion)

Paper: Obada Kraishan, arXiv:2511.04453, 2025. Analyzed 138 AI/tool repo launches.

| Metric | Value |
|--------|-------|
| Avg stars gained in 24h post-HN | 121 |
| Avg stars gained in 48h | 189 |
| Avg stars gained in 7 days | 289 |
| "Show HN" tag effect (after controls) | Not significant |
| Optimal posting window | 12–17 UTC |
| Best single hour | 12:00 UTC (12.2% breakout rate) |
| Sunday 0–2 UTC breakout rate | 15.7% |

## Optimal Posting Times (multiple sources)

| Day | Best Time (UTC) | Notes |
|-----|----------------|-------|
| Tuesday | 14:00–15:00 | Historically highest avg scores |
| Wednesday | 14:00–15:00 | Nearly as strong |
| Sunday | 0:00–2:00, 11:00–16:00 | Highest "Show HN" breakout rate |
| Weekdays general | 11:00–13:00 | Solid ~10–11% rate |
| Avoid | Fri after 14:00 UTC, 3–7 UTC any day | Troughs |

Sources: https://www.myriade.ai/blogs/when-is-it-the-best-time-to-post-on-show-hn,
https://github.com/JoseMarquezAlberti/hn-best-posting-times

## The Lazygit "First-Launch Star Nudge" Pattern

This is the single most actionable tactic from this research:
1. On first launch, show a brief, friendly message: "If QuickSheet is useful, ⭐ it on GitHub: <url>"
2. Only show once (flag in config/autosave).
3. Non-intrusive — a one-liner in the status bar or a 3-second toast, not a modal.

Why it works:
- Converts users who already like the tool (they installed it) into visible supporters.
- High star count triggers a virtuous cycle: trending → more installs → more stars.
- Jesse Duffield credits this as a major factor in lazygit's growth.

## Implications for QuickSheet (actionable queue items)

1. **Implement a first-launch star nudge** (Bucket E). On first run, if no
   autosave exists, display a one-time message in the grid or status bar:
   "⭐ If QuickSheet helps, star it: https://github.com/cemheren/QuickSheet".
   Store a `star_nudge_shown=true` flag. ~20 lines. High ROI per line of code.

2. **Time the HN launch precisely.** Target **Tuesday or Wednesday 14:00 UTC**
   (7am PT, 10am ET, 4pm CET). Or if feeling bold, **Sunday 12:00 UTC**. Have
   the Show HN draft, first comment, and README all polished before the clock
   hits. The existing `drafts/showhn.md` is ready — just needs final timing.

3. **Seed 30–100 stars before going public.** Ask friends, colleagues, community
   members who've already seen the project to star it. This primes the GitHub
   trending algorithm and adds social proof before the HN crowd arrives.

4. **Animated GIF is non-negotiable for HN.** The `cell` spreadsheet failed to
   break out partly because it lacked a visual hook. QuickSheet's wallpaper mode
   is its differentiator — a 10-second GIF of a spreadsheet AS the desktop
   wallpaper with live `i: htop` output would be jaw-dropping on HN. This
   requires a human with a running desktop, but it's the #1 blocker for launch.

5. **Reddit backup plan.** If HN underperforms, immediately cross-post to
   r/commandline and r/linux (drafts ready). Reddit can deliver 143K views from
   a single post if it catches. Different audience, different hour — no need to
   wait.
