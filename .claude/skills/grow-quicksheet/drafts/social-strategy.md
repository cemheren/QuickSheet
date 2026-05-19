# QuickSheet social strategy — execution playbook

Written 2026-05-18 after stars went 1→0. **You execute, skill drafts.** Every venue's copy already exists under `drafts/`; this file is the *order*, the *gating*, and the *don't-fuck-it-up* rules.

---

## The honest situation

- Stars: 0. One person starred, then unstarred.
- Distribution attempted: 0 venues.
- Cause of low stars: **nobody has seen it yet.** Not a product problem.

The repo is feature-complete enough to ship. Drafts are written and ranked. The only blocker is *you posting them*.

## North star

> Get the wallpaper screenshot in front of 50,000 people in the genres that actually like wallpaper tools, in the right order, with one shot per venue. Measure stars at +24h, +72h, +7d.

Not "go viral on HN." HN is the *wrong* primary venue for the wallpaper framing — `research/wallpaper-launch-venues.md` has the data (Übersicht: 5 HN attempts, never >10 pts; Parallax: 225 pts only by leading with technique, not product).

## Asset gate — DO NOT SKIP

Skill has no display, cannot make these. **You** produce them once, reuse everywhere.

| Asset | Why it matters | Spec |
|------|----------------|------|
| **A1. Wallpaper screenshot** (1080p+) | r/unixporn, r/Rainmeter, r/selfhosted, Twitter, Mastodon all live or die on this single frame. | Real desktop. Real data. Show: sparklines, one extension (weather/tls), one `r:` launcher row, one `i:` live cell. Crop to 16:9. PNG. |
| **A2. 10-second GIF / mp4** | Show typing into a cell → live output appearing. Differentiates from screenshot-only competitors. | Use `peek` (Linux) or ScreenToGif (Windows). ≤4 MB. Hosted on the repo (`docs/img/`) so embed works on HN. |
| **A3. README hero image** | First-impression upgrade. Repo currently leans on tour.md. | Static frame from A1, embed at top of README. PR-gate this. |

**No A1 → no post.** Posting without a screenshot is how 90% of these communities silently downrank you. Capture A1 first, do *literally nothing else* socially until it exists.

---

## Launch sequence — one venue per day, in order

Sequencing rule: **start where the wallpaper framing is rewarded; only widen after one post has traction.** Each step has a gate to the next.

### Day 0 — Capture assets

- Take screenshot A1.
- Save into `docs/img/wallpaper-hero.png`.
- (Optional) Record GIF A2.
- Commit + push (PR'd separately by the skill).

### Day 1 — **r/unixporn** (primary)

- **Draft:** `drafts/unixporn-rice.md`
- **Why first:** Per `research/wallpaper-launch-venues.md` — wallpaper-tool genre's actual cultural home. Screenshot does 80% of the work.
- **Post format:** Image post. Title in r/unixporn `[OC] / [<DE>] / <vibe>` convention. Body: short, no project pitch, link in comments.
- **Posting window:** Sun 09:00–11:00 ET or Tue 19:00–21:00 ET (rice subs skew evening + weekends).
- **Gate to Day 2:** ≥50 upvotes in 12h → proceed. <50 → pause, A1 needs work.

### Day 2 — **r/Rainmeter** + **r/desktops** (peer subs)

- **Draft:** reuse `drafts/unixporn-rice.md`, retitle for each. Body emphasizes "Rainmeter alternative that's also interactive."
- **Gate to Day 3:** any positive engagement → proceed. Hostile → pause 48h.

### Day 3 — **r/selfhosted**

- **Draft:** custom (skill will write next run). Frame: "Homepage.io / dashy alternative that doesn't live in a browser tab."
- **Why here, not r/homelab:** r/selfhosted rewards alternatives-to-popular-tools posts; r/homelab is harder without rack porn.
- **Asset:** A1 + one paragraph on the extension network (40+ extensions, JSON-lines protocol).

### Day 4 — **HN, reframed**

- **Draft:** `drafts/showhn.md` — already revised for the technique-first angle.
- **Critical:** Title leads with mechanism, NOT "wallpaper spreadsheet." Use the technique-first title option in the draft.
- **Best window:** Tue–Thu 08:00–10:00 ET.
- **First comment:** Post your own clarification comment within 60s of submission, explaining the "why it exists" — increases comment count, helps thread surface.
- **Gate:** if >50 pts in first hour, switch to engagement mode (reply fast to every top-thread comment under 30 min). If <10 pts at 2h, walk away, don't repost.

### Day 5 — **r/commandline** + **r/dotnet**

- **Drafts:** `drafts/reddit-commandline.md`, `drafts/reddit-dotnet.md`.
- These two are *different* posts with different angles. Don't cross-link.
- **r/commandline:** lead with TUI mode demo.
- **r/dotnet:** lead with zero-NuGet + P/Invoke story.

### Day 6 — **Lobsters** + **Mastodon/Bluesky**

- **Drafts:** `drafts/lobsters.md`, `drafts/mastodon-bluesky.md`.
- Lobsters is small but high-signal. Mastodon hits FOSS devs (Fosstodon/Hachyderm).

### Day 7 — Cool-off + assess

- No posting Day 7. Read every comment, fix every legitimate criticism. Star delta should be visible by now.
- If trajectory good: queue Day 8+ longer-form (`drafts/devto-blog.md`).
- If flat: stop launching, do one *very specific* improvement based on the most common criticism, then plan a re-launch in 30 days.

### Day 8+ — Slow burn

- **dev.to / Medium post** — long-form code walkthrough (draft already exists).
- **Newsletter pitches** — terminaltrove + console.dev (`drafts/email-pitches.md`). Send-and-forget.
- **Awesome-list submissions** — drafts under `drafts/awesome-*.md`. PRs to those repos *you* file, not the skill.

---

## What NOT to do — common failure modes

1. **Don't post the same post to 3 subs the same day.** Reddit catches cross-posting and downranks. One venue per day, customized copy each time.
2. **Don't post without A1.** A wall of text in r/unixporn gets 0 upvotes.
3. **Don't reply defensively to criticism.** Top comment on HN is often a "this is just X with Y" comparison — engage on the technique, acknowledge the comparison, don't litigate.
4. **Don't claim "production-ready" / "thousands of users" / "fastest."** Truthful only. The repo is honest about being a side-project — keep it.
5. **Don't pay for promotion, buy upvotes, or use alts.** Will get banned from the venues that matter (r/unixporn especially aggressive about this).
6. **Don't post during low-traffic windows.** Times in this doc are calibrated; off them, engagement drops 5x.
7. **Don't re-post after a flop.** Mods notice. A flop today is recoverable in 30 days with new framing; a re-post within a week is permanent damage.

## Engagement rules during a live post

- **First 60 min are decisive.** Block the calendar.
- Reply to every top-level comment ≤30 min after it lands.
- For comparison comments ("this is just lazygit/visidata/Rainmeter"): acknowledge → name the *one* differentiator → ask a question to keep the thread alive. Example: *"Fair comp. The piece Rainmeter doesn't do is the interactive editing — you can type into the wallpaper and Excel can open the CSV. What would you want from a Rainmeter-replacement that's actually a spreadsheet?"*
- Star count is the lagging indicator. The leading indicator is **comments-per-upvote**. Healthy = >1 comment per 5 upvotes. Below that → post is a flat curiosity, won't convert.

## Measurement

Track in `.claude/skills/grow-quicksheet/log.md` after each post:

```
## YYYY-MM-DD <venue>
- Asset: A1 / A2 / both
- Title used: <copy>
- Submitted: <time, day-of-week>
- T+1h: <pts/comments>
- T+12h: <pts/comments>
- T+24h: <pts/comments/⭐ delta>
- Top criticism: <one line>
- Lesson: <one line>
```

This is the dataset that informs round 2.

## Hard rules (project-wide)

- Truthful claims only.
- No paid promotion / bots / alts / astroturf.
- One venue per day max.
- One product story, customized per venue — not 8 different products.
- If a venue's rules say "no self-promo," follow them — link in comment, not in title.

---

## What the skill will keep doing in parallel

Drafts, extension scaffolds, doc polish — produced and PR'd on the existing cadence. Skill won't post anywhere. Skill won't pressure you to. **Distribution is your job; production is the skill's job.** This doc keeps that line clear.

Next-run skill follow-ups, queued:
- Draft the r/selfhosted "Homepage.io alternative" post once A1 lands.
- Add `docs/img/` to README's hero once you commit A1.
- Build the r/Rainmeter variant of the unixporn draft (different audience tone).
