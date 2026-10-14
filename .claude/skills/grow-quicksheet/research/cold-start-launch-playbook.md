# Cold Start Launch Playbook — From 0 to First 50 Stars

## TL;DR

QuickSheet is **over-produced, under-distributed**. The repo has 40+ extensions,
polished README, audience landing pages, CLI flags — but 0 stars. The bottleneck
is a single coordinated launch event, not more features. One well-timed HN post
with a demo GIF could deliver 50+ stars in 48 hours.

## Key Data Points

### Timing & Venue

- **Best HN window:** Tuesday 6:30–9:30am EST (least competition + high traffic).
  Monday is viable but ultra-competitive. Wednesday slightly lower energy.
- **Title pattern that works for TUIs:** "Show HN: [Tool] – [one-line pitch]"
  — concrete, no buzzwords, hint at novelty.
  - Example: `Show HN: QuickSheet – Your Linux/Windows desktop wallpaper is a spreadsheet`
- **Conversion rate:** 1–10% of HN page visitors star if README/demo is compelling.
  A front-page "Show HN" gets ~2,000–5,000 visitors → 20–500 stars potential.
- **First 48 hours:** 60–70% of stars from a launch post arrive in this window.

### What correlates with high scores on HN for TUI/desktop tools

1. **Animated demo at the top** — asciinema or GIF showing real usage.
2. **"Why this exists" narrative** — personal pain point, not feature list.
3. **Zero-dep / small-binary** angle — HN loves lightweight, auditable tools.
4. **Active founder engagement** — responding to every early comment within
   minutes boosts rank algorithm.
5. **Cross-platform** — "Works on Linux AND Windows" broadens appeal.

### Adjacent success stories

- **ycombo** (X11 HN widget) — exact same "desktop widget" space. Niche but
  got attention from r/unixporn and HN crowd.
- **edex-ui** — 44.9k stars. Sci-fi terminal dashboard. Proves "make your
  desktop do something" resonates massively if visuals are striking.
- **Preevy** — 1,500 stars in 12 weeks via coordinated launch: GIF demo,
  blog post, newsletter pitches, Discord sharing, "good first issue" labels.

### Distribution channels ranked by expected yield

| Channel | Expected stars | Effort | Notes |
|---------|---------------|--------|-------|
| Show HN (front page) | 50–500 | Medium | Needs GIF + active commenting |
| r/unixporn (rice post) | 20–100 | Low | Needs screenshot of polished setup |
| r/commandline | 10–50 | Low | Short demo, link to repo |
| r/linux | 10–30 | Low | Focus on X11 wallpaper angle |
| awesome-list PRs | 5–20 each | Low | Drafts ready, just submit |
| Console.dev / TerminalTrove | 5–15 | Low | Drafts ready |
| Dev.to blog post | 5–20 | Medium | Draft ready |
| Twitter/X thread | 5–50 | Medium | Needs visual assets |
| goodfirstissue.dev listing | 2–10 | Trivial | Add "good first issue" labels |
| Product Hunt | 10–50 | Medium | Good for dev tools |

### What's MISSING from QuickSheet right now for launch readiness

1. **Demo GIF/video** — THE single blocker. No amount of text sells a visual
   tool without a visual. This needs the human to capture.
2. **"Good first issue" labels** — Zero exist. Adding 3–5 would get the repo
   indexed by contributor-matching sites.
3. **GitHub Release** — No tagged release yet. PR #308 adds the workflow but
   isn't merged. A `v1.0.0` release signals maturity.
4. **Star-history badge** — Embed `star-history.com` graph in README. Creates
   FOMO once the first spike hits.

## Implications for QuickSheet — 3 Specific Actions to Queue

1. **Create 3–5 "good first issue" labels on real issues/feature requests.**
   This gets the repo indexed by goodfirstissue.dev, CodeTriage, Up For Grabs,
   and GitHub's own Explore feed. Zero cost, passive discovery. The skill CAN
   do this autonomously (just `gh issue create` with the label).

2. **Draft a coordinated launch-day checklist for the human.** A single doc
   that says: "On launch day, do these 7 things in order within 2 hours."
   Posting order: HN → Reddit (3 subs) → Twitter → awesome-list PRs.
   All drafts already exist. The human just needs to execute the sequence
   on a Tuesday morning EST. Save as `drafts/launch-day-checklist.md`.

3. **Add a "call to action" in README** — a subtle "⭐ Star this repo if you
   find it useful" near the Quick Start section. Many successful projects do
   this (Preevy, Hoppscotch). It converts passive visitors into stars. The
   skill can PR this.

## Anti-patterns to avoid

- **Drip-feeding features without launch** — More PRs won't help at 0 stars.
  The repo needs ONE coordinated moment.
- **Posting without visuals** — Text-only posts about a visual tool get
  scrolled past. The GIF is non-negotiable.
- **Spreading thin** — Better to nail HN + 2 subreddits on the same day than
  to post to 10 places over 10 days. Momentum compounds within a 48-hour window.
- **Over-engineering before launch** — The repo is already feature-complete
  enough. Ship ≠ more features; ship = tell people it exists.
