# Persona 4 — Visual designers, illustrators, 2D artists

Status: **done** · Last revised: 2026-05-16

## TL;DR (5 lines)

- Freelance illustrators, indie game artists, motion designers, comic artists. US/EN headcount ~250–400k freelancers + many more hobbyists. **Multi-monitor culture; second monitor often reference-only.**
- They are NOT terminal users. The wallpaper sell is "your reference grid, color palettes, commission queue, and time-spent-per-piece are pinned behind Photoshop/Procreate/Clip Studio — visible the moment you alt-tab."
- Highest-leverage wallpaper cells: per-commission time + price + status, hex-color swatch row, reference-image links, deadline countdowns, daily revenue running total.
- Where to seed: r/ArtistLounge, r/comics, r/IndieDev (art-side), Cara.app (post-Adobe migration), Bluesky #ArtCommunity, ArtStation forums. *Plus* r/desktops with the "artist's desktop" angle.
- Competition for the wallpaper slot: nothing native. Trello/Notion live in tabs. Color-palette tools (Coolors, Adobe Color) live in browser. **Nobody sells an "artist's ambient layer."**

## 1. Profile

- Subsegments: illustrators (children's books, editorial, gallery), 2D game artists, motion/After Effects designers, comic creators, graphic designers, character designers.
- Income shape: per-piece commission, hourly retainer, or per-spot rate. *Always* tracking time-per-piece vs price-charged to know if they're earning min wage.
- Hardware: high-spec desktop or Wacom Cintiq + secondary reference monitor; or laptop + iPad. **The secondary screen is wallpaper-prime territory** — already used for reference art today.
- Buying brain: "Does it interrupt my flow? Does it work offline? Does it cost a monthly fee?" Allergic to subscriptions after Adobe.
- Cultural moment: post-Adobe-CC-creep + post-Procreate-AI-controversy; community migrating to Cara, Bluesky, indie tools. **Receptivity to indie/open-source is unusually high in 2026.**

## 2. Why their desktop is wasted

What lives there now:

- Reference image dumped in `~/Desktop/refs/` with cryptic filenames.
- A static personal-portfolio render.
- Maybe one PureRef window if they know about it (closest competitor — but PureRef is reference-image-only, no data).
- A messy folder of WIP files.

What QuickSheet replaces: the *intent* of opening Trello/Notion to "see what commissions I owe." With the wallpaper, that view is on-screen the moment Photoshop loses focus. **Reference images + commission status + time + colors on one surface.**

Note: QuickSheet is not a reference-image board (PureRef wins on visual reference). It is the *data shell* around the work — who's paying, when's it due, what colors, how many hours.

## 3. Glanceable data they actually want behind windows

1. **Commission queue table** — N rows: client, title, price quoted, due date, status. Status colour-coded (red = overdue, yellow = revision, green = ready to deliver).
2. **Today's hours per piece** — per-commission timer cells. Click to start/stop. Critical for hourly retainer work.
3. **Effective $/hour cell** — single big cell = SUM(prices delivered this month) / SUM(hours tracked). The "am I undercharging?" cell.
4. **Color palette row** — `c:hex:#FF5A4D` style cells, the active project's palette. One glance, no more "what was that orange again."
5. **Reference link row** — clickable URLs to Pinterest boards / ArtStation refs / mood-board Notion pages.
6. **Daily revenue running total** — single cell, today's earned vs target.
7. **Streak / hours-this-week cell** — ambient motivation.
8. **Inbox: pending DMs** — count of "are you taking commissions?" messages.

## 4. Candidate extensions / desktop-only features (ranked)

| Rank | Item                                       | Type | Cost | Hit prob | Why                                                                       |
|------|--------------------------------------------|------|------|----------|----------------------------------------------------------------------------|
| 1    | **Hex-color cell prefix** (`c:#FF5A4D:`)   | feat | low  | high     | `c:color:` already takes named colours; extend to `#RRGGBB`. Palette row.  |
| 2    | **Per-cell ticking timer** (shared w/ lawyers/accountants) | feat | low | high | Foundational; unlocks per-commission tracking.                              |
| 3    | **Image-thumbnail cell** (`img: path|url`) | feat | high | high     | Render a small inline thumbnail. Would make artist personas viable.        |
| 4    | **Value-driven cell colour**                | feat | low  | high     | Overdue commissions red, paid green. Trinity unlock again.                 |
| 5    | `etsy:` open-order count                    | ext  | med  | high     | Etsy API; sticker/print sellers live in it.                                |
| 6    | `gumroad:` revenue today                    | ext  | low  | med-high | Free API; print/asset sellers.                                              |
| 7    | `instagram:` / `bsky:` mention count        | ext  | high | med      | OAuth-heavy; defer until artists actually ask.                              |
| 8    | `palette:` extract dominant colours from image| ext | med | med      | Open image → fill row with hex cells. Stunt-feature for posts.            |
| 9    | `commish:` template — pre-built CSV layout  | ext  | low  | med      | Just a curated starter `commissions.csv` + README. Frictionless onboard.   |
| 10   | `pomo:` already shipped — promote it        | docs | —    | —        | Artists love Pomodoro; re-pitch with art-friendly screenshot.               |

Feature 3 (image thumbnails) is the **single biggest unlock** for this persona. Without it the wallpaper is text-only and PureRef wins. Building inline thumbnails is non-trivial cross-platform — needs investigation. Could be desktop-only (skip TUI), since wallpaper is the only place it matters.

## 5. Where they hang out (desktop-relevant first)

- **r/desktops, r/macsetups** — "my artist desktop" is a popular post format here.
- **r/ArtistLounge** (~400k) — process + tooling discussions. Less hostile to indie tools than r/learnart.
- **r/comics, r/webcomics, r/IndieDev (art side), r/Illustration** — profession-specific.
- **Cara.app** — Adobe-disenchanted artists migrated here in 2024–25. High receptivity to indie/open-source.
- **Bluesky** — #ArtCommunity / #CommissionsOpen hashtags. Higher signal than Twitter in 2026.
- **ArtStation forums + ArtStation Magazine** — pro side; review pieces are linked widely.
- **Discords:** Crit-the-Boards, Sketch Daily, individual artist Patreon Discords.
- **People to be visible to:** Aaron Blaise (free education), Marc Brunet (cubebrush.co), David Revoy (Krita evangelist, open-source-friendly), Ethan Becker, Ross Tran (Ross Draws). Especially **David Revoy** — open-source advocate, blogs about Linux art workflows; would notice an indie tool natively.

## 6. Discoverability hooks

Hero image = **Photoshop or Procreate open with an active piece; QuickSheet wallpaper visible around/behind it: commission queue with one red overdue cell, hex-color palette row underneath the active piece, "$/hour this month" cell glanceable.**

Headlines that land:

- "I built a free desktop dashboard for freelance artists (commissions, time, colors, reference)"
- "My wallpaper now shows my commission queue and effective $/hour"
- "Open-source, offline, your data stays on your machine — for artists post-Adobe"

The post-Adobe + post-AI-controversy moment is the lead, not the protocol. **Frame as "you own your tools, you own your data."**

Avoid: terminal screenshots, dev jargon, comparing to Trello (Notion-comparison is fine).

## 7. Implications (queue these)

1. **Bucket E: extend `c:color:` to accept `#RRGGBB` hex.** ~20 LOC. Without this, the palette row doesn't work.
2. **Bucket R sub-investigation: feasibility of inline image thumbnails in `--desktop` mode.** Windows (WinForms `Graphics.DrawImage` — easy). Linux X11 (`XPutImage` / `Xrender` — non-trivial). Decide build vs skip. Save findings as `research/inline-image-thumbnails.md`. Defer until investigation done.
3. **Bucket F: ship `gumroad:` extension.** Lowest-hanging fruit on the ranked list — free API, no OAuth, instant "wait this is my revenue on my wallpaper?" reaction.
4. **Bucket A: `commish:` starter CSV** under `examples/commissions-starter.csv` + a `docs/for-artists.md` landing page. Zero-LOC effective onboarding for this persona.
5. **Bucket C draft: David Revoy outreach + Cara.app post + r/desktops "artist desktop" screenshot post.** Save in `drafts/artists-launch.md` only after hex-colour + gumroad ship.

Cross-link:
- [lawyers.md](lawyers.md), [accountants.md](accountants.md) — shared timer + value-colour features.
- Future `research/inline-image-thumbnails.md` — feasibility brief.
