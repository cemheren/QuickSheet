# Persona 2 — Solo + small-firm lawyers (single-monitor desktop)

Status: **done** (v2, desktop-mode focus) · Last revised: 2026-05-16

## TL;DR (5 lines)

- ~400k solo US attorneys + 2–10-attorney firms; mostly **single-monitor laptops**. The desktop wallpaper is the only surface that's *always visible* between document tasks.
- They are not TUI users. The wallpaper sell is "open the laptop → see today's billable hours, deadlines, and unbilled matters before opening Word."
- Highest-leverage wallpaper cells: 6-min billing timer running per matter, deadline countdowns (FRCP-day math), conflict-check search, today's unbilled total.
- Where to seed: r/Lawyertalk, r/solopractice, Lawyerist Lab Slack, Lawyerist Podcast. *Plus* r/macsetups / r/desktops for the "lawyer desktop rice" angle — surprisingly engaged.
- Direct competition for wallpaper slot: Stardock Fences (Windows), GeekTool (macOS), nothing native. Practice-management SaaS (Clio) does NOT touch the desktop.

## 1. Profile

- ~1.3M licensed US attorneys; ~50% in firms <10 attorneys. ABA Profile 2024.
- Solo = ~400k. Their tooling decisions are personal, not committee-driven.
- Hardware: laptop with one external monitor at most. Often *no* external monitor in court / coffee shop. **Wallpaper is high-value real estate because it is reliable real estate.**
- Buying brain: "Where does my data live, and will this still work if the SaaS company dies?" CSV-first answers this.

## 2. Why their desktop is wasted

What lives behind their windows now:

- Static law-firm wallpaper or vacation photo.
- A messy folder grid of "Smith v Jones drafts," "tax stuff 2024," "intake forms."
- Maybe Clio in a browser tab they have to click into.

What QuickSheet replaces: the *intent* of opening Clio every morning to "see where I am." With the wallpaper, that view is already on-screen the moment the laptop wakes — no login, no tab. **And the data is the user's own CSV, not a SaaS captive.**

## 3. Glanceable data they actually want behind windows

1. **Today's billable hours so far** — single big cell that sums all matter timers. The day's score.
2. **Per-matter timer row** — N cells, one per active matter, each `bill: smith-v-jones` ticking up in 6-min increments. Click to stop/start.
3. **Deadline strip** — N cells, "3d to motion response," "12d to discovery cutoff," red if <24h.
4. **Unbilled matters list** — column of matter names with last-billed date; old ones go red.
5. **Conflict-check search** — top cell is a search box that filters a `clients.csv` next to the file. Live.
6. **Court calendar countdown** — next hearing in days.
7. **A scratch column** for "remember to follow up with X" — autosaved every 5s.

All visible without opening anything. The whole UX is "wake the laptop, glance, get back to drafting."

## 4. Candidate extensions / desktop-only features (ranked)

| Rank | Item                                          | Type | Cost | Hit prob | Why                                                              |
|------|-----------------------------------------------|------|------|----------|-------------------------------------------------------------------|
| 1    | **Per-cell ticking timer** (`bill:` semantic) | feat | low  | high     | The single highest-leverage desktop primitive for this persona.   |
| 2    | `bill:` extension on top of timer feature     | ext  | low  | high     | Wraps timer with matter ID + .1h rounding. CSV-exportable.        |
| 3    | `dock:` court deadline (FRCP-day math)        | ext  | med  | high     | Skip-weekends + federal holidays. Red countdown on wallpaper.     |
| 4    | **Cell value-driven colour** (`<24h = red`)    | feat | low  | high     | Already needed for SREs; equally critical for lawyers.            |
| 5    | `conflict:` local CSV fuzzy-search             | ext  | low  | high     | Pure-local, no network. Trust-builder.                            |
| 6    | `cite:` Bluebook formatter (rename current → `doi:`) | ext | low | med-high | Solo lawyers cite constantly; useful daily.                       |
| 7    | **Always-on-top "pinned cell"** mode           | feat | med  | med      | One cell stays visible even on focused windows. Deadline strip.   |
| 8    | `stat:` USC / state statute lookup            | ext  | med  | med      | Cornell LII scrape. Useful but lower frequency than billing.      |
| 9    | `caselaw:` Caselaw Access Project              | ext  | low  | med      | Free; case-name + snippet lookup.                                 |
| 10   | `pacer:` docket pull                           | ext  | high | med      | Paid + auth. Defer until users explicitly ask.                    |

The timer feature (#1) is the unlock — without it, none of the billing cells work as a wallpaper.

## 5. Where they hang out (desktop-relevant first)

- **r/macsetups / r/desktops** — surprisingly engaged with professional desktop setups; a "lawyer desktop with live billing and deadlines on wallpaper" post is novel here.
- **r/Lawyertalk** (~120k private), **r/solopractice**, **r/LawFirm** — profession side. Post after the desktop post earns the screenshot.
- **Lawyerist Lab Slack** — small but high-signal solo-lawyer community. Bob Ambrogi's LawSites blog reviews tools genuinely.
- **Podcasts:** Lawyerist Podcast, Above the Law. Pitch "your wallpaper as your billing assistant" — concrete enough to land as a segment.
- **Conferences:** ABA TECHSHOW (Chicago, spring); the legal-tech press attends.
- **People to be visible to:** Bob Ambrogi (LawSites), Carolyn Elefant (MyShingle), Sam Glover (Lawyerist), Nicole Black.

## 6. Discoverability hooks

Hero image = **a MacBook desktop, Word open with a brief draft, QuickSheet wallpaper visible to the right of the document showing a ticking timer for "smith-v-jones," today's total billable hours, and a red 18-hour countdown to "motion response due."**

Headlines that land:

- "My laptop wallpaper now tracks my billable hours and court deadlines (open source, CSV)"
- "Built a free desktop billing capture for solos — no SaaS, your data stays local"
- "A wallpaper that reminds you that the motion is due tomorrow"

Avoid: terminal screenshots, "TUI" anywhere in the copy, dev-jargon (`stdin`, `JSON-lines`, etc. — those go in a separate "how it works" section).

## 7. Implications (queue these)

1. **Bucket E: ship per-cell ticking timer prefix.** Foundational for billing/Pomodoro/anything time-shaped. <80 LOC additive. Without this, lawyer extensions are not viable.
2. **Bucket E: value-driven cell colour rules.** Required so deadlines turn red automatically — non-negotiable for the persona.
3. **Build `bill:` extension** after timer + colour land. Pair with already-existing `--export-md` for end-of-day billing summaries.
4. **Bucket C drafts:** `drafts/lawyers-launch.md` — r/macsetups post + LawSites email pitch + Lawyerist Podcast outreach. Lead with hero image above. *Save only after timer feature merged.*

Cross-link: [accountants.md](accountants.md) — same audit/CSV-trust value-prop; the timer feature serves both.
