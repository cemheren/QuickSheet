# Persona 9 — Students (CS / STEM, undergrad + bootcamp + self-taught)

Status: **done** · Last revised: 2026-05-16

## TL;DR (5 lines)

- ~3M US CS/STEM undergrads + ~500k bootcamp learners + uncountable self-taught Discord-and-YouTube cohort. **Highest star-per-install ratio of any persona** — students star to bookmark, share aggressively in group chats, post their `~/.config` rices.
- Single-monitor laptop most of them; the wallpaper is the *only* dashboard surface. Free + zero-install + GitHub-shaped (clone-build-run) is exactly the install profile this audience reliably completes.
- Highest-leverage wallpaper cells: assignment-due countdowns, GPA / class-grade tracker, schedule strip ("Calc II in 47min"), LeetCode streak, GitHub contribution today, daily Pomodoro target, Discord/Slack unread count, job-application tracker.
- Where to seed: r/csMajors (~150k), r/learnprogramming (~4M), r/cscareerquestions (~1M), r/college, r/GetStudying, **r/unixporn (huge student crowd — rice culture)**, CS Discord servers (theprimeagen, Fireship), university subreddits, dev.to "student" tag.
- Direct competition: Notion (free for students! actually competitive), but lives in a browser tab. Apple/Google Calendar (no glance). My Study Life (student SaaS, abandoned). **No wallpaper-shaped student dashboard exists.**

## 1. Profile

- US: ~3M undergrad CS/STEM majors + ~500k bootcamp learners (BLS-adjacent; bootcamp number from Course Report). Globally many millions more.
- Hardware: single laptop (Mac dominant in CS undergrad; ThinkPad/Linux indie crowd large). Often a single external monitor in dorm. **Wallpaper is primary glance surface.**
- Day-shape: classes 2–6 hours, study 4–8 hours, side projects, gaming, Discord. Heavy multitasking.
- Buying brain: "Is it free? Is it on GitHub? Will my friends think it's cool?" Star ceremony: open the README, see a screenshot, star, share in Discord, install next week.
- Cultural moment: **post-LLM CS undergrad** — anxiety about job market + "everyone uses ChatGPT" lifestyle. Tools that prove you're not just AI-vibing are status symbols.

## 2. Why their desktop is wasted

What lives there now:

- Anime/Hyprland rice (Linux side) or Sequoia-default (Mac) or Win11-stock (Windows).
- Sticky note app with one TODO from 3 weeks ago.
- Notion in a perma-tab they avoid opening.
- Discord taking 800 MB RAM.

What QuickSheet replaces: the *idea* of a personal dashboard that students try (and abandon — Notion, Obsidian, Trello, paper planner) every semester. **The wallpaper version sticks because there's no app to open and no muscle memory to build.** Boot the laptop → see assignments due → study.

The rice angle: r/unixporn-style customisation is *the* way CS students express identity online. A wallpaper grid is a *new rice surface*, not a "productivity app."

## 3. Glanceable data they actually want behind windows

1. **Today's classes strip** — N cells, "Calc II 10:00", "CS101 14:00", current class highlighted.
2. **Assignment-due countdowns** — "CS final 3d," "ML hw 18h," "essay 4h" (red).
3. **GPA / per-class grade tracker** — column of classes with current grade, plus running GPA cell.
4. **GitHub contribution today** — green if pushed, red if not. Tiny streak counter.
5. **LeetCode / HackerRank streak** — "12-day streak" cell.
6. **Pomodoro target** — "3/4 sessions today" with progress bar.
7. **Job-application tracker** (seniors) — applied / interview / offer counts.
8. **A "park here" column** — random thoughts that interrupt studying go here. Same as writers' park column.

## 4. Candidate extensions / desktop-only features (ranked)

| Rank | Item                                       | Type | Cost | Hit prob | Why                                                                            |
|------|--------------------------------------------|------|------|----------|---------------------------------------------------------------------------------|
| 1    | `leetcode:` profile streak + solved count   | ext  | low  | high     | GraphQL endpoint public. CS-student-identity. Every CS Discord posts streaks.  |
| 2    | `gh:` GitHub today commits / streak         | ext  | low  | high     | Free API. "Contribution today" cell. Massively shareable.                       |
| 3    | `gpa:` GPA + class-grade calculator         | ext  | low  | med-high | Local CSV: class, credits, current grade. Computes GPA. Pure math.             |
| 4    | `schedule:` ICS / class schedule reader     | ext  | low  | high     | Reads an `.ics` (uni LMS exports them). Today's classes strip.                  |
| 5    | **In-cell progress bar prefix**            | feat | low  | high     | "3/4 pomos" → `▓▓▓░ 3/4`. Cross-persona feature now 5/8 personas want it.    |
| 6    | **Value-driven cell colour**                | feat | low  | high     | Red overdue, green done. 8/8 personas now.                                     |
| 7    | `wakatime:` coding-hours-today              | ext  | low  | med-high | WakaTime free tier. CS students using it = a flex.                              |
| 8    | `jobs:` LinkedIn/Indeed saved-search count  | ext  | high | med      | OAuth-heavy. Defer to senior-year persona slice.                                |
| 9    | `canvas:` Canvas LMS assignments            | ext  | med  | med      | Most US unis run Canvas; REST API exists with token. Hits exactly assignment-due cells. |
| 10   | `pomo:` already shipped — re-pitch          | docs | —    | —        | Pomo extension already exists. Re-frame for students with screenshot.           |

The rice culture means **screenshots are the entire product** for this persona. The `examples/student-dashboard.csv` is more important than any individual extension.

## 5. Where they hang out (rice + student channels)

- **r/unixporn** — *the* student rice community. Catppuccin/Gruvbox/Nord rices dominate. A QuickSheet-as-rice-element screenshot (with the freshly-shipped themes) belongs here.
- **r/csMajors (150k), r/cscareerquestions (1M), r/learnprogramming (4M), r/AskCS, r/computerscience**.
- **r/college (1.5M), r/GetStudying (350k), r/StudyTips, r/GetMotivated** — general study, not just CS.
- **r/leetcode (250k)** — focused; tools that surface streaks land.
- **University subreddits** — too many to list but each has 5–50k. r/uwaterloo, r/UCSD, r/UTAustin, etc.
- **Discords:** theprimeagen, Fireship Cult, ThePrimeagen, NeetCode, CS50, individual university servers.
- **Twitter/Bluesky:** student-tech Twitter (#100DaysOfCode), CS YouTubers' replies.
- **dev.to** — student tag, beginner-friendly publishing.
- **People to be visible to:** ThePrimeagen (Twitch streamer, signal-boosts hacker-aesthetic tools), Fireship (Jeff Delaney), Theo Browne, NeetCode (Navi). All have huge student followings.

## 6. Discoverability hooks

Hero image = **a student's Hyprland desktop**. VSCode open with a homework `.cpp` file. Wallpaper: QuickSheet grid with today's classes ("Calc II in 47min"), "ML hw due 18h" in red, "GitHub streak: 47 days," "LeetCode: 312 solved," GPA 3.72, Pomo `▓▓▓░ 3/4`. Catppuccin theme.

Headlines that land:

- "[r/unixporn] [Hyprland] My CS-student wallpaper has my schedule, LeetCode streak, GitHub today, and grade tracker"
- "I built a free, open-source student dashboard — it's your wallpaper"
- "Stop opening Notion. Your assignments are on your wallpaper now."
- "[r/csMajors] A free assignment + grade tracker that lives on your wallpaper (zero deps, MIT)"

Avoid: "productivity" jargon, anything that sounds like a Notion competitor, AI-feature claims.

## 7. Implications (queue these)

1. **Build `leetcode:` extension first** (Bucket F, post-research). Lowest cost, highest cultural-resonance for the persona. CS students *post* their LeetCode streaks; surfacing one on a wallpaper = instant share.
2. **Build `gh:` user-streak extension** second. Combines with `leetcode:` to form the "CS student flex bundle."
3. **Build `schedule:` ICS reader** third. Universities export `.ics` from LMS; one extension hits every student with a `.ics` link.
4. **Bucket A: `examples/student-dashboard.csv` + `docs/for-students.md`.** Pre-built with `gh:`, `leetcode:`, `schedule:`, `pomo:`, GPA cells, Pomodoro target. Screenshot = entire post.
5. **Bucket C: r/unixporn rice post** with student-themed dashboard. Likely the **single most-viral candidate** in the entire 10-persona slate (combines rice + student status + free open-source). Save in `drafts/students-unixporn.md`. Pair with theme choice — Catppuccin or Gruvbox is the right rice aesthetic for this audience.

Cross-link:
- [gamedev-ttrpg.md](gamedev-ttrpg.md) — same r/unixporn / rice culture.
- [writers-academics.md](writers-academics.md) — overlap on park-column + progress-bar.
- [developers-sre.md](developers-sre.md) — SRE-equivalent recurring features (colour + staleness) apply.
