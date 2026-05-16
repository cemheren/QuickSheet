# Persona 6 — Writers, researchers, academics

Status: **done** · Last revised: 2026-05-16

## TL;DR (5 lines)

- Three overlapping subsegments: long-form non-fiction writers (journalists, authors, Substack), fiction writers (novelists, screenwriters), and academics (PhD students, profs, post-docs). US/EN active: ~200k working journalists/authors + ~1.5M actively-publishing academics + ~5M Substack writers.
- Allergic to context-switching and SaaS. The desktop wallpaper is the calmest surface they own — perfect for ambient "where am I in the project?" cells.
- Highest-leverage wallpaper cells: today's word count vs target, manuscript section list with progress bars, citation queue (DOIs to format), open-loop research questions, distraction-free pomo timer, deadline countdowns.
- Where to seed: r/writing, r/academia, r/PhD, r/AskAcademia, r/Scrivener, r/ObsidianMD, r/zettelkasten, r/writingadvice, Mastodon scholar.social. *Plus* r/desktops with the "writer's desk" angle.
- Direct competition for wallpaper slot: nothing. Scrivener/Obsidian/Notion occupy a window. **The wallpaper is the orientation surface above all that.**

## 1. Profile

### 1a. Long-form non-fiction + Substack

- Mostly solo. Many publish on Substack (~5M total writers; ~80k earn revenue).
- Hardware: laptop, often single-monitor. Coffeeshop-frequent.
- Income shape: drip subscriptions, advances, occasional contract pieces.

### 1b. Fiction / novelists / screenwriters

- NaNoWriMo crowd, hybrid-published authors, indie screenwriters. ~400k self-pub authors on Amazon KDP in the US.
- Hardware: similar laptop-bound.
- Workflow-fetishists by culture — *will* install a third tool for a 5% productivity nudge.

### 1c. Academics

- Grad students + faculty across STEM and humanities. ~1.5M in US higher-ed.
- Hardware: laptop + library desktop + sometimes campus dual-monitor.
- Cycle: lit review → drafting → revisions → submission → peer review → revisions. Constant context-switch between sources, manuscript, citation manager.
- Buying brain: institutional Zotero/Mendeley/EndNote already in place. New tool must *complement*, not replace.

## 2. Why their desktop is wasted

- Drafts in Word/Scrivener/Obsidian (focused window).
- 14 browser tabs of sources, JSTOR, arXiv, Substack analytics.
- Citation manager (Zotero) in a perma-tab.
- Scattered "research notes" in Apple Notes / Notion / one_doc_per_question.docx.

What QuickSheet replaces: the *intent* of opening their Notion / TaskPaper "writing dashboard" every hour. With the wallpaper, **the dashboard is always behind the manuscript window** — alt-tab into the editor, alt-tab out, glance at progress, back in.

Critically: **writers protect their focus. The wallpaper appears only when they alt-tab out** — it does not steal focus, never pings. This is exactly what this audience wants from a dashboard. Slack is the opposite (always pings); Obsidian's sidebar is intrusive; QuickSheet's wallpaper is invisible-until-glance.

## 3. Glanceable data they actually want behind windows

1. **Today's word count vs target** — single big cell. The day's score, same as commission earnings for artists.
2. **Manuscript section list** — N rows: chapter/section name, current word count, target, % done (sparkline bar). Click-to-open `r: code chapter-03.md`.
3. **Citation queue** — DOIs / arxiv IDs to be formatted. `cite: 10.1038/nature12373` resolves to a row of author/year/title/journal.
4. **Open-loop research questions** — column of "what was the X about Y again?" — written as discovered, answered later.
5. **Reading queue** — papers/books-to-read list with date added; stale entries dim (cell staleness already a queued feature).
6. **Pomodoro / writing-sprint timer** — large countdown.
7. **Deadlines** — submission, review return, grant due. Red countdown when ≤7d.
8. **A "park here" column** — distracting thoughts get parked, not chased. Single-cell drop.

## 4. Candidate extensions / desktop-only features (ranked)

| Rank | Item                                       | Type | Cost | Hit prob | Why                                                                         |
|------|--------------------------------------------|------|------|----------|------------------------------------------------------------------------------|
| 1    | **Per-cell ticking timer (writing sprint)**| feat | low  | high     | Pomo + word-sprint. Same shared timer feature, fifth persona to want it.    |
| 2    | **Value-driven cell colour**               | feat | low  | high     | Stale tasks dim, missed-deadlines red. Same feature, 6th persona want.       |
| 3    | **In-cell progress bar prefix** (`p: 700/1500`) | feat | low | high  | Per-chapter % visible at a glance. Tiny render change; high WOW.            |
| 4    | `cite:` Crossref / DOI lookup *(shipped)*  | ext  | —    | —        | Already exists. Re-pitch with academic screenshot.                          |
| 5    | `arxiv:` arXiv ID → title/authors/abstract | ext  | low  | high     | Free API; STEM academics live here.                                          |
| 6    | `pubmed:` PMID → citation                   | ext  | low  | med-high | NCBI E-utils; life-science audience.                                         |
| 7    | `wc:` live word-count of file               | ext  | low  | high     | `wc: chapter-03.md` re-reads file every N seconds; pairs with `L:` loop.    |
| 8    | `substack:` subscriber + revenue            | ext  | med  | med      | Substack has no public stats API; would need scraping. Defer.               |
| 9    | `zotero:` last 5 added items                | ext  | med  | med      | Zotero local DB read; cross-platform path discovery. Useful for academics.  |
| 10   | `goodreads:` reading queue                  | ext  | low  | med      | Public RSS per shelf; non-fiction crowd.                                     |

The most overlooked feature for this persona is **the in-cell progress bar (#3)**. Writers measure progress in word counts, and a one-line ASCII bar (`▓▓▓▓▓▓░░░░ 712/1500`) on the wallpaper is what makes the surface emotionally engaging, not just informational. ~30 LOC, high WOW.

## 5. Where they hang out (desktop-relevant first)

- **r/desktops, r/macsetups** — "my writing desk" / "my dissertation desktop" posts are recurring genres.
- **r/writing (3.5M), r/writingadvice, r/screenwriting, r/selfpublish, r/PubTips**.
- **r/academia, r/PhD, r/AskAcademia, r/GradSchool**.
- **r/ObsidianMD (350k), r/Scrivener, r/zettelkasten, r/PKMS** — tool-curious overlap.
- **Mastodon scholar.social** — academics moved here post-Twitter; high signal for STEM.
- **Bluesky academic skies** — humanities side migrated here.
- **Newsletters/blogs:** Patrick Rhone, Daring Fireball's writing-tool side, John Walker's writing-software roundups, Andy Matuschak (note-taking philosophy).
- **People to be visible to:** Andy Matuschak (note-taking + tools-for-thought), Maggie Appleton (open-tools advocate), Steph Ango (Obsidian CEO, friendly to local-first), Tyler Cowen (Marginal Revolution — would post a screenshot if interesting).

## 6. Discoverability hooks

Hero image = **a writer's laptop: VSCode or Obsidian or iA Writer open with a manuscript draft. QuickSheet wallpaper visible around/behind: today's word count cell ("723/1500" with a green bar), chapter list with per-chapter progress bars, two red citation queue rows, a pomo timer in the corner.**

Headlines that land:

- "My writing dashboard is now my wallpaper (open source, your data is CSV)"
- "I track word count, chapters, and citations on my desktop wallpaper instead of in Notion"
- "A local-first ambient surface for academics — your wallpaper as orientation surface"

Avoid: TUI screenshots, "spreadsheet" framing without the "but it's on your wallpaper" follow-up, anything that smells like "yet another Notion competitor."

## 7. Implications (queue these)

1. **Bucket E: in-cell progress bar prefix (`p: 712/1500` → `▓▓▓▓░░ 712/1500`).** ~30 LOC, additive, hits writers immediately, also useful for accountants (per-client status) and artists (commission %), gamedev (project completion). Cross-persona feature.
2. **Build `arxiv:` extension** (Bucket F, post-research). Free API, lowest auth, hits STEM academics directly. Pairs with already-shipped `cite:` (Crossref DOIs).
3. **Bucket A: `examples/manuscript-tracker.csv` + `docs/for-writers.md` + `docs/for-academics.md`.** Two starter sheets — fiction novelist (chapter-by-chapter) and PhD student (lit review + manuscript + citations). One image of each = the entire post.
4. **Bucket C: Andy Matuschak / Maggie Appleton outreach pitch** + r/ObsidianMD post.

Cross-link:
- [lawyers.md](lawyers.md), [accountants.md](accountants.md), [artists-visual.md](artists-visual.md), [gamedev-ttrpg.md](gamedev-ttrpg.md) — same timer + value-colour features.
- The progress-bar feature is **net-new** to the recurring backlog and the highest novel implication of this paper.
