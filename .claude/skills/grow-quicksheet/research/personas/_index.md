# Persona research index

Phase: market research before next round of code. **Scope locked to `--desktop` wallpaper mode.**
Terminal/TUI users are not the target audience for this research. The differentiator we are
selling is **"your wallpaper is now an interactive grid that lives behind every window."**

That single sentence is the test. If a persona only benefits when QuickSheet runs *inside* a
terminal, they do not belong here.

Comparators to anchor every paper:

- **Rainmeter** (Windows) — skinnable desktop widgets, ~10M users.
- **GeekTool** (macOS) — shell output painted on the desktop.
- **Conky** (Linux) — system-info wallpaper overlays.
- **Übersicht** (macOS) — HTML widgets on the desktop.
- **BumpTop / Stardock Fences** — desktop-as-workspace tools.

Not comparators: lazygit, k9s, visidata, harlequin. Those are TUIs that occupy a terminal pane.

## Paper template

1. **Profile** — who they are, day-to-day, US/EN-speaking headcount.
2. **Why their desktop is wasted** — what currently sits behind their windows (static wallpaper,
   nothing, a Slack channel they never look at). What they *could* see instead.
3. **Glanceable data they actually want behind windows** — concrete cells, not aspirational.
4. **Candidate extensions / desktop-only features** — ranked by build cost × hit probability.
   Every entry must answer "would they keep this on their wallpaper for >1 week?"
5. **Where they hang out** — *desktop-customization* communities (r/unixporn, r/Rainmeter,
   r/desktops, r/macsetups) first; profession-specific second.
6. **Discoverability hooks** — screenshot-first angle; the hero image is a full desktop with
   windows over the grid, never a terminal.
7. **Implications** — 2–3 concrete queued actions for `log.md`.

## Persona list (desktop-mode-relevant)

| # | Persona                              | Status   | File                                | Why their wallpaper is the right surface                                                       |
|---|--------------------------------------|----------|-------------------------------------|------------------------------------------------------------------------------------------------|
| 1 | On-call SRE / DevOps (second monitor) | done     | `developers-sre.md`                 | Second monitor is already a dashboard TV; they buy "glanceable" instinctively.                  |
| 2 | Solo / small-firm lawyers            | done     | `lawyers.md`                        | Single-monitor laptop most of the day; deadline anxiety = always-visible countdowns.            |
| 3 | Accountants & bookkeepers            | done     | `accountants.md`                    | Tax-season deadline storm — desktop is the only surface they never close.                       |
| 4 | Visual designers & illustrators      | done     | `artists-visual.md`                 | Reference grids, color palettes, commission queue pinned behind Photoshop/Procreate windows.    |
| 5 | Indie game devs / TTRPG GMs          | queued   | `gamedev-ttrpg.md`                  | r/unixporn-adjacent crowd; love custom desktops; campaign trackers as wallpaper.                |
| 6 | Writers, researchers, academics      | queued   | `writers-academics.md`              | Distraction-free desktop; word-count, citation queue, draft list visible without context switch.|
| 7 | Day traders / quant hobbyists        | queued   | `traders.md`                        | Multi-monitor culture; tickers behind every window is exactly what they pay Bloomberg for.      |
| 8 | Homelab / sysadmin hobbyists         | queued   | `homelab.md`                        | r/homelab/r/selfhosted; second monitor of Grafana → second monitor of QuickSheet.               |
| 9 | Students (CS / STEM)                 | queued   | `students.md`                       | One laptop, one desktop; deadline + assignment list as wallpaper. Free + zero-install = sticky. |
|10 | Teachers / educators                 | queued   | `teachers.md`                       | Class schedule + grading queue glanceable between teaching blocks.                              |

Drop or merge as overlap surfaces. Add adjacent niches (e.g. "remote workers with single monitor")
only if a queued paper turns one up.

## Workflow

- One paper per `/grow-quicksheet` run, Bucket R, **desktop-mode-only framing**.
- Skill self-edit (push direct to main, no PR).
- After paper N is `done`, append its top-3 implications to `log.md ## Queued`.
- When all 10 are `done`, rank queued implications by build cost × hit probability and resume
  Bucket E (desktop features) / F (extensions that produce glanceable cells).
