# Persona research index

Phase: market research before next round of code. Goal — identify which professional groups
QuickSheet could realistically grab beyond its current dev-tooling beachhead, then write a
research paper per persona before deciding what to build.

Each paper follows the template:

1. **Profile** — who they are, day-to-day job, scale of US/EN-speaking workforce.
2. **Current toolchain** — what they actually use; what's painful.
3. **Pain points QuickSheet could touch** — concrete, not aspirational.
4. **Candidate extensions / features** — minimum 5, ranked by build cost × hit probability.
5. **Where they hang out** — subreddits, forums, podcasts, Slack/Discords, conferences.
6. **Discoverability hooks** — angle for a "Show HN"-style intro, headline format.
7. **Implications** — 2–3 concrete actions for `## Queued` in `log.md`.

Status legend: `queued` | `in-progress` | `draft` | `done`.

## Persona list

| # | Persona                              | Status   | File                                | Why included                                                                                  |
|---|--------------------------------------|----------|-------------------------------------|------------------------------------------------------------------------------------------------|
| 1 | Developers (refined: SRE/DevOps)     | done     | `developers-sre.md`                 | Current beachhead. Sharpen which sub-segment converts.                                          |
| 2 | Lawyers (solo/small-firm)            | done     | `lawyers.md`                        | High willingness-to-pay-with-attention, fragmented tooling, plain-text deeply familiar.        |
| 3 | Accountants & bookkeepers            | done     | `accountants.md`                    | Spreadsheets are native language; ext angle already started (`accounting-extensions.md`).      |
| 4 | Designers, illustrators, 2D artists  | queued   | `artists-visual.md`                 | Underserved by TUI tools; commission tracking + reference dashboards plausible.                |
| 5 | Indie game devs / TTRPG GMs          | queued   | `gamedev-ttrpg.md`                  | Highly online, evangelize freely, custom-tooling culture, screenshot-loving.                   |
| 6 | Writers, researchers, academics      | queued   | `writers-academics.md`              | Citation, todo, distraction-free desktop — already have one ext (`cite`), can grow.            |
| 7 | Day traders / quant hobbyists        | queued   | `traders.md`                        | Live data on wallpaper is the actual product they buy. `stock`, `fx`, `price` already exist.   |
| 8 | Homelab / sysadmin hobbyists         | queued   | `homelab.md`                        | Adjacent to current audience, k8s/docker/ping/portck extensions already there.                 |
| 9 | Students (CS / STEM)                 | queued   | `students.md`                       | Zero-budget audience but viral. Star-per-install ratio likely high.                            |
|10 | Teachers / educators (K–12, college) | queued   | `teachers.md`                       | Grading dashboards, attendance, schedule — CSV-native. Untapped.                                |

Initial cut: 10 personas. Drop or merge as research surfaces overlap. Add new ones if a Bucket R
run on a paper turns up an adjacent niche worth its own file.

## Workflow

- One paper per `/grow-quicksheet` run, Bucket R.
- Skill self-edit (push direct to main, no PR).
- After paper N is `done`, append its top-3 implications to `log.md ## Queued` so later runs
  can pick them up.
- When all 10 are `done`, switch back to building: rank queued implications by build cost × hit
  probability, then ship one per run via Bucket E or F.
