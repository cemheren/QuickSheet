# PR backlog merge-triage cheat sheet — 2026-06-24

## Summary (≤5 lines)
- 50 open PRs on `cemheren/QuickSheet`, 0 stars. A 50-PR pile reads as "chaotic/abandoned" to a visitor deciding whether to star — this is now a *first-impression* problem, not just a queue problem.
- The bottleneck is **merge throughput**, not production (56 ext repos already exist; more production has not moved stars).
- Many issues have 2–4 duplicate PRs; several PRs are **contaminated giant diffs** (+3012/+2898/+3015) from bad rebases — close on sight.
- 14 open PRs are skill-authored `chore(skill)` log/draft commits that, under the current SKILL.md rule, should be **direct-to-main commits, not PRs**.
- Action for the user: merge one PR per issue (list below), close the rest, clear the skill-only noise. ~5–10 min → backlog 50 → ~15.

## Recommended merges (one per issue — smallest clean diff, all MERGEABLE)
| Issue | MERGE this | Why | CLOSE these |
|------|-----------|-----|-------------|
| #318 NO_COLOR | **#411** (+15) | minimal, exactly the issue | #442 (+54, adds extra `--no-color` flag — scope creep; merge instead of #411 only if you want that flag), #391 (+2898 contaminated) |
| #319 delimiter | **#447** (+253) | newer of an identical pair | #403 (+253, older dup) |
| #320 stdin | **#484** (+66) | clean, recent, `Closes #320` | #396 (+44, alt `-` arg approach), #397 (+3014 contaminated) |
| #321 --info | **#426** (+57) | clean | #394 (+3015 contaminated) |
| #372 completions | **#393** (+123) | only remaining PR | — |
| #374 help examples | **#486** (+8) | tiny, `closes #374` | #485 (+9 near-dup), #470 (+33), #413 (+208) |

## Contaminated PRs — close outright (bad branch base, thousands of unrelated lines)
- #391 (+2898), #397 (+3014), #394 (+3015), #390 (+3012). Review #413 (+208) and #405 (+735) — likely same issue.

## Independent feature/ext-crosslink PRs (the "(none)" cluster, 21 PRs)
These don't conflict with each other — review on merit, but they're the bulk of the pile. High-value, low-risk picks to merge first:
- Headless export modes: #388 (--export-sql), #422 (--export-tsv), #437 (--export-yaml), #473 (--export-latex), #439 (--print), #453 (--sort). These are additive, screenshot-friendly, and several issues asked for export flexibility. Consider merging the cleanest 2–3 and closing redundant ones.
- Extension-directory sync PRs (#477, #475, #464, #450, #455, #435) keep re-drifting. **Merge #450 (the big "69→90" sync) and close the smaller incremental ones** — they overlap. Going forward, stop opening one sync PR per ext; batch.
- Landing pages (#431 accountants, #419 writers, #416 DMs, #400 artists, #390 security): #390 is contaminated; the others are clean docs. Merge on taste.

## Skill-only noise PRs — should be direct-to-main, close as PRs
Per SKILL.md ("changes scoped entirely to `.claude/skills/grow-quicksheet/` … pushed directly to main. No PR for skill self-edits"), these 14 should not be PRs:
#488, #487, #410, #409, #408, #406, #404, #402, #399, #398, #395, #392, #389 (and #483 is itself a cleanup PR).
- **Before closing**, confirm each one's draft/log content is already on `main` (drafts under `drafts/`, log in `log.md`). If not committed, the content is lost on close. PR #483 already did a batch of 52 this way.

## Implications for QuickSheet (queue these)
1. **Stop opening per-extension directory-sync PRs.** Batch directory updates, or generate the directory from a manifest. Each one-off sync PR is filler that re-drifts. → queue a single "generate ext directory from list" task.
2. **Skill log/draft commits go direct to main**, never as PRs. Audit that the cron path does this (recent runs still opened #487/#488).
3. **A clean PR list is itself a star lever.** Drained backlog + recent merges = "actively maintained" signal on the repo homepage. Re-check pile size next run; if still >30, no new production — the move is triage, not more PRs.
