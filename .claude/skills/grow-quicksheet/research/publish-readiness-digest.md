# Publish-readiness digest — collapse the PR pile into decisions

**Generated:** 2026-06-22 by grow-quicksheet run.

## Summary (≤5 lines)

- Stars: 0. The bottleneck is **not** lack of material — it is ~80 open `grow/*` PRs the user has not triaged, plus dozens of finished drafts.
- The skill's own cron runs created **redundant PRs**: NO_COLOR ×4, `--info` ×4, stdin ×3, `--delimiter` ×3, help-examples ×4, shell-completions ×2, `--sort` ×2.
- Every open issue across the whole repo family already has ≥1 covering PR → producing more PRs is pure padding. **Stop adding; start triaging.**
- Below: one recommended PR per cluster + the duplicates to close. This turns an unreviewable 80-PR pile into ~12 yes/no decisions.
- All listed PRs report `MERGEABLE` (no conflicts) as of generation, except directory-sync PRs which touch the same file and will conflict pairwise.

## Recommended decisions — issue-backed clusters

For each, merge the **Keep** PR, close the **Close** duplicates (older / superseded). Closing a PR is reversible; it removes review noise.

| Issue | Feature | Keep (merge) | Close (duplicates) |
|-------|---------|--------------|--------------------|
| #374 | `--help` usage examples | **#470** (newest "Examples section") | #376, #379, #413 |
| #372 | shell completion (bash/zsh) | **#393** | #378 |
| #321 | `--info` (size + mtime) | **#426** (newest) | #357, #394 |
| #320 | read CSV from stdin | **#397** (pipe) | #360, #396 |
| #319 | `--delimiter` / TSV | **#447** (newest) | #368, #403 |
| #318 | `NO_COLOR` env var | **#442** (also adds `--no-color` flag) | #367, #391, #411 |

Net: 6 merges resolve 6 open issues; 10 stale duplicate PRs closed.

## Independent feature PRs — each unique, user picks à la carte

No duplicates; decide on merit:

- #453 `--sort` / `--sort-desc` (supersedes older #382 — close #382 if #453 preferred)
- #388 `--export-sql`, #422 `--export-tsv`, #437 `--export-yaml`, #473 `--export-latex`
- #439 `--print` pretty table, #384 `--grep` row filter, #365 `--columns`, #361 `--head`/`--tail`
- #386 Catppuccin / Tokyo Night themes (most screenshot-worthy of this batch)

Note: the `--export-*` family + `--print`/`--grep`/`--columns`/`--sort`/`--head`/`--tail` are a coherent "headless data-tooling" story. If the user wants one flagship PR instead of eight, that's a consolidation task for a future run — flag, don't auto-do.

## Directory-sync PRs — CONFLICT cluster (same file `docs/extensions.md`)

These edit the same directory file and will conflict once one merges:

- #450 (adds 22, 69→90) — most comprehensive, **merge first**
- #385 (adds 17, →85+) — likely subset of #450; close after #450 merges
- #477 (dict/gitlog/stocks), #475 (gomod), #464 (aqi) — rebase onto merged #450, then merge whichever rows are still missing

Action: merge #450, then close/rebase the rest. Do NOT merge two directory PRs without rebasing.

## Landing-page PRs — one duplicate

- #359 and #416 are **both** Dungeon-Master landing pages → keep one, close the other.
- Unique: #364 freelancer, #390 security, #400 artist, #419 writer, #431 accountant. Decide whether the README should link this many persona pages (risk: docs sprawl). Recommend keeping 2–3 strongest (accountant, writer, freelancer) and closing the rest unless a real visitor asked.

## Skill-log PRs

Many `grow/log-*` PRs are skill-only bookkeeping that, per the skill's own rules, should have been **direct-to-main commits, not PRs**. They can be batch-closed; the log content is already on those branches and low-value to merge. Future runs must push skill-only changes straight to `main` (no PR) per the skill-self-edit exception.

## Implications for QuickSheet (the actual deliverable)

1. **The next 3–5 grow runs should be NO-OPs or triage digests, not new PRs.** The pile is the problem; adding to it lowers the user's odds of ever merging anything. (Reinforces [[feedback-stop-padding]].)
2. **Stars come from publishing, not from PRs.** All the launch drafts (showhn.md, unixporn-rice.md, reddit-*.md) are gated on the user capturing **one** wallpaper screenshot (`A1`). That single human action unblocks the entire Bucket C/D pipeline. Everything else is downstream of it.
3. **Fix the cron-dup root cause:** runs keep re-implementing the same 6 issues because they sweep issues but not *open PRs targeting those issues*. Future runs MUST check `gh pr list` for an existing PR referencing the issue number before implementing. Queue this as a hard pre-check.

## Queued actions (added to log)

- Pre-implementation guard: every feature run greps open PR titles/bodies for the issue number first; if covered, no-op.
- When user captures the `A1` wallpaper screenshot: skill swaps README hero (PR) + writes r/unixporn + Show HN final copy.
