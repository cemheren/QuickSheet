---
name: grow-quicksheet
description: >
  Autonomously pick and execute one high-leverage action to grow GitHub stars on
  QuickSheet. Designed for scheduled, unattended runs (no human present). Picks
  one action, does it, logs it, reports summary. Never asks the user mid-run.
  Trigger: /grow-quicksheet, "grow quicksheet", or scheduled invocation.
---

# Mission

Grow GitHub stars on github.com/cemheren/QuickSheet. The project is a zero-dependency
.NET terminal spreadsheet that also runs as a Linux/Windows desktop wallpaper. It is
genuinely useful — job is to make more people aware of it and make the first-touch
experience better.

You have **limited time per run** and **limited total runs**. Treat each run as one
focused action, not a sprawl. Compounding daily progress beats one giant push.

# How a run works

Assume no human is watching. Do not ask questions. Do not block on approval.
Pick, execute, log, report.

1. **Read the log** at `.claude/skills/grow-quicksheet/log.md`. This is durable
   memory across runs — what was tried, what worked, what didn't, what's queued.
2. **Check current state**: `gh repo view --json stargazerCount,forks,issues` and
   note star count + delta since last run.
3. **Pick ONE action** from the menu below. Selection rules:
   - Prefer items under `## Queued` in the log.
   - Bias toward variety: do not repeat the same bucket two runs in a row.
   - Prefer high expected value × low risk.
   - If unsure, default to Bucket A (product polish, local-only).
4. **Execute** the action end-to-end. No mid-run questions.
5. **Update the log** with date, action, outcome, star count, follow-ups.
6. **Open a PR for every code/doc change.** Never push directly to `main`.
   Per run:
   - `git checkout -b grow/<short-slug>`.
   - Commit (Conventional Commits, no Claude attribution lines).
   - Push the branch.
   - `gh pr create --base main --head <branch> --title "..." --body "..."`.
   - User merges. The skill never merges its own PRs.
   - For destructive / publishable-elsewhere actions, drafts still go to
     `.claude/skills/grow-quicksheet/drafts/` — also on a branch + PR.
7. **Report** a 3-6 line summary at end of run: action, PR URL, star delta, next.

## Action priority

1. **Outstanding issues first — across the whole repo family.** Sweep open
   issues on the main repo *and every* `cemheren/quicksheet-*` extension repo:
   ```bash
   gh issue list --repo cemheren/QuickSheet --state open --json number,title,labels,url
   for r in $(gh repo list cemheren --limit 200 --json name \
       -q '.[].name | select(. | startswith("quicksheet"))' ); do
     gh issue list --repo "cemheren/$r" --state open --json number,title,labels,url \
       | jq --arg r "$r" 'map(. + {repo: $r})'
   done
   ```
   Pick the smallest qualifying issue across that union. Fix on a branch *in the
   repo where the issue lives* (main repo OR the extension repo). PR with
   `Closes #N` in body, opened against that repo. The user merges. The skill
   never merges its own PRs.

2. **No open issues anywhere → feature PRs are fair game.** Either:
   - A new feature on the main QuickSheet repo (Bucket E menu), OR
   - A targeted feature/fix on an existing extension repo, OR
   - A new extension repo (Bucket F workflow — still create the ext via `gh repo
     create`, then file the README/tour cross-link change against QuickSheet as a PR).

3. **Other buckets (A/B/C/D/R)** still allowed, each via PR.

The "plausibly causes stars" filter is *relaxed* under this regime — feature
work is welcome even if star-ROI is not certain, because PRs let the human gate
what merges.

# Action menu

## Bucket A — Product polish (local, low-risk, autonomous-safe)
- README improvements: clearer hero, sharper first-paragraph pitch, animated demo
  GIF near top, badges, install one-liner, "Why this exists" section
- Add new screenshot showing compelling use (live process output, weather extension,
  hyperlinks, formula cells, loop coloring)
- Write `docs/` page: "tour of features in 60 seconds"
- Fix an open issue or TODO in `roadmap.md`
- Add small example extension under `Extensions/`
- Improve `--help` output and first-run UX
- Polish error messages, copy, typos

## Bucket B — Discoverability (metadata, SEO, autonomous-safe via gh)
- Set/refine repo description: `gh repo edit --description "..."`
- Set topics: `gh repo edit --add-topic terminal,tui,spreadsheet,csv,dotnet,...`
- Set homepage if a site exists
- Write `CONTRIBUTING.md` (small, alive signal)
- Cut a GitHub release with built binaries when feature-meaningful change exists

## Bucket C — Content drafts (DRAFT ONLY, never publish)
Save to `.claude/skills/grow-quicksheet/drafts/<topic>.md`. User publishes manually
on their own time. Log as "draft saved at <path>".
- Hacker News "Show HN" post — title + first comment
- Reddit posts: r/commandline, r/dotnet, r/programming, r/linux
- Lobsters submission text
- Twitter/X thread (3-5 tweets)
- Mastodon/Bluesky post
- dev.to / Medium blog post

## Bucket D — Network effects (DRAFT ONLY, never push to other repos)
- Identify awesome-* lists this project fits. Draft PR description + diff snippet.
  Save to `drafts/awesome-<listname>.md`. User submits.
- Identify terminal-tool directories (terminaltrove, console.dev). Draft submission.

## Bucket F — Vertical extensions (scaffold locally, create repo, push)

Extensions live in **separate GitHub repos** referenced via `ext: github:user/repo`.
QuickSheet itself stays small; the network of `quicksheet-*-ext` repos is the
discoverability flywheel.

For inspiration, see existing extensions linked in the README (e.g.
`quicksheet-copilot-ext`). Manifest + JSON-lines protocol details are in README
under "Extensions — Make Your Desktop Do More".

Workflow:
1. Pick a vertical. Scaffold the extension at
   `.claude/skills/grow-quicksheet/drafts/extensions/<name>/`.
2. Build it locally with `dotnet build` — must be green, zero NuGet deps.
3. Clean build artifacts: `rm -rf bin obj`.
4. Init local repo + commit: `cd <draft-path> && git init -b main && git add . && git -c user.email=cemheren@gmail.com -c user.name=cemheren commit -m "Initial: <pitch>"`.
5. Create the GitHub repo + push:
   `gh repo create cemheren/<name> --public --description "..." --source=. --push`.
6. **Critical:** Back in the QuickSheet repo root, the inner `.git` makes the draft
   look like a submodule (gitlink, mode 160000) when added to the outer index.
   Remove it: `git rm --cached <draft-path> && rm -rf <draft-path>/.git && git add <draft-path>`.
   Then commit the draft files normally in the outer repo so the scaffold travels
   with the QuickSheet repo as plain files.
7. Add a link to the new extension in QuickSheet's main README **and**
   `docs/tour.md` (same commit).

Never push extension code to the QuickSheet repo itself — it goes to its own
new repo on the user's account.

Vertical ideas (research-driven — pick one per run):
- **Tax / accounting**: `ext: tax: <amount>,<state>` → returns federal+state tax
  estimate; `ext: 1099: <gross>` → quarterly estimated payment.
- **Legal**: `ext: case: <citation>` → case lookup (free Caselaw API); `ext: stat:
  <us-code>` → statute snippet.
- **Real estate**: `ext: zillow: <address>` → Zestimate; `ext: mortgage:
  <principal>,<rate>,<years>` → monthly payment.
- **Crypto / finance**: `ext: price: <symbol>` → last trade; `ext: yield: <ticker>`
  → dividend yield. Use free APIs (CoinGecko, Yahoo via yfinance JSON).
- **Devops / SRE**: `ext: ping: <host>` → status code + latency; `ext: tls:
  <host>` → cert expiry days.
- **Writing / research**: `ext: define: <word>`, `ext: thes: <word>`, `ext:
  cite: <doi>` → formatted citation.
- **Email / contacts**: `ext: gravatar: <email>`, `ext: mxck: <domain>` → MX
  records.

What each draft must include:
- `README.md` — install one-liner (`ext: github:user/repo`), examples, demo
  screenshot placeholder.
- `manifest.json` (or whatever QuickSheet's protocol spec calls it — read the
  protocol section in the main README before drafting).
- A minimal working implementation in **any zero-dep language** (Python stdlib,
  Go, Rust, .NET) — pick whatever's smallest.
- License (MIT).
- A 2-sentence pitch paragraph for the eventual repo description.

Submission strategy:
- Tag extension with the vertical's terms (e.g. `tax,1099,accounting`).
- Cross-link from QuickSheet README (separate Bucket A action — wait for user
  to push the extension repo first).

**Hard rules**:
- Never push extension code to the QuickSheet repo. Drafts only.
- Same boundaries apply: no social posting, no NuGet deps in dotnet projects,
  truthful claims.
- One extension draft per run, not a sprawl of half-done verticals.

## Bucket R — Deep research (use WebSearch + WebFetch)

Spend a run *researching* instead of *producing*. Goal: gather concrete intel that makes the next 5–10 publish/PR actions hit harder. Save findings as a markdown brief in `.claude/skills/grow-quicksheet/research/<topic>.md`.

Good research targets (pick one per run):
- **Comparable TUI launch case studies** — repos that crossed 500 / 1k / 5k stars from a single HN or Reddit moment. Title pattern, posting time, who they replied to, what the top comment was. Distill into a "what worked" cheat sheet.
- **Specific niche communities** — beyond the big subs, find smaller forums where this project specifically fits (e.g. r/Rainmeter, r/unixporn, r/i3wm, vim/emacs lists, terminal-tool newsletters like Console Weekly, TerminalTrove podcasts). For each: rules, posting cadence, what they reward.
- **Top contributors / influencers** for terminal-tool genre — accounts that signal-boost when they discover something they like. Curate a small list (≤10) with what each is known for. No mass-tagging; this is *who to be visible to*, not who to spam.
- **Adjacent-project teardown** — pick 2–3 projects QuickSheet is adjacent to (e.g. `visidata`, `lazygit`, `nyxt`, `harlequin`) and read their README / `docs/` / launch posts. What hooks do they lead with? What does the first screenshot show? What's their one-line pitch? Steal the *structure* of what works, not the words.
- **What works on HN for "Show HN: a TUI..."** — search HN for the last year of TUI submissions, capture title + score + comments. Identify three patterns that correlate with high scores.

Output of each research run is a `research/<topic>.md` with:
- Bullet summary at top (≤5 lines).
- Concrete data points with URLs.
- "Implications for QuickSheet" section — 2–3 specific actions to queue in the log.

Then queue those actions in the log's `## Queued` section so the next run picks one.

Hard rule: research without distilled implications is wasted work. The deliverable is the implications, not the survey.

## Bucket E — Quality-of-life features that get screenshotted
Code change in QuickSheet repo. Autonomous-safe — commit and push.
- Theme presets
- Sparkline-in-cell rendering
- Live web fetch cell prefix (`w: url`)
- Markdown export
- Vim-style keybinding mode
- Better autocomplete on cell prefixes

Pick small. One feature per run. Keep CLAUDE.md conventions:
zero NuGet deps, CSV persistence, cross-platform pattern.

**Build before commit.** Run `dotnet build ExcelConsole.csproj`. If build fails,
fix until green. Do NOT commit broken code. If `dotnet` is unavailable in the
environment (e.g., remote sandbox without .NET 9 SDK), skip Bucket E this run
and pick from A–D instead. Log "skipped: no dotnet" rather than shipping
unverified code.

# Selection rule (relaxed under the PR regime)

PRs make the human the gatekeeper, so produce liberally — the filter is mostly
"is this a real change, truthful, and not breaking?" rather than "am I sure
this causes stars?" Specifically:

- Feature PRs are welcome on the main repo. Even if star-ROI is uncertain, the
  user can decline to merge.
- New extension repos are welcome. Same gating: cross-link PR against QuickSheet
  can be declined if the ext doesn't earn its place.
- Bucket-variety is fine. Pick the highest-quality concrete action available.

Still **avoid:**
- Breaking the build or existing features (see Boundaries).
- Re-doing work already in `drafts/` or already merged.
- Manufactured filler (duplicate drafts, near-identical extensions, flag bikesheds).

If genuinely nothing concrete to do, log a no-op and exit.

# Boundaries (hard rules — no exceptions)

- **DO NOT BREAK THE PROJECT.** A broken repo loses stars, doesn't earn them.
  Before every code/doc commit:
  - Run `dotnet build ExcelConsole.csproj` — must be 0 errors, 0 warnings.
  - Do not change existing public API or rename CLI flags without a deprecation path.
  - Do not modify keybindings, default behavior, or UI layout that users rely on.
  - Do not remove or rename screenshots referenced in README.
  - Do not change CSV format, autosave path, or persistence layout.
  - Touch the smallest surface area that achieves the action. Additive > refactor.
  - If a change is non-trivial (>~50 lines, or touches >2 files outside docs),
    skip and pick a smaller action this run.
  - If build fails and the fix isn't obvious, revert local changes and switch to a
    Bucket A/B/C/D action instead. Never push broken code.
- **No social posting.** Never post to HN, Reddit, Twitter, Mastodon, Bluesky,
  Lobsters, dev.to, Medium, etc., even if credentials exist. Drafts only.
- **No PRs to other repos.** Draft branch + PR body saved to `drafts/`. User submits.
- **No paid promotion, no bots, no astroturfing, no fake accounts.**
- **Truthful claims only.** No "production-grade" / "thousands of users" lies.
- **Never push to `main` directly.** All changes go via feature branch + PR.
  User merges. No force-push, no history rewrite, no branch deletion either.
- **One action per run.** Pick, execute, log, stop.
- **No NuGet dependencies added.** Hard repo policy.

# Log format

`.claude/skills/grow-quicksheet/log.md` is persistent memory. Append entries:

```
## YYYY-MM-DD

- Stars: N (Δ +M since last)
- Action: <one-line summary>
- Bucket: A/B/C/D/E
- Outcome: <shipped commit abc1234 | draft saved at path | blocked on X>
- Follow-up: <next step queued, or "none">
```

Keep entries short. Maintain a `## Queued` section at bottom for next-run hints.

# End-of-run report

Print to user (always, even if no human will read it — the log of runs matters):

```
QuickSheet grow run YYYY-MM-DD
Stars: N (Δ +M)
Did: <action>
Outcome: <commit hash | draft path>
Next: <queued follow-up>
```

# First run

If `log.md` does not exist, create it with header and a `## Queued` seed:
"Audit README first-impression; pick top fix." Then execute that audit as the
first action.
