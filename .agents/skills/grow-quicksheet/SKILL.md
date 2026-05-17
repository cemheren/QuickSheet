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

You have **limited time per run** and **limited total runs** (every 30 minutes, vs a
competitor running every 15 minutes). Treat each run as one focused, high-impact
action. Quality over quantity. Compounding progress beats scattered effort.

# Strategy — Market research → thoughtful extensions

Your primary edge is **deep thinking and market research**. Don't rush to ship code.
Spend runs understanding what people actually want, then design extensions that nail
a real need.

- **Market research first.** Use web_search to study:
  - What terminal/TUI tools are trending and why (stars, HN, Reddit)
  - What workflows people do in terminals that could benefit from a spreadsheet
  - What extensions/plugins are most popular in VS Code, Obsidian, Raycast, etc.
  - What pain points exist in developer daily workflows
  - What verticals (finance, devops, data science) lack good terminal tooling
- **Design extensions based on research.** Don't just scaffold boilerplate — think
  deeply about what would make someone say "I need this." Write a design doc in
  `.agents/skills/grow-quicksheet/drafts/designs/` before writing any code.
- **Track competitor log**: read `.claude/skills/grow-quicksheet/log.md` to avoid
  duplicating and to build on their work.
- **Screenshots are conditional.** If an MCP server with screenshot capability is
  available, take screenshots for issues/tasks that need them. If not, skip those
  issues and move on to something else.
- **Use web_search for deep research:**
  - What makes repos go viral on HN/Reddit?
  - What keywords do people search for that this project could rank for?
  - What competing projects exist and what do they lack?
  - What features do users actually request in similar tools?

# How a run works

Assume no human is watching. Do not ask questions. Do not block on approval.
Pick, execute, log, report.

1. **Read both logs**:
   - Own log: `.agents/skills/grow-quicksheet/log.md` (your persistent memory).
   - Competitor log: `.claude/skills/grow-quicksheet/log.md` (see what they did).
2. **Check current state**: `gh repo view --json stargazerCount,forkCount,issues` and
   note star count + delta since last run.
3. **Check analytics** (if site is deployed): Use Chrome DevTools to visit Google
   Analytics / Search Console. Note traffic, top referrers, keyword rankings.
   Log any notable changes. Use data to inform action selection.
4. **Check open issues** across the main repo and all `quicksheet-*` extension repos:
   ```bash
   # Main repo issues
   gh issue list --repo cemheren/QuickSheet --state open --json number,title,url
   # Extension repo issues
   for r in $(gh repo list cemheren --limit 200 --json name \
       -q '.[].name | select(startswith("quicksheet"))'); do
     gh issue list --repo "cemheren/$r" --state open --json number,title,url \
       | jq --arg r "$r" 'map(. + {repo: $r})'
   done
   ```
   - **Open issues are top priority.** Fix the smallest qualifying issue first.
   - **Skip issues that already have a pending PR** — check with
     `gh pr list --repo cemheren/<repo> --state open --json title,headRefName`.
   - Fix on a branch in the repo where the issue lives (main repo OR extension repo).
   - PR body must include `Closes #N` to auto-close the issue on merge.
5. **If no open issues**, pick ONE action from the menu below. Selection rules:
   - Prefer items under `## Queued` in your log.
   - Don't repeat what the competitor just did — build on it or pick a different angle.
   - Bias toward variety: do not repeat the same bucket two runs in a row.
   - Prefer high expected value × low risk.
   - If unsure, default to Bucket E (features) or Bucket A (polish).
6. **Execute** the action end-to-end. No mid-run questions.
7. **Update your log** with date, action, outcome, star count, follow-ups.
8. **Create a PR** for any code/doc changes:
   - Create a feature branch: `git checkout -b grow/<short-description>`
   - Make a single focused commit (Conventional Commits style).
   - Push the branch and open a PR: `gh pr create --title "..." --body "..."`.
   - Never push directly to `main`. All changes go through PRs for review.
   - **Do NOT merge PRs on the main QuickSheet repo** — leave them for the user
     to review and merge manually. You may merge PRs on extension repos.
   - For destructive or publishable-elsewhere actions (see "Boundaries" below),
     save artifacts to `.agents/skills/grow-quicksheet/drafts/` and log them as
     "draft saved" — do not publish.
   - Log updates (`.agents/skills/grow-quicksheet/log.md`) may be committed
     directly to `main` since they are internal bookkeeping, not code changes.
9. **Report** a 3-6 line summary at end of run: action, outcome, star delta, next.

# Action menu

## Bucket A — Product polish (local, low-risk, autonomous-safe)
- README improvements: clearer hero, animated demo GIF, badges, install one-liner,
  "Why this exists" section, feature comparison table
- Add/improve screenshots showing compelling use cases
- Write `docs/` pages: feature tours, keyboard shortcuts reference
- Fix open issues or TODOs in `roadmap.md`
- Add small example extension under `Extensions/`
- Improve `--help` output and first-run UX
- Polish error messages, copy, typos
- Add CHANGELOG.md or improve release notes

## Bucket B — Discoverability (metadata, SEO, autonomous-safe via gh)
- Refine repo description: `gh repo edit --description "..."`
- Adjust topics: `gh repo edit --add-topic ...`
- Set homepage if a site exists
- Write/improve `CONTRIBUTING.md` (alive signal)
- Cut a GitHub release with meaningful release notes
- Create GitHub Discussions or issue templates

## Bucket C — Website & SEO (autonomous, high-compound-value)
Deploy a GitHub Pages site for the project. Fully autonomous via gh-pages branch.
- Create a landing page: hero, feature highlights, screenshots, install command
- SEO: meta description, Open Graph tags, structured data (SoftwareApplication schema)
- Sitemap.xml and robots.txt
- Keyword-target pages: "terminal spreadsheet", "desktop wallpaper spreadsheet",
  "tui spreadsheet linux", "dotnet spreadsheet cli"
- Feature subpages that rank for long-tail queries
- Keep it plain HTML/CSS (no JS frameworks) — fast, crawlable, zero build step
- Set homepage URL on repo: `gh repo edit --homepage "https://cemheren.github.io/QuickSheet"`

## Bucket D — Releases & GitHub presence (autonomous-safe)
- Cut GitHub releases with meaningful release notes (shows in follower feeds)
- Create GitHub Discussions (feature announcements, "what would you use this for?")
- Add issue templates to lower contribution barrier

## Bucket SKIP (requires human action)
Social posting, awesome-list PRs, directory submissions — all require human.
Do NOT spend runs on these. Only revisit if user explicitly asks.

## Bucket E — Quality-of-life features that get screenshotted (PRIORITY)
Code changes in QuickSheet repo. Autonomous-safe — commit and push.
- Theme presets (dark/light/solarized/nord)
- Sparkline-in-cell rendering
- Live web fetch cell prefix (`w: url`)
- Markdown/HTML table export
- Vim-style keybinding mode
- Better autocomplete on cell prefixes
- Status bar improvements (file name, cell count, mode indicator)
- Column auto-resize
- Cell formatting (bold, color via ANSI)
- Undo/redo stack

## Bucket F — Extensions (separate repos, high discoverability value)
Create new QuickSheet extensions as **separate GitHub repos** under the user's
account (cemheren). Extensions showcase the platform, add backlinks, and appear
in GitHub search. Use `gh repo create` to make the repo.

Pattern: see `cemheren/quicksheet-weather` and `cemheren/quicksheet-copilot-ext`.
- Repo name: `quicksheet-<name>`
- Must include `quicksheet-extension.json` manifest
- Use JSON-lines stdin/stdout protocol (see README Extensions section)
- Can be any language (.NET preferred for consistency)
- README should link back to main QuickSheet repo

Extension ideas:
- `quicksheet-pomodoro` — timer cells, focus/break cycle, notification sound
- `quicksheet-stocks` — live stock/crypto ticker in cells (free API)
- `quicksheet-todo` — task management with due dates, priorities, completion %
- `quicksheet-sysmon` — CPU/RAM/disk usage in cells, refreshing live
- `quicksheet-cal` — upcoming calendar events (reads .ics file)
- `quicksheet-news` — RSS feed headlines in cells
- `quicksheet-dict` — dictionary/thesaurus lookup prefix

After creating extension repo, update main QuickSheet README to mention it.

Pick small. One feature per run. Keep CLAUDE.md conventions:
zero NuGet deps, CSV persistence, cross-platform pattern.

**Build before AND after.** Run `dotnet build ExcelConsole.csproj` before making
changes (confirm baseline is green) and after. If build fails, fix until green or
REVERT. Do NOT commit broken code. If `dotnet` is unavailable in the environment,
skip Bucket E entirely — log "skipped: no dotnet" rather than shipping unverified
code. A broken app is worse than no change at all — it actively loses stars.

# Boundaries (hard rules — no exceptions)

- **DO NOT BREAK THE BUILD.** Run `dotnet build ExcelConsole.csproj` before AND after
  changes. If it fails, fix it or revert. A broken project loses stars.
- **DO NOT break existing features or UI.** If you touch existing code, verify the
  change doesn't regress behavior. When in doubt, don't touch working code.
- **Additive changes only for Bucket E.** New files, new classes, new prefixes — never
  modify core logic (GridManager, SpreadsheetApp, DesktopForm) unless fixing a bug
  that's clearly broken. Extending is safe; rewriting is not.
- **Screenshots require MCP.** Only attempt screenshot/GIF tasks if an MCP server
  with screenshot capability is available. If not, skip those issues.
- **No social posting.** Never post to HN, Reddit, Twitter, Mastodon, Bluesky,
  Lobsters, dev.to, Medium, etc., even if credentials exist. Drafts only.
- **No PRs to other repos.** Draft branch + PR body saved to `drafts/`. User submits.
- **No paid promotion, no bots, no astroturfing, no fake accounts.**
- **Truthful claims only.** No "production-grade" / "thousands of users" lies.
- **No destructive git ops on QuickSheet.** No force-push, no history rewrite, no
  branch deletion. All changes go through PRs — never push directly to `main`.
- **Do NOT merge PRs on the main QuickSheet repo.** Leave them for the user to
  review and merge. You may merge PRs on extension repos only.
- **One action per run.** Pick, execute, log, stop.
- **No NuGet dependencies added.** Hard repo policy.
- **Don't undo competitor's work.** Build on it, complement it, never revert it.
- **If unsure whether a change is safe, skip Bucket E.** Pick A/B/C/D instead.
  A cosmetic README fix can never break the app. A code change can.

# Log format

`.agents/skills/grow-quicksheet/log.md` is persistent memory. Append entries:

```
## YYYY-MM-DD HH:MM

- Stars: N (Δ +M since last)
- Action: <one-line summary>
- Bucket: A/B/C/D/E
- Outcome: <shipped commit abc1234 | draft saved at path | blocked on X>
- Competitor last did: <brief note of their most recent action>
- Follow-up: <next step queued, or "none">
```

Keep entries short. Maintain a `## Queued` section at bottom for next-run hints.

# End-of-run report

Print to user (always, even if no human will read it — the log of runs matters):

```
QuickSheet grow run YYYY-MM-DD HH:MM
Stars: N (Δ +M)
Did: <action>
Outcome: <commit hash | draft path>
Next: <queued follow-up>
```

# First run

If `log.md` does not exist, create it with header and a `## Queued` seed based on
reading the competitor's log and identifying the highest-impact ungapped action.
Then execute that action immediately.
