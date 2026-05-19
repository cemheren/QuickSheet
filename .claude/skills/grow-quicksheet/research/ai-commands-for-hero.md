# AI commands for hero PNG — ranked top 10

Goal: pick the AI-tool commands a scrolling dev recognizes within 1 second. Hero shows QuickSheet as a wallpaper launcher/dashboard for an AI-coding workflow.

## TL;DR (5 lines)

- **Claude Code dominates the slash-command universe right now**: `/init`, `/compact`, `/cost`, `/clear`, `/review` are the universally cited daily-driver set.
- **Copilot CLI is 2 commands**: `gh copilot suggest "<x>"` + `gh copilot explain "<cmd>"` (plus the `ghcs`/`ghce` aliases). That's it. Both are recognizable instantly.
- **Aider's hot pair is `/ask` ↔ `/code`** (plan-then-execute pattern). `/add <file>` rounds out the trio.
- **Copilot Chat in-IDE**: `/explain`, `/fix`, `/tests`, `/doc` are the canonical four.
- For QuickSheet hero: pick **5 runnable launchers + 1 live cost cell + 1 extension cell** — that mix shows off prefixes (`r:`, `i:`, `ext:`) without looking staged.

---

## Ranked top 10 — by recognizability for an AI-developer audience

Rank weighted: **recognition** (dev sees the command, knows the tool) × **demo value** (visually meaningful as a cell). High-rank = must appear in hero.

| # | Command | Tool | Why it's in | Cell prefix |
|---|---------|------|-------------|-------------|
| 1 | `claude` | Claude Code | Brand pull. Single token. Most recognizable AI-CLI invocation on Reddit/HN in 2026. | `r: claude` |
| 2 | `claude --continue` | Claude Code | Daily-driver per every workflow blog. "Resume last session" — instantly understood. | `r: claude --continue` |
| 3 | `gh copilot suggest "..."` | GitHub Copilot CLI | The Copilot CLI's flagship verb. Recognizable even to non-users from the docs. | `r: gh copilot suggest "find files larger than 1GB"` |
| 4 | `/compact` | Claude Code | THE cost-management ritual. Every "23 tips" / "60% cost reduction" blog leads with it. | text cell: `/compact (run at 70% ctx)` |
| 5 | `/cost` | Claude Code | Universally cited token-usage check. Visual: pair with a `s:` sparkline of last week's spend. | text + sparkline neighbour |
| 6 | `/init` | Claude Code | Onboarding ritual — creates CLAUDE.md. Every onboarding doc leads here. | text cell: `/init → CLAUDE.md` |
| 7 | `/ask` (aider) | aider | The plan-first half of aider's hot loop. Distinct, terminal-y. | `r: aider --ask "refactor parser"` |
| 8 | `gh copilot explain "<cmd>"` | GitHub Copilot CLI | Pair with #3. "Explain this bash one-liner." | `r: gh copilot explain "tar xzvf"` |
| 9 | `/review` | Claude Code | PR review ritual. Pairs naturally with `r: gh pr view` cells. | text cell: `/review (in claude)` |
| 10 | `/fix` `/tests` `/explain` `/doc` | Copilot Chat | The IDE-side four. One row, four cells, shows breadth. | row of four text cells |

Sources at bottom.

## Lower-ranked (didn't make hero, kept for reference)

- `claude -p "<prompt>"` — headless mode. Useful but visually identical to `claude` to most viewers.
- `claude --resume` — multi-session resume. `--continue` is the more recognized variant.
- `/clear` — preceded by `/compact` in every workflow doc; ship the more impressive one.
- `/diff`, `/rewind`, `/security-review` — known to power users but not first-glance recognizable.
- aider `/add <file>` — important inside aider but doesn't read as AI from a screenshot.
- Cursor `Ctrl+K`, `Ctrl+L` — Cursor is GUI-only; doesn't fit a wallpaper-cells frame.
- `gh copilot alias` — config-time, not daily.

## Recommended hero PNG layout (one frame, ~10 cells visible)

```
┌────────────────────────────────────────────────────────────────────────┐
│ A         B                          C                     D           │
├────────────────────────────────────────────────────────────────────────┤
│ AI                                                                      │
│ r: claude              r: claude --continue        i: claude --version  │  ← Row 2
│ r: claude /init        r: claude /review           /compact at 70%      │  ← Row 3
│ r: gh copilot suggest  r: gh copilot explain       /fix /tests /explain │  ← Row 4
│ r: aider --ask "..."   ext: cemheren/quicksheet-copilot-ext             │  ← Row 5
│                                                                         │
│ TOKENS THIS WEEK                                                        │
│ s: 12,8,15,22,9,7,14   $/day: i: claude /cost      Σ: 87,341 tok        │  ← Row 7
│                                                                         │
│ TODAY                                                                   │
│ - refactor parser      https://github.com/me/proj/pull/42 (hyperlink)   │  ← Row 9
│ - r: code .            r: gh pr list --json number,title                │  ← Row 10
└────────────────────────────────────────────────────────────────────────┘
```

What this layout shows in one screenshot:
- **3 cell prefixes visible** (`r:`, `i:`, `s:`, `ext:`, hyperlink) — answers "what is this thing?"
- **AI-tool literacy** — every command above is recognizable to the target audience.
- **A real cost-tracking section** — `s:` sparkline + `/cost` cell + sum. Solves a real pain (token spend).
- **The copilot extension cell** — visible proof QuickSheet integrates AI directly, not just launches it.
- **Today list with hyperlinks + run cells** — shows the wallpaper is a working surface, not eye candy.

## Hard rules for the screenshot

- **Real data only.** Don't fake token counts. Run `claude /cost` once, put the real number in.
- **No editor visible.** No VS Code window, no Cursor — keep wallpaper-pure. The point is that QuickSheet *is* the surface.
- **Resolution ≥1920×1080.** Crops well for every subreddit's image limits.
- **Theme: Nord or Dracula.** Both read as "developer aesthetic" instantly. HotdogStand is a gag; Light theme washes out on r/unixporn.
- **Anti-aliasing on.** Default `--desktop` mode is fine. Verify font is Consolas/JetBrains Mono — not the OS default.
- **No personal handles in PATH or window titles.** Crop or alias.

## Implications for QuickSheet — actions to queue

1. **Hero PNG capture.** User produces this, saves to `docs/img/ai-hero.png`. Skill then PRs the README hero swap.
2. **Pre-seed CSV.** Skill will ship `examples/ai-workflow.csv` matching the layout above so users can `dotnet run -- examples/ai-workflow.csv` and reproduce the screenshot exactly. One-line "see exactly the screenshot" install line — high copy-paste value, matches the dev-onboarding pattern Aider's docs use.
3. **r/MachineLearning + r/LocalLLaMA variant.** Once the AI-hero PNG exists, write a venue draft framing QuickSheet as the "ambient AI cost monitor" — lateral audience the current drafts don't address.

Sources:
- [Claude Code Commands Cheat Sheet (2026)](https://www.scriptbyai.com/claude-code-commands-cheat-sheet/)
- [Claude Code Cheat Sheet — Every Command (Claude Directory)](https://www.claudedirectory.org/blog/claude-code-cheat-sheet)
- [Essential Claude Code Slash Commands — BSWEN](https://docs.bswen.com/blog/2026-05-13-claude-code-slash-commands-guide/)
- [50 Claude Code Tips and Best Practices — builder.io](https://www.builder.io/blog/claude-code-tips-best-practices)
- [Reduce Claude Code Costs 60% — systemprompt.io](https://systemprompt.io/guides/claude-code-cost-optimisation)
- [Claude Code --continue and --resume](https://pasqualepillitteri.it/en/news/366/claude-code-continue-resume-guide)
- [GitHub Copilot CLI command reference](https://docs.github.com/en/copilot/reference/copilot-cli-reference/cli-command-reference)
- [GitHub Copilot CLI (github/gh-copilot)](https://github.com/github/gh-copilot)
- [Top 10 Copilot Chat slash commands in VS Code](https://medium.com/@shrinivassab/top-10-github-copilot-slash-commands-every-vs-code-developer-must-know-in-2025-4f866360fdad)
- [aider in-chat commands](https://aider.chat/docs/usage/commands.html)
- [aider tips — /ask + /code workflow](https://aider.chat/docs/usage/tips.html)
