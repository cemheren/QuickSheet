# Niche communities: where QuickSheet specifically fits

Research run 2026-05-14. Source: r/unixporn AutoMod source (github.com/unixporn/upmo), terminaltrove.com about page, console.dev selection-criteria page. Reddit blocks scrape so r/commandline/r/i3wm gathered indirectly.

## Summary

- **r/unixporn is the underexploited highest-fit channel.** QuickSheet's `--desktop` mode is literally what the sub is about — a desktop screenshot with custom stuff on it. Different framing than the dev-tool subs.
- **terminaltrove.com submission is just an email** to `hello@terminaltrove.com`. Free shot, no risk.
- **console.dev is editorial** (Thursdays, 2-3 tools/week). QuickSheet ticks every criterion they list. Worth the pitch email.
- All three are "send and forget" — no thread to babysit, no posting-cadence stress.

## Data

### r/unixporn (~600k subs)

AutoMod (via [u/upmo](https://github.com/unixporn/upmo/blob/master/readme.md), now AutoModerator) enforces:

- **Title must tag a desktop environment (DE).** Posts without a DE tag get removed.
- **No deprecated tags** (e.g. old `[i3]` style retired in favor of approved tags).
- **No OS tags** (`[Linux]` etc. get removed — be specific about DE/WM).
- **Image must be on an approved host** (imgur, etc.).
- **A details-comment is mandatory within 30 minutes** of submission — distro, DE/WM, font, theme, dotfiles repo, etc. Bot warns at 15min, removes at 30min.
- **Minimum karma requirement** to post.

Implication: QuickSheet's wallpaper mode IS the screenshot — the post should be styled as a desktop screenshot ("rice"), NOT as a project announcement. The "project" framing goes in the details-comment.

### terminaltrove.com

[About page](https://terminaltrove.com/about/): two submission paths.

1. "Post a Tool" button (form, exact fields not public until you click).
2. Email `hello@terminaltrove.com`.

No documented quality bar. They've published a "Tool of the Week" feature, plus a "Newly Added" carousel. Newly Added examples include `cheznav` ("A TUI for managing chezmoi dotfiles") — single-purpose TUIs with one-line descriptions get listed.

### console.dev (weekly newsletter, ~Thursday)

[Selection criteria](https://console.dev/selection-criteria):

- Subjective-yet-objective. They feature 2-3 tools per weekly issue.
- Must-have signals they explicitly check:
  - "Is this interesting and useful to developers?"
  - "Is the primary user a developer?"
  - "Is there a self-service signup?" (open-source = yes by default).
  - "Would this form part of a regular-use set of developer tools?"
  - "Does it make me a better developer?" (faster builds, better code quality, etc.).
  - "Is the tool high quality?" / "Actively maintained?"
  - "Does it have good documentation?" / "Is it fast?"
  - "Does it work on multiple platforms?"
  - Advanced-user nods: dark mode, API, CLI, keyboard shortcuts.

QuickSheet score against these (honest read):
- Developer audience: ✓ (TUI users skew dev)
- Self-service: ✓ (clone, build, run)
- Regular-use: ✓ (wallpaper mode = literally always-on)
- Makes-me-better: ~ (saves context-switches, debatable)
- Quality / actively maintained: ✓ (recent commits, v0.2.0 tagged, CHANGELOG, CONTRIBUTING, SECURITY)
- Documentation: ✓ (README + tour + recipes + extensions docs)
- Fast: ✓ (no boot, no DB)
- Cross-platform: ✓ (Windows + Linux)
- Advanced-user nods: ✓ (CLI, keyboard, headless `--export-md`)

QuickSheet fits 8/9 of console.dev's checklist. Worth the pitch.

## Implications for QuickSheet (the deliverable)

1. **Draft a r/unixporn submission as a rice post, not a project announcement.** Distinct framing from the existing `drafts/reddit-*.md` files. Title pattern: `[<DE>] something interactive on my wallpaper`. Image: wallpaper-mode screenshot of QuickSheet in actual use (real cells, real data, real desktop). Details-comment template: distro, DE/WM, terminal emulator, font, theme, dotfiles repo, **and** "the grid you're seeing is QuickSheet, my project, [link]". Bury the project mention near the bottom of the details-comment — leading with it gets read as self-promo and downranked.

2. **Draft a one-paragraph terminaltrove pitch email.** Just text, one line description, one line "why it's different", repo link. Single-shot send to `hello@terminaltrove.com` — no need for a polished form submission first.

3. **Draft a console.dev pitch email.** Lead with the 8/9-criteria-match summary, repo link, ask for nothing specific (no "please feature us"). Their criteria page implies they self-evaluate when nudged.

## Queue these in the log

- Bucket D: write `drafts/unixporn-rice.md` — rice-style submission template + details-comment template, with the r/unixporn-specific framing rules.
- Bucket D: write `drafts/email-pitches.md` — short pitch emails for terminaltrove + console.dev, each ≤120 words, ready to send.

Both pass the "plausibly causes stars" filter: r/unixporn has the highest demographic fit found so far, and console.dev gets QuickSheet directly into a curated dev newsletter inbox.
