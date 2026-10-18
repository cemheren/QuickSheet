# Product Hunt submission — draft (2026-06-11)

> User submits manually at producthunt.com. Do not post.

## Pre-launch checklist

- [ ] Demo GIF or 30-second video walkthrough (HUMAN BLOCKER — skill cannot capture)
- [ ] At least one GitHub release tagged (v1.0.0+)
- [ ] README screenshot shows a compelling filled-in desktop
- [ ] Have 2–3 early supporters ready to upvote + leave genuine comments in the first hour

## Product name

**QuickSheet**

## Tagline (60 chars max)

```
Your desktop wallpaper is now an interactive spreadsheet
```

Alternative taglines (pick the one that feels best on the day):

```
Turn dead desktop space into a live CSV grid
```

```
A spreadsheet that lives behind every window
```

## Description (Product Hunt "About" field, ~260 words)

```
QuickSheet replaces your static wallpaper with a transparent, interactive spreadsheet grid. Click anywhere on the desktop to jot a note, paste a URL to make it a clickable link, prefix a cell with `r: code .` to make it a launcher. Everything autosaves to a plain CSV file every 5 seconds.

It runs on Windows (embeds behind desktop icons via the WorkerW trick) and Linux (X11, using _NET_WM_WINDOW_TYPE_DESKTOP). Cross-platform, same CSV file, same shortcuts.

Key features:
• Always-on scratchpad — no window to find, no app to open.
• App launcher — prefix cells with `r:` and hit Enter. Multi-select to launch your morning stack.
• Link dashboard — paste URLs, they're highlighted and open in your browser.
• Live data — column sums (Σ), sparklines (`s:`), inline subprocess output (`i: htop`).
• 69+ extensions — weather, stocks, RSS, GitHub streak, system monitoring, and more. Install with one cell: `ext: github:user/repo`.
• Zero dependencies — clone → `dotnet build` → run. No NuGet packages, no npm, no Docker.

Built with .NET 9, MIT licensed, entirely P/Invoke for native interop. The CSV file round-trips through Excel, vim, Miller, or any other tool.

If you spend your day in a terminal or IDE and want your desktop to DO something, QuickSheet is for you.
```

## Maker's comment (first comment after launch)

```
Hey PH! I'm the maker of QuickSheet.

I built this because I have a second monitor that shows a wallpaper doing nothing. I wanted that space to be useful — quick notes, links I open every day, a few shell commands I run to start my workflow.

The hard constraint I set: zero external dependencies. All the Windows and Linux native interop is hand-written P/Invoke. The data format is plain CSV, so it works with every tool in the unix pipeline ecosystem.

A few things I'm proud of:
- The extension system: one cell (`ext: github:user/repo`) installs a JSON-lines extension. There are 69+ community extensions now.
- Inline subprocess output: type `i: tail -f /var/log/syslog` and the cell streams live output.
- It genuinely replaced my sticky-notes workflow. I haven't opened a separate note app in months.

Would love feedback on what you'd want from a desktop-as-spreadsheet. What cells would you fill?
```

## Topics / Categories

Primary: **Developer Tools**
Secondary: **Productivity**, **Open Source**

## Thumbnail / Gallery guidance (for human)

1. **Hero image:** Full-desktop screenshot showing QuickSheet with a mix of notes, runnable commands (green `r:` cells), URLs (blue underline), and a sparkline. Dark theme, clean font.
2. **GIF (if possible):** 10-second loop: click desktop → type a note → it autosaves → type `r: code .` → hit Enter → VS Code opens.
3. **Screenshot 2:** Extension in action — weather, stock ticker, or GitHub streak cell updating live.
4. **Screenshot 3:** Side-by-side Windows + Linux showing the same CSV file.

## Launch timing

- **Day:** Tuesday or Wednesday (Mon–Thu are strongest on PH; Tue/Wed have slightly less competition than Mon).
- **Time:** 12:01 AM Pacific (PH resets at midnight PT; early listing = more total upvote hours).
- **Avoid:** Don't launch same day as a major Apple/Google event or a YC Demo Day — the front page gets crowded.

## Amplification plan

Within first hour of going live:
1. Share PH link on Twitter/X, Mastodon, Bluesky (use the pre-drafted threads in `drafts/twitter-thread.md` and `drafts/mastodon-bluesky.md`).
2. Drop the PH link in the existing Reddit posts if already live, or post with `[Product Hunt]` note.
3. Update the GitHub README with a "Featured on Product Hunt" badge after launch day.
4. Email early supporters (see `drafts/email-pitches.md` list).

## Expected outcome

- Product Hunt is strongest for polished dev tools with a visual hook. QuickSheet's "wallpaper as spreadsheet" IS visual — that's the differentiator.
- Realistic: 50–200 upvotes for a novel open-source tool, 30–100 GitHub stars from PH traffic directly.
- Long-tail: PH listing is permanent, shows up in Google searches for "desktop spreadsheet" and "wallpaper productivity tool."

## Why Product Hunt specifically

- The visual hook ("your wallpaper is a spreadsheet") is inherently demo-able — it's the kind of thing PH voters like to upvote because it looks cool in a screenshot.
- Open-source dev tools do well on PH when they're genuinely different (not another TODO app or note-taker).
- The comparison is Rainmeter/Stardock Fences but with an *editable data grid* — that's novel enough to generate "I didn't know you could do that" reactions.
- Free, MIT-licensed, no signup, no SaaS — PH voters appreciate that.

## Differentiation talking points (for comment replies)

- vs **Rainmeter/Conky:** Those render pre-built widgets; QuickSheet renders editable cells you type into. Different primitive.
- vs **Notion/Obsidian:** Those are apps you switch to. QuickSheet is always visible behind your windows — zero context-switch cost.
- vs **sc-im/VisiData:** Those are terminal-only viewers. QuickSheet does that too, but the wallpaper mode is the unique angle.
- vs **Excel/Google Sheets:** Those are full-fat spreadsheets. QuickSheet is intentionally minimal — CSV, plain text, shell integration.
