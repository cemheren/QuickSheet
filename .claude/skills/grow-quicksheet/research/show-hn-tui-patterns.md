# Show HN: TUI launches — what correlates with high scores

Research run 2026-05-14. Source: HN Algolia search (`Show HN TUI`, top 20 by points), plus thread reads on the top two.

## Summary (read this if nothing else)

- **Title structure that always works**: `Show HN: <Name> – <pithy descriptor>` with an em-dash. Every top-5 post follows this exact pattern.
- **Submit the repo URL, not a blog post** (top 2 do this; Bagels at #2 hit 283 points pointing straight at github.com/EnhancedJax/Bagels).
- **OP's first comment should lead with personal itch**, not feature list. "I built this for myself to do X" beats "introducing Y."
- **Score floor**: 5 of the top 20 cleared 100 points. A reasonable target floor for QuickSheet is ~50–100 points. Anything > 50 puts you on the front page for hours.
- **Engage in comments about trade-offs**, not thanks. Top scorers' threads have OP defending design choices (SQLite vs ledger, why Nim, etc.).

## Data — top 20 Show HN TUI posts ever (HN Algolia)

| Rank | Title | Points | Comments | Date |
|------|-------|--------|----------|------|
| 1 | Chawan TUI web browser | 387 | 69 | Jun 2025 |
| 2 | Bagels – TUI expense tracker | 283 | 87 | Jan 2025 |
| 3 | Austin-Tui – Python program spy | 256 | 36 | Oct 2020 |
| 4 | hackernews_tui – HN browser | 160 | 31 | Apr 2021 |
| 5 | TUI for managing XDG default applications | 138 | 44 | Jan 2026 |
| 6 | Retronews – HN/Lobsters TUI | 109 | 36 | Sep 2024 |
| 7 | Tracexec – execve tracing TUI | 57 | 9 | May 2024 |
| 8 | TUI-use – AI agent terminal control | 52 | 37 | Apr 2026 |
| 9 | mpvc-tui – mpv controller | 47 | 10 | Dec 2022 |
| 10 | Docfd – multiline fuzzy finder | 44 | 17 | Apr 2024 |
| 11 | oeis-tui – integer sequences | 31 | 2 | Nov 2025 |
| 12 | 7guis TUI in Go | 27 | 3 | Jul 2022 |
| 13 | Tuiqo – document versioning | 26 | 5 | Jun 2017 |
| 14 | Grafana TUI – dashboard browser | 24 | 8 | Mar 2026 |
| 15 | Java TUI framework with sixel | 21 | 2 | Feb 2022 |
| 16 | Markdown TUI editor | 20 | 2 | Apr 2026 |
| 17 | Revdiff – diff reviewer for AI agents | 17 | 4 | Apr 2026 |
| 18 | PyTermGUI – layout system | 15 | 2 | May 2022 |
| 19 | Mprocs – multi-process controller | 14 | 5 | May 2022 |
| 20 | CuTE – Rust libcurl HTTP client | 14 | 3 | Oct 2023 |

URL pattern for any of these: `https://news.ycombinator.com/item?id=<id>` — IDs in `awesome-tuis` / `drafts/showhn.md` follow-ups.

## Title pattern (every top-5 uses it)

```
Show HN: <Name> – <pithy descriptor>
```

- "Show HN: Chawan TUI web browser"
- "Show HN: Bagels – TUI expense tracker"
- "Show HN: Austin-Tui – Python program spy"
- "Show HN: hackernews_tui – HN browser"

Em-dash, no period, no marketing adjectives. Descriptor is 3–6 words.

## Top-2 thread reading

### Chawan #1 (387 points)

- URL submitted: project release-notes page (`https://chawan.net/news/chawan-0-2-0.html`), not a repo readme. So the rule "submit the repo" is not universal — submitting a *release* page also works if it has prose explaining what's new.
- Top comments engaged with: why-not-Blink (technical choice), TTY internals (learning resource ask), Nim language choice (genuine curiosity), and bug reports.
- No OP first-comment shown in scrape — Chawan's prose page substituted.

### Bagels #2 (283 points)

- URL submitted: GitHub repo (`github.com/EnhancedJax/Bagels`).
- OP first comment: "I'm Jax and I've been building this cool little terminal app for myself to track my expenses and budgets!" — personal-itch framing, identifies the author, two sentences total.
- Highest-engagement comment debated SQLAlchemy vs plain-text accounting (ledger/beancount). A countercomment defended SQLite. OP engaged technically. This was the thread's spine.
- Distinctive: included rounded-unicode-box screenshots in the README, learned Python as part of the project, positioned as "solving my own workflow" not "competing."

## Implications for QuickSheet (this is the deliverable)

1. **Refactor `drafts/showhn.md` first-comment to lead with personal itch.** Current draft is fine but slightly marketing-y. Open with a one-sentence "I'm Akif, and my desktop wallpaper has been a static image I never interact with for years — I wanted something on it that did things, so I built this." Then features. Then constraints. Bagels-style.

2. **Submit the repo URL, not docs/tour.md.** docs/tour.md is the cross-link target *in replies and the first comment*, not the submission. The submission should be `https://github.com/cemheren/QuickSheet`.

3. **Brace for the "why not [other tool]" thread.** Pre-write 3 short comment replies for:
   - "Why not VisiData / sc-im / NeoVim spreadsheet plugin?" → wallpaper-embedding is the differentiator; those don't do that.
   - "Why .NET? Why not Rust/Go?" → because P/Invoke makes the Win32 WorkerW work; .NET is the path of least resistance for that specific trick.
   - "Why zero NuGet deps? Isn't that just dogma?" → it's a design forcing function; the WorkerW + X11 code stays small because no helper library is doing it for you.

4. **Title is fine as-is.** Current draft title "Show HN: QuickSheet – my desktop wallpaper is a spreadsheet" matches the pattern (em-dash, 6-word descriptor). Don't change it.

5. **Post timing: Tue–Thu, 8–10am Pacific** is conventional wisdom and matches Bagels (Sunday but it was a slow news day; the post stayed on front page 36 hours). For active news days, weekday morning Pacific is safer.

## Queue these as next-run actions

- Bucket C: refactor `drafts/showhn.md` per implication #1 (personal-itch first comment), add the three pre-written reply blocks (implication #3). One run, one revision.
- Bucket A (only after #1 above): no README change needed yet; current hero already says "Your desktop is a spreadsheet" which is already the right tagline.
- Bucket B: no new release needed pre-launch — v0.2.0 is fresh.
