# What features actually get loved — inverse teardown of adjacent TUI tools (2026-05-17)

> Inverse of `adjacent-pitch-teardown.md`. That brief asked "what's their hero pitch."
> This one asks "what *single feature* gets repeatedly mentioned in praise of these
> tools — what's the thing readers actually click for?"

## TL;DR (5 lines)

- **Lazygit's win is keymap speed, not the Git wrapper.** "Context-based keybindings eliminate the steep curve" recurs across 4 top HN threads. The thing that's loved is the *reduction in context-switching*, not the feature surface.
- **VisiData's win is the "vi for X" metaphor + a lightning demo video.** Show HN at 221 pts uses that exact title; a follow-up "VisiData lightning demo at PyCascades" video alone hit 195 pts.
- **Harlequin's win is "drop-in replacement for X."** Even sharper than "alternative to X." Frames itself as a swap-in for the DuckDB CLI.
- All three live in the *speed-of-iteration* category, not the *richness-of-feature* category. None of them advertise "300 things you can do" — they advertise "the thing you already do, in half the keystrokes."
- For QuickSheet: lean into **cell-prefix keystroke speed** in the pitch (lazygit's lesson) and a **"X for CSV-first folks"** framing (VisiData's). Demo video is unavailable to the skill but is queued for human capture.

## Data — top HN posts (HN Algolia, 2026-05-17)

| Tool | Top HN post title | Points | Comments | URL |
|------|-------------------|--------|----------|-----|
| Lazygit | The lazy Git UI you didn't know you need | 436 | 222 | bwplotka.dev/2025/lazygit |
| Lazygit | Lazygit: A simple terminal UI for Git commands | 395 | 141 | github.com/jesseduffield/lazygit |
| Lazygit | Lazygit: Simple terminal UI for Git commands | 239 | 79 | (resubmit) |
| Lazygit | Lazygit Turns 5: Musings on Git, TUIs, and open source | 211 | 69 | jesseduffield.com |
| VisiData | Show HN: VisiData – vi for data | 221 | 23 | github.com/saulpw/visidata |
| VisiData | Visidata 2.0 | 203 | 22 | visidata.org/blog |
| VisiData | VisiData Lightning Demo at PyCascades 2018 [video] | 195 | 39 | youtube.com |
| VisiData | VisiData in 60 Seconds | 112 | 15 | jsvine.github.io |
| VisiData | VisiData 1.0 released | 84 | 3 | visidata.org/releases |
| Harlequin | Harlequin: SQL IDE for Your Terminal | 183 | 85 | github.com/tconbeer/harlequin |

## Patterns across all three tools

### 1. The loved feature is *iteration speed*, not breadth

Every comment-thread theme that recurs across the three tools reduces to "I open it less / do the thing in fewer keystrokes / don't have to leave the terminal." None of them are loved for being "complete" or "powerful."

Implication for QuickSheet: **Stop emphasising 40+ extensions in the hero.** It signals breadth; readers want speed. The cell-prefix list (`r:`, `i:`, `s:`, etc.) is the speed angle; the extension count is the breadth angle. Lead with prefixes.

### 2. The "drop-in replacement for X" framing is sharper than "alternative to"

Harlequin's "drop-in replacement for the DuckDB CLI" is a stronger commit than "an alternative to the DuckDB CLI." It promises the reader *they don't have to learn a new mental model* — same workflow, better surface.

Implication for QuickSheet: "the data is just CSV" already implements this (drop-in for any CSV workflow). The `--export-md` flag (`csv → markdown table` in one command, no UI) also implements it for the csvkit-style audience. **Surface `--export-md` more prominently** — it's the most "drop-in replacement"-shaped feature in the project.

### 3. Demo video correlates with reach

VisiData's lightning demo (195 pts) is *separately scored from* its Show HN (221 pts). Lazygit's most-recent top post (436) is a blog walkthrough with embedded GIFs. Reach compounds: demo → blog → resubmit → top of front page.

Implication for QuickSheet: the skill cannot create a video. Queue it as a human-required action. Format that worked: short (60s), no narration, one continuous shot showing keystrokes → cell behaviour → screenshot of the wallpaper-mode result.

## Implications (queue these)

### 1. Bucket A: README hero re-edit (later run, not now)

The 13-word add-line from PR #108 (CSV / runnable cells / cross-platform) is good. The current "40+ extensions" badge is *the wrong selling point* — peers don't sell on breadth. Consider replacing or de-emphasising the extensions badge in favour of one that names the keystroke-speed angle (`cell prefixes`, `one-keystroke launchers`). The `cell_prefixes-6` badge added in PR #108 already does this; **remove or downplay the `extensions-40+` badge** to amplify the prefix-speed framing.

This is contentious — extensions are real flywheel content. Don't act on it without checking with the user. **File this as a question for the user, not a unilateral edit.**

### 2. Bucket A: surface `--export-md` more

`--export-md` is the most "drop-in replacement" feature in the project (any csvkit/Miller user would understand it instantly). Currently it gets one line in README "Quick Start" and a row in `--help`. Could merit:
- A small `docs/csvkit-comparison.md` page (one paragraph + one example).
- A separate paragraph in README under a "Drop into your existing CSV workflow" subhead.

### 3. Queue (human-required): demo video

Format target: 60 seconds, no narration, one continuous shot. Open QuickSheet → type `r: code .` → Enter (IDE opens) → cell shows `s: 1,2,3,4,5,6` rendering live → window comes up over the wallpaper → screenshot the result. Upload to YouTube or as a `.webm` linked from README. Worth ~50-200 stars based on VisiData precedent.

### 4. Bucket C: re-revise `drafts/showhn.md` (later run, not now)

The current draft leads with personal-itch (Bagels pattern). The lazygit/VisiData lesson: also lean into the **iteration-speed** angle in the body, not just the protocol angle. Concrete: a paragraph that lists "the things I open less often now" — Trello (replaced by a column of `r: code .` cells), htop (replaced by `sysmon:`), Slack (replaced by, nothing, still need Slack), browser tabs for status pages (replaced by `apistatus:`, `tls:`, `health:`).

Queue this as a 3rd-pass revision of `drafts/showhn.md`. Don't do it this run (last C was the unixporn-rice revision; need at least one non-C run between).
