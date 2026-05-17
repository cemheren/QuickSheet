# r/unixporn rice-style submission — QuickSheet

Drafted 2026-05-14; revised 2026-05-17 per `research/wallpaper-launch-venues.md`. **Distinct from `drafts/reddit-commandline.md` etc. — that's a project announcement; this is a rice post.**

> **Position in launch sequence:** This is now the **#1 launch venue** (per the venue brief — HN underperforms hard for the wallpaper-tool genre). Submit this *before* Show HN / Lobsters / Reddit-commandline. The screenshot does 80% of the work.

## Framing rule

r/unixporn is not a project-announcement sub. It is a desktop-screenshot sub. The wallpaper-mode screenshot of QuickSheet IS the post; the project mention belongs near the bottom of the mandatory details-comment, not in the title.

## Title

Per AutoMod, title must tag a desktop environment (DE). Pattern:

```
[<DE>] <vibe-y descriptor>
```

Examples (pick one matching your actual desktop):

1. `[GNOME] my wallpaper is a spreadsheet now`
2. `[KDE] interactive grid in place of the wallpaper`
3. `[Xfce] live cells on the desktop`
4. `[i3] tiling, but the wallpaper is a TUI`
5. `[Hyprland] interactive grid as the desktop` *(blocked today — Wayland not yet supported. r/unixporn 2024–26 is Hyprland-dominated, so this is the biggest missed-audience tag. Workaround: capture the screenshot from a separate X11 session on the same machine (Xorg login, or a TTY-launched Xfce/i3) and post the [DE] tag matching THAT session. Don't claim Hyprland if it's not.)*

Recommended: **#1 or #2** depending on which DE you actually use. If neither, swap the bracketed tag for whichever DE you screenshotted on.

## Image

Single screenshot, full-resolution, uploaded to imgur (or another r/unixporn approved host).

Composition to aim for:
- Multiple visible windows over the wallpaper grid (one editor, one browser) — proves it's actually behind everything.
- Grid showing **real, plausible content**: a few labelled launchers, a couple `i: ` cells with live output, a `s: …` sparkline, maybe a `wthr:` weather widget. NOT a demo/lorem ipsum grid.
- Σ in the status bar visible.

Avoid: cropped screenshots, dark wallpaper that hides the grid, no other windows (looks like a regular TUI shot).

## Details-comment template (MANDATORY within 30 minutes of post)

Standard rice details first, project mention LAST. The bot wants distro/DE/font/etc.; readers downvote if you skip those for marketing.

```
- Distro: <e.g. Arch / Fedora 41 / Ubuntu 24.04>
- DE: <GNOME 47 / KDE 6.x / Xfce 4.x — pick yours>
- WM: <Mutter / KWin / Xfwm4 / i3 / etc.>
- Terminal: <Alacritty / Kitty / Wezterm / GNOME Terminal>
- Font: <e.g. JetBrains Mono 11>
- Theme: <GTK theme / icon theme — what you actually use>
- Wallpaper (the grid): QuickSheet — open-source side project of mine. Cell prefixes for runnable commands, live subprocess output, sparklines, and extensions installed by git URL. Embeds via Win32 WorkerW on Windows and `_NET_WM_WINDOW_TYPE_DESKTOP` on X11. Zero NuGet deps, MIT. → https://github.com/cemheren/QuickSheet
- Dotfiles: <your repo URL or "not public yet">
```

Notes:
- Put the QuickSheet bullet **last** in the list, framed as the wallpaper.
- Don't say "I made this" — say "open-source side project of mine." Different tone.
- If asked in the comments "what's that grid thing", the answer's already linked; don't double-down on self-promo in replies. Just answer the question.

## What to do in the comments

- Answer questions about the rice elements (font, theme, dotfiles) first.
- If someone asks about QuickSheet specifically, answer with a feature, not a pitch. ("Yeah, `r: code .` cells launch on Enter. There's a list of prefixes in docs/tour.md.")
- Don't post the same screenshot to r/commandline / r/dotnet within the same week — that reads as crossposting for stars.

## Timing

- Post Saturday or Sunday morning Pacific. Weekend ricing crowd skews active.
- Avoid: same day as a big rice contest result, same day a major DE/WM hits front page (your post gets buried under the hype).

## Pitfalls

- Including a video/GIF instead of a still image — most rice subs prefer stills for the first post; you can link a GIF in the details-comment.
- Title that says "QuickSheet" — reads as self-promo. The screenshot title is about the *desktop*, not the *project*.
- Title that says "Show HN" or "I built" — wrong sub culture.
- Forgetting the details-comment — bot will warn at 15min and remove at 30min.
- Posting without the minimum karma — check your account's karma in the sub before submitting.

## What success looks like

For r/unixporn specifically:
- 500+ upvotes within 24h is a strong day. 100+ is fine.
- 30+ comments with rice-element questions = the screenshot landed.
- Star delta: indirect but consistent. Rice screenshots with project link in details get a clear "this looks cool, what's the project?" → click → star flow. Typically a single decent post = 10-50 stars.

## Persona-variant screenshots (pick one — the highest-virality is the student rice)

The persona research (`research/personas/`) identified four rice-shaped wallpaper screenshots, in expected-virality order:

### 1. CS-student rice — highest viral candidate

Identified in `research/personas/students.md` as the single highest-virality opportunity across the 10-persona slate. The student rice combines (a) rice culture, (b) student-flex identity, (c) free + open-source.

Grid contents to compose:
- Today's class strip ("Calc II 10:00 · CS101 14:00") with current highlighted via `c:yellow:`.
- GitHub commit-today cell (`gh:` ext — currently scaffolded, not yet a public repo).
- LeetCode streak cell (`leetcode:` ext — not yet scaffolded; user can fake this with a static cell for the screenshot).
- GPA / per-class grade column.
- Pomodoro session count.

Theme suggestion: **Gruvbox** or **Nord** (already shipped Ctrl+T themes). Catppuccin works too — needs adding to theme presets first.

### 2. TTRPG GM-screen rice — second viral candidate

Identified in `research/personas/gamedev-ttrpg.md`. The GM-side monitor is never seen by players, always glanceable to the GM — a uniquely clean wallpaper use case.

Grid contents:
- Initiative tracker column (current turn highlighted).
- NPC HP grid.
- Quest log column.
- Loot table.
- A `roll: 1d20` cell ready to fire (the `roll:` extension is scaffolded but not yet a public repo).

Theme suggestion: **Dracula** (already shipped) or any dark theme.

### 3. Trader multi-monitor rice

`research/personas/traders.md`. r/battlestations is a closer fit than r/unixporn for the multi-monitor angle, but a single-monitor trader rice still works on r/unixporn.

Grid contents:
- Watchlist column (`stock:`, `price:` extensions — already shipped).
- Today's P/L mega-cell.
- News strip (`news:` ext — already shipped).
- FX side-row (`fx:` ext — already shipped).

### 4. Homelab rice

`research/personas/homelab.md`. Lower r/unixporn fit than r/selfhosted, but works as crossover.

Grid contents:
- Service-health row (`health:` ext — currently scaffolded).
- Pi-hole stats (`pihole:` ext — currently a PR per #109).
- Disk usage strip + cert expiry strip.

### Choosing one for the actual post

Pick whichever the user can authentically screenshot from their own machine. The persona narratives are scaffolding — what matters is that the screenshot looks lived-in, not staged. Faking with placeholder cells reads as marketing and gets called out in comments.
