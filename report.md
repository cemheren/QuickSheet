# QuickSheet Audience Research

**Date:** 2026-04-29
**Method:** WebSearch + WebFetch + Puppeteer MCP across reddit, Hacker News, dev blogs, productivity forums
**Goal:** find communities where people are frustrated by something QuickSheet plausibly solves

---

## Executive Summary

QuickSheet sits at an unusual intersection: **always-visible desktop surface** + **spreadsheet UX** + **runnable cells** + **zero dependencies**. Each axis has its own audience, and the overlap is small but motivated. The strongest signal across every search axis was the same complaint: **"my notes/tasks are out of sight, so they stay out of mind"** — and the fix users keep gravitating to is *pinning something to the desktop itself*, not opening another app.

Five distinct audience segments emerged. The strongest fit for early adopters is the **Rainmeter / desktop-dashboard crowd** (proven willingness to install hacky desktop tooling, already accept config-as-text) and the **HN terminal-spreadsheet crowd** (`sc-im` / `VisiData`) — the latter explicitly wants "vim for spreadsheets" and care about supply-chain hygiene, which maps exactly to QuickSheet's positioning.

The largest gap between QuickSheet's pitch and what users currently complain about is **sync** — every persistent-note alternative thread mentions cross-device sync as table stakes. QuickSheet's CSV-only, local-only model is a feature for the security-minded but a deal-breaker for the mass productivity audience. Lean into the former, don't fight the latter.

---

## Segment 1 — Desktop Dashboard / Anti-Default-Wallpaper Crowd

These people already replace their wallpaper with information. Rainmeter (Windows) and Conky (Linux) are the established tools. They write/install skins to surface clocks, weather, system stats, todo lists, notes — the exact surface area QuickSheet provides, minus the spreadsheet metaphor.

**Direct evidence (quote, MakeUseOf author building a Rainmeter dashboard):**

> "Every morning when I start my computer, I immediately see everything important: time and date, system performance to avoid slowdowns, Bluetooth, volume adjustment, and a quick note for reminders."
> — *I built my own desktop dashboard with Rainmeter…* ([article](https://www.makeuseof.com/how-to-build-custom-desktop-dashboard-rainmeter/))

> "Technology should adapt to us, not the other way around. My desktop finally works the way I think, displaying the information I need when I need it."
> — same author

**Sub-evidence — Win11 widgets resentment (large secondary audience):**

> "Widgets now only take up a third of the available space, with the remaining two-thirds taken up by 'My feed' … mostly junk … random celebrity gossip, clickbait, auto-playing videos."
> — *Microsoft Already Ruined Windows 11 Widgets* ([howtogeek](https://www.howtogeek.com/microsoft-already-ruined-windows-11-widgets/))

> "The inability to pin widgets to the desktop — one of the most-requested features — remains a glaring omission."
> — *Top Third-Party Widgets for Windows 11* ([WindowsForum](https://windowsforum.com/threads/top-third-party-widgets-for-customizing-your-windows-11-desktop-in-2025.368171/))

**Maps to QuickSheet:** core wallpaper-as-spreadsheet idea, autosave, lightweight footprint.

**Existing tooling shows demand is real (these all exist because users built them):**
- [alperenozlu/rainmeter-todo](https://github.com/alperenozlu/rainmeter-todo) — "to-do skin that will always be in the spotlight"
- [Pernickety/rainmeter-todo-list](https://github.com/Pernickety/rainmeter-todo-list)
- [Rainmeter forum: Project "Desktop Notes"](https://forum.rainmeter.net/viewtopic.php?t=36060)
- [Rainmeter forum: Skin for directly input notes](https://forum.rainmeter.net/viewtopic.php?t=15850)
- [Conky as a Todo list (Arch Forums)](https://bbs.archlinux.org/viewtopic.php?id=74705)
- [MOSAID: Sticky Todo Reminders with Cron, systemd & Conky](https://mosaid.xyz/articles/from-terminal-to-desktop--sticky-todo-reminders-with-cron,-systemd-&-conky-14.html)
- [Mabox 26.03 — Conky + new TODOlist tool](https://maboxlinux.org/mabox-26-03-conky-improvements-todolist-and-keyboard-shortcuts/)

**Where to engage:** r/Rainmeter, r/conky, r/unixporn, r/Windows11, r/desktops, Rainmeter forum.

---

## Segment 2 — Heavy-Notes-App Refugees (Notion / Obsidian / Sticky Notes burnout)

People who tried "the right way" to take notes and bounced off either the setup overhead (Obsidian) or the bloat/cloud dependency (Notion / Sticky Notes' OneDrive sync).

**Direct evidence — Obsidian fatigue:**

> "Obsidian is great, but it also seems like a tool that asks for a lot upfront … users get stuck in setup mode before doing any actual work."
> — *Stop overcomplicating notes* ([XDA Developers](https://www.xda-developers.com/stop-overcomplicating-notes-lightweight-apps-does-everything-obsidian-does/))

> "There's a whole ecosystem of guides and YouTube tutorials on how to set up Obsidian the 'right way'."
> — same article

**Direct evidence — desktop-widget-vs-app friction:**

> "Even opening them to update or review tasks felt like extra work … Your tasks are there, staring back at you every time you glance at your desktop."
> — *I Ditched To-Do Apps for a Desktop Widget* ([MakeUseOf](https://www.makeuseof.com/windows-11-widget-to-do-list-check/))

**Direct evidence — Sticky Notes pain:**

> "OneDrive sync delays broke their workflow, with 32% experiencing this issue. Additionally, the app doesn't offer proper Sticky Notes desktop widgets, multiple windows, or screen splits."
> — *Top Sticky Notes Alternatives for Enhanced Note-Taking* ([WindowsForum](https://windowsforum.com/threads/top-8-alternatives-to-microsoft-sticky-notes-for-enhanced-note-taking.344191/))

**Maps to QuickSheet:** zero-setup (just type), no plugins, no cloud, no account — the exact opposite of Obsidian's setup-first ritual. Cells autosave; nothing to configure before first use.

**Caveat:** this audience expects sync. Be honest in messaging — QuickSheet is single-machine, file-based, and that's the *point* (the README already does this; reinforce).

**Where to engage:** r/productivity, r/ObsidianMD (carefully — frame as complement, not replacement), r/Notion, r/PKMS, MacRumors note-taking threads.

---

## Segment 3 — HN Terminal-Spreadsheet Crowd (sc-im, VisiData)

Smallest audience by raw count, largest by signal-to-noise. These people *already want* a non-Excel spreadsheet, value keyboard-driven UX, and explicitly cite supply-chain hygiene — which is QuickSheet's stated zero-NuGet stance.

**Direct evidence — the exact ask QuickSheet answers:**

> "If I didn't have so many projects already, I'd give this a shot because I would really love a 'vim' for spreadsheets."
> — *freedomben*, [HN sc-im thread](https://news.ycombinator.com/item?id=47662658)

> "I think a great TUI could get it done though, but it remains to be seen how it could really stack up."
> — *freedomben*, same thread

> "Still from an esthetical perspective I love those simple TUI interfaces. They invoke a weird sense of comfort in me that I can't fully explain."
> — *dodomodo*, same thread

**Direct evidence — supply-chain framing (matches QuickSheet's CLAUDE.md "zero NuGet deps" hard policy):**

> "A spreadsheet that runs in a RISC-V+Core-V device is less susceptible to supply-chain issues and geopolitical stresses."
> — *dsign*, same thread

**Direct evidence — the gap they keep hitting:**

> "I often have some longform text field (like 'Notes') that becomes more fiddly to deal with than a GUI."
> — *VariousPrograms*, same thread

**Other rich threads with overlap:**
- [Sc-im: Spreadsheets in your terminal — HN](https://news.ycombinator.com/item?id=47662658)
- [SC-IM (older HN thread, 2020)](https://news.ycombinator.com/item?id=24318367)
- [VisiData – open-source spreadsheet for the terminal — HN](https://news.ycombinator.com/item?id=45934260)
- [SCIM: Ncurses based, Vim-like spreadsheet — HN](https://news.ycombinator.com/item?id=40876848)
- [Show HN: A terminal spreadsheet editor with Vim keybindings](https://news.ycombinator.com/item?id=47920289)

**Maps to QuickSheet:** keyboard-first navigation, cell prefixes (`r:`, `i:`) as escape hatches, zero deps, CSV persistence (org-mode tables, sc-im, VisiData all use plain text). QuickSheet's GUI mode actually answers the recurring "longform text in cells is fiddly in TUI" complaint while preserving the keyboard-first ergonomics.

**Where to engage:** Show HN with explicit framing ("Vim-flavored spreadsheet, but it lives on your wallpaper, zero deps"), r/commandline, r/vim, lobste.rs.

---

## Segment 4 — Dev Workflow / Morning-Routine Automation

Devs who already write bash scripts to launch their daily repos, terminals, browser tabs. QuickSheet's `r:` cells (multi-select → enter → run all) is a direct alternative to a startup script, but discoverable and tweakable without editing files.

**Direct evidence:**

The DEV Community post ["The 15-Minute Morning Routine That 10x'd My Coding Output"](https://dev.to/cumulus/the-15-minute-morning-routine-that-10xd-my-coding-output-55i5) describes a `focus-mode` bash script that kills distractions and opens tmux + nvim. The author's pain — "the friction of remembering and typing the same setup commands every morning" — is what QuickSheet's runnable-cell launcher targets, with the added benefit of *visible* state (you can see what your morning looks like, and edit it inline).

**Adjacent: launcher users wanting persistence, not popup:**

> "Alfred is one of those tools you only realise how valuable it is when you don't have it."
> — *joaoco*, [HN: PowerToys Run vs Alfred](https://news.ycombinator.com/item?id=31306620)

> "PowerToys Run is fantastic as an app launcher. Not quite as good as Alfred on MacOS but then nothing is."
> — *NelsonMinar*, same thread

> "The thing that Alfred does that the various methods of launching an app don't do is that it switches to an already-running app."
> — *jen729w*, same thread

PowerToys' recently-added [Command Palette Dock](https://www.xda-developers.com/powertoys-added-feature-replaced-third-party-app-launcher-entirely/) — "a second taskbar that can be pinned to any edge of the screen" — proves Microsoft itself sees demand for *pinned, persistent* launcher surfaces, not just popup ones. QuickSheet sits adjacent: not a launcher you summon, but a launcher you *read*.

**Maps to QuickSheet:** `r: code .`, `r: gh pr list`, multi-select repo opener. Inline `i:` output cells (live subprocess) are unique vs. all listed competitors — none of Rainmeter/Conky/PowerToys Run keep live process output as a first-class pinned widget.

**Where to engage:** r/ExperiencedDevs ("show your morning routine" angle), r/commandline, r/programming, lobste.rs, dev.to.

---

## Segment 5 — "I use Excel for everything" personal-tracking crowd

People who track habits, weight, sleep, expenses, books read, etc. in a single spreadsheet because every app is overkill. They open Excel/Sheets for things that should not require a window.

**Direct evidence:**
- [Medium: I Replaced Trello, Notion, and Todolist With Excel](https://medium.com/@kbala7092/i-replaced-trello-notion-and-todolist-with-excel-heres-how-d59a293ca321) — author added a "Today" tab, manages tasks + time blocks + notes in a single .xlsx because "maintaining multiple systems was exhausting … spending more time managing tools than doing meaningful work."
- [MrExcel forum: Using Excel For Note Taking](https://www.mrexcel.com/board/threads/using-excel-for-note-taking.1001864/)

**Maps to QuickSheet:** auto Σ per column, auto Π per row in status bar (the README already calls out "I hate opening the calculator for simple operations" — that's exactly this segment). CSV import/export means their existing Excel files come over.

**Caveat:** this audience often wants charts, formulas, conditional formatting. QuickSheet has none of that. Lead with the "ambient calculator" framing, not "Excel replacement."

**Where to engage:** r/excel ("show me your weird personal tracker" threads do well), r/personalfinance (budget pinned to wallpaper), r/getdisciplined.

---

## Feature Gaps Surfaced by Research

Things the audience repeatedly asks for that QuickSheet lacks. Not all worth building — flagged for awareness:

| Gap | Mentioned by | Worth building? |
|-----|--------------|-----------------|
| Cross-device sync | every notes/todo segment | **No** — kills zero-deps + supply-chain story. Lean into local-first. |
| Cell formulas / `=A1+B1` | HN sc-im threads ("not a spreadsheet without formulas") | **Maybe** — Σ/Π is in the right direction. A `=` prefix would close a real gap and matches the existing prefix system. |
| Markdown / longform in cells | *VariousPrograms* on HN, Obsidian refugees | **Yes** — small win. Already have multi-line cell editing per README; surface it. |
| Cloud / mobile companion | productivity segments | **No** — wrong audience, kills positioning. |
| Plugin/skin ecosystem | Rainmeter / Flow Launcher users | **Maybe later** — risks supply chain, but `r:` cells are already plugin-shaped. |
| Multiple "sheets" / tabs | Excel-as-everything segment | **Probably yes** — natural extension, low risk. |

---

## Risks / Wrong-Audience Signals

- **Notion power-users**: will not switch. Their workflow is database-views and shared pages. Don't pitch there.
- **r/productivity at large**: dominated by Notion/Todoist/TickTick discussion. Posting "yet another todo tool" gets ignored. Pitch as *desktop dashboard*, not *todo app*.
- **Mobile-first users**: hard pass. QuickSheet has no mobile story by design.
- **Enterprise / team users**: collaboration absent. Personal tool, frame accordingly.

---

## Concrete Posting Recommendations (priority order)

1. **Show HN**: "QuickSheet — a Vim-flavored spreadsheet that lives on your desktop wallpaper, zero deps". HN audience already proven receptive (sc-im / VisiData threads). Lead with supply-chain framing + CSV + screenshots.
2. **r/Rainmeter / r/conky**: "I built a thing that replaces my wallpaper with an interactive grid — anyone want to roast it?" Frame as a peer-built tool, not a Rainmeter replacement.
3. **r/unixporn**: a screenshot of a heavily-customized QuickSheet desktop is the entire post. Will succeed or fail on aesthetic, not features.
4. **r/commandline / r/vim**: lead with keyboard shortcuts table + `r:` cells. Audience overlaps heavily with HN sc-im crowd.
5. **lobste.rs**: similar to HN, more curated. Submit after HN if it goes well.
6. **r/ExperiencedDevs / dev.to**: "show your morning routine" angle, with QuickSheet as the visible front-end.

---

## Sources (consolidated)

**Hacker News threads:**
- https://news.ycombinator.com/item?id=47662658 — Sc-im: Spreadsheets in your terminal
- https://news.ycombinator.com/item?id=45934260 — VisiData – open-source spreadsheet for the terminal
- https://news.ycombinator.com/item?id=24318367 — SC-IM (2020)
- https://news.ycombinator.com/item?id=40876848 — SCIM: Ncurses based, Vim-like spreadsheet
- https://news.ycombinator.com/item?id=47920289 — Show HN: terminal spreadsheet w/ Vim keybindings
- https://news.ycombinator.com/item?id=31306620 — PowerToys Run vs Alfred

**Articles / blogs:**
- https://www.makeuseof.com/how-to-build-custom-desktop-dashboard-rainmeter/
- https://www.makeuseof.com/windows-11-widget-to-do-list-check/
- https://www.howtogeek.com/microsoft-already-ruined-windows-11-widgets/
- https://www.xda-developers.com/stop-overcomplicating-notes-lightweight-apps-does-everything-obsidian-does/
- https://www.xda-developers.com/apps-give-desktop-widgets-microsoft-wont/
- https://www.xda-developers.com/powertoys-added-feature-replaced-third-party-app-launcher-entirely/
- https://medium.com/@kbala7092/i-replaced-trello-notion-and-todolist-with-excel-heres-how-d59a293ca321
- https://dev.to/cumulus/the-15-minute-morning-routine-that-10xd-my-coding-output-55i5
- https://mosaid.xyz/articles/from-terminal-to-desktop--sticky-todo-reminders-with-cron,-systemd-&-conky-14.html
- https://maboxlinux.org/mabox-26-03-conky-improvements-todolist-and-keyboard-shortcuts/

**Forums:**
- https://forum.rainmeter.net/viewtopic.php?t=36060 — Project "Desktop Notes"
- https://forum.rainmeter.net/viewtopic.php?t=15850 — Skin for directly input notes
- https://forum.rainmeter.net/viewtopic.php?t=43406 — Obsidian Daily Notes Display Rainmeter Skin
- https://bbs.archlinux.org/viewtopic.php?id=74705 — Conky as a Todo list
- https://bbs.archlinux.org/viewtopic.php?id=74174 — Conky todo list (Arch)
- https://bbs.archlinux.org/viewtopic.php?id=133245 — simple ToDo list for conky
- https://www.mrexcel.com/board/threads/using-excel-for-note-taking.1001864/
- https://windowsforum.com/threads/top-third-party-widgets-for-customizing-your-windows-11-desktop-in-2025.368171/
- https://windowsforum.com/threads/top-8-alternatives-to-microsoft-sticky-notes-for-enhanced-note-taking.344191/

**Existing similar tools (proves demand):**
- https://github.com/alperenozlu/rainmeter-todo
- https://github.com/Pernickety/rainmeter-todo-list
- https://github.com/saulpw/visidata
- https://github.com/andmarti1424/sc-im

---

## Notes on Method

- Puppeteer MCP was used to verify the HN sc-im thread renders as expected and to capture page state for the MakeUseOf dashboard article (screenshots viewed but not persisted to disk — the configured `puppeteer-mcp-server` returns screenshots inline only, not as savable artifacts).
- WebFetch was used for direct-quote extraction; quotes above are verbatim from the linked sources.
- Search depth was bounded by the 30-min window per the user's request; further depth would benefit from authenticated reddit thread scraping (puppeteer can do this if reddit login is provided).
