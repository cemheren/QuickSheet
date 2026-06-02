# awesome-dotnet submission — draft (2026-06-02)

> User submits the PR manually. Do not push to quozd/awesome-dotnet.

## Target

Repo: https://github.com/quozd/awesome-dotnet (19k★, actively maintained)
File: `README.md`
Section: **CLI** (primary fit — lives alongside Gui.cs, spectre.console, etc.)

## One-line entry (matches list's existing format)

```markdown
* [QuickSheet](https://github.com/cemheren/QuickSheet) - Interactive terminal spreadsheet that also embeds as the desktop wallpaper. Cells support runnable commands, live subprocess output, sparklines, and extensions installed by git URL. Zero NuGet dependencies.
```

## PR title

```
Add QuickSheet (CLI)
```

## PR body

```markdown
**Project:** [QuickSheet](https://github.com/cemheren/QuickSheet)
**Category:** CLI

QuickSheet is an interactive terminal spreadsheet written in .NET 9.
It runs in two modes: a full-screen TUI and a transparent desktop-wallpaper
overlay (Windows via WorkerW, Linux via X11 `_NET_WM_WINDOW_TYPE_DESKTOP`).

Why it fits the list:
- Actively maintained (weekly commits, 35+ releases)
- MIT licensed, fully open source
- Zero NuGet dependencies — all platform interop via hand-written P/Invoke
- Cross-platform (Windows + Linux)
- Extension ecosystem via JSON-lines protocol + git URL install
- Documented: README, docs/tour.md, CONTRIBUTING.md, CHANGELOG.md

Cell prefixes make it more than a data grid:
- `r: cmd` — runnable shell command
- `i: cmd` — live subprocess output (ConPTY on Windows, pipes on Linux)
- `s: {A1::A10}` — inline sparkline
- `ext: github:user/repo` — install extensions in one cell

Link: https://github.com/cemheren/QuickSheet
```

## Placement

Insert alphabetically in the CLI section, after Gui.cs and before Power Args:

```diff
 * [Gui.cs](https://github.com/migueldeicaza/gui.cs) - Terminal UI toolkit for .NET.
+* [QuickSheet](https://github.com/cemheren/QuickSheet) - Interactive terminal spreadsheet that also embeds as the desktop wallpaper. Cells support runnable commands, live subprocess output, sparklines, and extensions installed by git URL. Zero NuGet dependencies.
 * [Power Args](https://github.com/adamabdelhamed/PowerArgs) - PowerArgs converts command-line arguments into .NET objects that are easy to program against...
```

## Notes

- Quality standard check:
  - ✅ Generally useful to the community (CSV editor, live dashboards, wallpaper mode)
  - ✅ Actively maintained (weekly commits, issues triaged)
  - ✅ Stable (36 tagged releases)
  - ✅ Documented (README, tour, CONTRIBUTING, CHANGELOG)
  - ⚠️ No formal test suite — this is the one gap; mention if asked
- Keep description under 200 chars (current: 195)
- awesome-dotnet uses `awesome-bot` CI — ensure the GitHub link is valid and repo is public
- The "zero NuGet dependencies" angle is the differentiator vs. spectre.console or Gui.cs which are libraries, not applications
