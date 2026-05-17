# QuickSheet for students

If you're juggling coursework deadlines, a job search, a side project on GitHub, and a
budget that resets every month — **your desktop wallpaper can track all of it.** No
browser tab. No Notion subscription. No "wait let me open that app."

> Already familiar with Notion, Obsidian, or Logseq? Think "those, but on the actual
> wallpaper, zero cloud, and talking to live APIs through extensions you can write in
> 50 lines of any language."

## Five-minute setup

```bash
git clone https://github.com/cemheren/QuickSheet.git
cd QuickSheet
dotnet run -c Release --project ExcelConsole.csproj -- --desktop examples/student-dashboard.csv
```

Edit cells live, autosaves every 5 seconds. Survive reboots — the grid is exactly where
you left it.

## What goes on the wallpaper

The starter sheet at `examples/student-dashboard.csv` covers four zones you can resize or
swap out:

| Zone              | What it shows                                              | How                                                         |
|-------------------|------------------------------------------------------------|-------------------------------------------------------------|
| **Coursework**    | Class, assignment, due date, status                        | Plain cells — edit directly                                 |
| **Budget**        | Spend vs budget per category, colour-coded progress bars   | [`budget`](https://github.com/cemheren/quicksheet-budget)  |
| **Tech tools**    | JWT decoder, regex explainer, base64, URL encoder in-cell  | [`jwtdec`](https://github.com/cemheren/quicksheet-jwtdec) · [`regex`](https://github.com/cemheren/quicksheet-regex) · [`b64`](https://github.com/cemheren/quicksheet-b64) · [`urlenc`](https://github.com/cemheren/quicksheet-urlenc) |
| **News & GitHub** | HN top 5, your GitHub commit log                           | [`hntop`](https://github.com/cemheren/quicksheet-hntop) · [`gitlog`](https://github.com/cemheren/quicksheet-gitlog) |

All extensions install with a single `ext: github:cemheren/quicksheet-<name>` cell —
QuickSheet clones the repo and the prefix is live instantly.

## Coursework + deadline tracking

Plain cells are all you need for a class tracker:

```
Course,Assignment,Due,Status
CS 301,Homework 3,2026-05-20,🟡 In progress
CS 420,Term paper,2026-06-01,⬜ Not started
MATH 201,Problem set 5,2026-05-18,🟢 Done
```

Ctrl+B to sort by due date. Ctrl+F to search. Ctrl+Z if you accidentally cleared a cell.

Add a world-clock strip for remote study partners:

```
ext: github:cemheren/quicksheet-worldtm
---
worldtm: ny
worldtm: london
worldtm: tokyo
```

## Budget tracker on your desktop

The [`budget`](https://github.com/cemheren/quicksheet-budget) extension turns a CSV into a
live envelope-budget display:

```
ext: github:cemheren/quicksheet-budget
---
budget: Rent:1200:1200
budget: Groceries:180:250
budget: Coffee:34:40
budget: Textbooks:0:150
```

Each cell renders a progress bar and a 🟢🟡🟠🔴 indicator.  No subscription. No bank API
— you just type in what you spent.

## Developer tools always one glance away

Paste a JWT and decode it in-cell without opening jwt.io (privacy win):

```
ext: github:cemheren/quicksheet-jwtdec
jwtdec: eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9...
```

Stuck on a regex in an assignment? Explain it in-cell:

```
ext: github:cemheren/quicksheet-regex
regex: ^(\d{4})-(0[1-9]|1[0-2])-(0[1-9]|[12]\d|3[01])$
```

Decode a base64 blob from an API response without leaving the terminal:

```
ext: github:cemheren/quicksheet-b64
b64: SGVsbG8gQ1MgMzAxIQ==
```

## Stay in sync with the tech world

Keep HN's top 5 stories visible on your wallpaper — never miss a link that matters:

```
ext: github:cemheren/quicksheet-hntop
hntop: 5
```

See your own recent commits without running `git log`:

```
ext: github:cemheren/quicksheet-gitlog
gitlog: /path/to/your/project, 8
```

## Hackathon mode

The day of a hackathon, swap the coursework zone for a task list and the budget zone for
a service-health strip:

```
ext: github:cemheren/quicksheet-health
health: http://localhost:3000
health: http://localhost:8080/api/health
health: https://api.github.com
```

Every service glows green/red on your wallpaper. You know immediately if your API server
crashed without a browser tab.

## Make it yours

Cycle themes with **Ctrl+T**: Dark, Light, Nord, Solarized, Matrix, Dracula, Synthwave,
Gruvbox, Monokai.  Nord pairs well with a dark terminal rice; Matrix is fun for late-night
sessions.

Color-code priority cells:

```
c?: >90=red, >70=yellow, *=green: 85
```

Useful for "percent of budget used" cells.

## Why wallpaper > app

- **Always visible** — the tracker is behind every window, no tab-switching required.
- **Zero cloud** — your schedule is a CSV on your machine.
- **Hackable** — write a 50-line `quicksheet-myCourse` extension for your syllabus API
  ([extension-protocol.md](extension-protocol.md)). JSON-lines over stdin/stdout, any language.
- **CS course credit potential** — some professors accept "built an extension for a TUI
  spreadsheet" as a systems programming assignment. Just saying.

## See also

- [Full extension directory](extensions.md)
- [Keyboard shortcuts](keyboard-shortcuts.md)
- [Extension protocol spec](extension-protocol.md) — write your own in an afternoon
- [for-homelab.md](for-homelab.md) — if you're also running a home server
