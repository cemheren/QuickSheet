# Awesome-list submission drafts — QuickSheet

Drafted 2026-05-14. Each target list has its own submission style. User reviews,
forks the target repo, opens the PR.

---

## 1. awesome-tuis — https://github.com/rothgar/awesome-tuis

Style: alphabetically sorted, one-line description, link first.

**Section:** `### Productivity` (or `### Spreadsheets` if it exists; otherwise `Productivity`).

**Diff line to add:**

```markdown
- [QuickSheet](https://github.com/cemheren/QuickSheet) - Terminal spreadsheet that doubles as your desktop wallpaper. Cell prefixes for runnable commands, live subprocess output, hyperlinks, and sparklines. Zero NuGet dependencies, CSV persistence. Windows + Linux.
```

**PR title:** `Add QuickSheet`

**PR body:**

```
Adds [QuickSheet](https://github.com/cemheren/QuickSheet) under Productivity.

QuickSheet is an interactive terminal spreadsheet (C# / .NET 9) with an
unusual second mode: it embeds itself as a transparent, interactive desktop
wallpaper. Cell prefixes let you mix data with launchers, live subprocess
output, hyperlinks, and inline sparklines.

Notes:
- Cross-platform: Windows (WinForms + WorkerW) and Linux (raw X11).
- Zero NuGet dependencies — all native interop is hand-written P/Invoke.
- MIT licensed, active, has working screenshots in the README.
- Repo is alphabetized; placed under Productivity per CONTRIBUTING.
```

---

## 2. awesome-dotnet — https://github.com/quozd/awesome-dotnet

Style: alphabetical inside each section, one-line description.

**Section candidate:** `Console` (if exists) or `Frameworks, Libraries and Tools → Tools` or `Applications`.

**Diff line:**

```markdown
* [QuickSheet](https://github.com/cemheren/QuickSheet) - Terminal spreadsheet with a desktop-wallpaper mode. Runnable command cells, live subprocess output, sparklines. Zero NuGet dependencies. Windows + Linux. .NET 9.
```

**PR title:** `Add QuickSheet (terminal spreadsheet with desktop-wallpaper mode)`

**PR body:**

```
Adds QuickSheet — an interactive terminal spreadsheet for .NET 9.

Why it fits awesome-dotnet:
- Pure .NET, zero NuGet dependencies — all native interop is hand-written
  P/Invoke (X11, WinForms, ConPTY).
- Cross-platform via OS-conditional TFMs in the csproj.
- MIT, README + screenshots, working binaries.

Link: https://github.com/cemheren/QuickSheet
```

---

## 3. awesome-cli-apps — https://github.com/agarrharr/awesome-cli-apps

Style: nested categories, one-line description, alphabetical.

**Section candidate:** `### Productivity` → `#### Spreadsheets` (create if missing).

**Diff line:**

```markdown
- [QuickSheet](https://github.com/cemheren/QuickSheet) - Terminal spreadsheet with a desktop-wallpaper mode.
```

**PR title:** `Add QuickSheet under Productivity`

(Keep PR body short — this list prefers minimal additions. Reference the README screenshots.)

---

## 4. awesome-console-services / awesome-console — https://github.com/chubin/awesome-console-services

Skip unless an explicit "TUI apps" section appears — list is mostly remote services.

---

## 5. terminaltrove.com submission — https://terminaltrove.com/submit

terminaltrove uses a form, not a PR. Fields to fill:

| Field        | Value                                                         |
|--------------|---------------------------------------------------------------|
| Name         | QuickSheet                                                    |
| Tagline      | Terminal spreadsheet that doubles as your desktop wallpaper.  |
| Repo URL     | https://github.com/cemheren/QuickSheet                        |
| Categories   | spreadsheet, productivity, csv, dotnet                        |
| Description  | (see paragraph below)                                         |
| Screenshot   | upload one of the README PNGs (image-1.png or image-2.png)    |
| License      | MIT                                                           |
| OS           | Windows, Linux                                                |

**Description:**

```
QuickSheet is an interactive terminal spreadsheet for .NET 9 that can also
embed itself as a transparent, interactive desktop wallpaper. Cell prefixes
let you mix plain text with runnable commands (r:), live subprocess output
(i:), hyperlinks, and inline sparklines (s:). It has an extension system —
`ext: github:user/repo` clones a repo and registers a new cell prefix that
talks to QuickSheet over JSON-lines. Persistence is plain CSV. Zero NuGet
dependencies; all native interop (X11, WinForms, ConPTY) is hand-written
P/Invoke. Cross-platform: Windows (WorkerW) and Linux (X11).
```

---

## 6. console.dev — https://console.dev/submit-a-tool

Editorial; they pick. Submit form fields are similar to terminaltrove. Same
description as above works.

---

## Order of operations (suggested for user)

1. **terminaltrove + console.dev first** — they're submission forms, fastest to
   queue. Editorial review takes weeks but costs nothing to submit early.
2. **awesome-tuis PR** — smallest list, fastest review, highest signal.
3. **awesome-dotnet PR** — biggest list, biggest reach, slower review.
4. **awesome-cli-apps PR** — last; the maintainer has strict style rules.

Stagger by a few days so PRs don't read as a blast.
