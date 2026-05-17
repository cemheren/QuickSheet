# Awesome-csharp + awesome-windows-apps submission drafts

Drafted 2026-05-14. Two more target lists beyond what's already in `awesome-lists.md`. Each is a separate PR; submit a few days apart so it doesn't read as burst.

> **2026-05-17 update:** the `Awesome-Windows/Awesome` target in §2 below is dead (404 on GitHub). Use the fresh draft at [`awesome-windows.md`](awesome-windows.md) which targets the active `0PandaDEV/awesome-windows` list (2.4k stars) instead. §1 (awesome-csharp at `uhub/awesome-csharp`) is still live and submitable.

---

## 1. awesome-csharp — https://github.com/uhub/awesome-csharp

(Distinct from awesome-dotnet; this list is more application-focused.)

**Style:** alphabetical inside each section, one-line description.

**Section candidate:** `Tools` or `Applications` (whichever your read shows is closer to TUIs / developer tools).

**Diff line:**

```markdown
* [QuickSheet](https://github.com/cemheren/QuickSheet) - Terminal spreadsheet that doubles as your desktop wallpaper. Cell prefixes for runnable commands, live subprocess output, sparklines. Cross-platform (Windows + Linux), zero NuGet dependencies, .NET 9.
```

**PR title:** `Add QuickSheet`

**PR body:**

```
Adds QuickSheet under {Tools|Applications} — pick the section that fits best at review time.

QuickSheet is an interactive terminal spreadsheet (C# / .NET 9) with an
unusual second mode: it embeds itself as a transparent, interactive desktop
wallpaper. Cell prefixes let you mix data with launchers, live subprocess
output, hyperlinks, and inline sparklines.

Why it fits awesome-csharp:
- Pure C# / .NET 9.
- Zero NuGet dependencies — all native interop is hand-written P/Invoke.
- MIT licensed, README with screenshots, active.
- 14 separate extension repos (also in C#) hook in via stdin/stdout JSON-lines.

Link: https://github.com/cemheren/QuickSheet
```

---

## 2. awesome-windows-apps — https://github.com/Awesome-Windows/Awesome

**Style:** GitHub-style table per section, with logo + name + description.

**Section candidate:** `Productivity → Note-taking & Office Tools`, or `Developer Tools`. Per category, items are alphabetical.

**Row to add (Markdown):**

```markdown
| [QuickSheet](https://github.com/cemheren/QuickSheet) | Terminal spreadsheet that also embeds as the desktop wallpaper. Runnable commands, live subprocess output, sparklines, extensions installed by git URL. |
```

**PR title:** `Add QuickSheet under Developer Tools (or Productivity)`

**PR body:**

```
Adds QuickSheet under {Developer Tools|Productivity}.

A free, open-source (MIT) terminal spreadsheet for Windows that can also
embed itself as the desktop wallpaper using the standard WorkerW technique.
Notes, runnable commands, live subprocess output, hyperlinks, sparklines,
and a small extension system. Cross-platform with a Linux build too.

Repo: https://github.com/cemheren/QuickSheet

Honors the awesome-windows-apps CONTRIBUTING guide:
- Open-source ✓
- Active development ✓
- README + screenshots ✓
- Windows binary requires .NET 9 SDK (noted in install steps).
```

---

## Submission strategy

- **awesome-csharp** first (smaller list, faster review).
- **awesome-windows-apps** ~5 days later (huge list, more selective, but high signal to Windows users).
- Both reference the canonical repo, not a fork.
- Do not include "please star" in the PR body or in the review thread.

## Coordination with prior drafts

Earlier `drafts/awesome-lists.md` covered:
- awesome-tuis
- awesome-dotnet
- awesome-cli-apps
- terminaltrove
- console.dev

Stagger across all of these so PRs don't burst over a single week. Suggested order (one PR per ~3 days):

1. awesome-tuis
2. terminaltrove form submit
3. awesome-dotnet
4. console.dev form submit
5. awesome-cli-apps
6. **awesome-csharp** (this draft)
7. **awesome-windows-apps** (this draft)
