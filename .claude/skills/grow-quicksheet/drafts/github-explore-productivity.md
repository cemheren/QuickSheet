# GitHub Explore: productivity-tools Collection Submission

## Target
Repository: [github/explore](https://github.com/github/explore)
File: `collections/productivity-tools/index.md`
Change: Add `cemheren/QuickSheet` to the `items` list.

## Diff

```yaml
# Add this line to the items list in collections/productivity-tools/index.md:
 - cemheren/QuickSheet
```

## PR Title

> Add cemheren/QuickSheet to productivity-tools collection

## PR Body

```markdown
### What is QuickSheet?

[QuickSheet](https://github.com/cemheren/QuickSheet) is a .NET 9 terminal
spreadsheet that replaces your desktop wallpaper with an interactive,
transparent grid. It's a zero-dependency productivity tool for developers
who want their desktop to do something useful:

- **Always-on scratchpad** — click the desktop to jot notes, autosaves every 5s
- **App launcher** — prefix cells with `r: code .` and hit Enter
- **Link dashboard** — paste URLs, they open on Enter
- **Live data** — column sums, sparklines, inline subprocesses, 69+ extensions
- **Zero dependencies** — clone → `dotnet build` → run. No NuGet, no npm

It runs on Windows (WinForms + Win32 interop) and Linux (raw X11 P/Invoke).

### Why productivity-tools?

QuickSheet fits alongside other terminal/desktop productivity tools in this
collection (Terminal, ripgrep, bat, zoxide, ShareX). It turns idle desktop
space into a persistent workspace — notes, launchers, links, and live data
without opening a window.

### Links

- Repository: https://github.com/cemheren/QuickSheet
- License: MIT
- Language: C#
- Platform: Windows + Linux
```

## Instructions for user

1. Fork https://github.com/github/explore
2. Edit `collections/productivity-tools/index.md`
3. Add ` - cemheren/QuickSheet` to the `items` list (anywhere in the list)
4. Commit and open a PR with the title and body above
5. Wait for GitHub maintainers to review

Note: The collection currently has ~46 repos. GitHub maintainers are selective
but accept well-maintained, useful projects. QuickSheet's zero-dep policy,
dual-platform support, and active development make a strong case.
