# awesome-cli-apps submission — draft (2026-06-03)

> User submits the PR manually. Do not push to agarrharr/awesome-cli-apps.

## Target

Repo: https://github.com/agarrharr/awesome-cli-apps (~19.7k★, actively maintained)
File: `readme.md`
Section: **Data Manipulation** (alongside visidata and sc-im — both spreadsheet tools)

## Prerequisites before submission

- **Star count ≥ 20** — contributing guidelines require it. Currently 0. Wait.
- **Repo age > 90 days** — ✅ created 2026-03-04 (91 days as of today).
- One PR per app. Title it exactly `Add QuickSheet`.

## Entry (matches list's format exactly)

Add at the **bottom** of the Data Manipulation top-level bullet list (after `nless`):

```markdown
- [QuickSheet](https://github.com/cemheren/QuickSheet) - Interactive spreadsheet that doubles as a desktop wallpaper with runnable cells and extensions.
```

### Why this description works

- Starts with capital, ends with full stop.
- No "CLI" or "terminal" (per guidelines: no redundant info).
- Short, concrete, differentiating (wallpaper + runnable cells + extensions — nothing else in that section does this).
- Under 120 chars.

## PR title

```
Add QuickSheet
```

## PR body

```markdown
[QuickSheet](https://github.com/cemheren/QuickSheet) is an interactive spreadsheet that can also embed itself as a transparent desktop wallpaper (Windows via WorkerW, Linux via X11 `_NET_WM_WINDOW_TYPE_DESKTOP`). Cells support shell commands (`r: git status`), live subprocess output (`i: htop`), sparklines, and a JSON-lines extension protocol.

- **License:** MIT
- **Language:** C# (.NET 9), zero NuGet dependencies
- **Platforms:** Windows, Linux (X11)
- **Install:** `dotnet run` or self-contained binaries from GitHub Releases

It fits **Data Manipulation** alongside visidata and sc-im as a spreadsheet tool — but with a unique angle: it's both a CSV editor and a composable desktop surface.
```

## Notes for the submitter

- The 20-star minimum is a hard gate. Do not submit before reaching it.
- Use the repo's pull request template if one is provided.
- One commit, one app, clean diff. Recent merged PRs follow this pattern.
- If asked how it differs from visidata: visidata is read-mostly data exploration; QuickSheet is a read-write grid that persists to CSV, runs commands, and optionally replaces your wallpaper.
- If asked how it differs from sc-im: sc-im is a Vim-like calculator; QuickSheet is a composable desktop surface with extensions, inline processes, and sparklines.

## After submission (expected impact)

- Inbound from a 19.7k-star list. The Data Manipulation section is browsed by people specifically looking for spreadsheet/data CLI tools.
- Realistic: 10–30 stars in long-tail traffic if accepted (based on similar-sized entries getting steady organic clicks).
- The list is widely referenced in "best CLI tools" blog posts, amplifying reach.

## Comparison to other awesome-list drafts

| List                                  | Stars  | Fit section          | Status        |
|---------------------------------------|--------|----------------------|---------------|
| luong-komorebi/Awesome-Linux-Software | ~22k   | Utilities            | draft ready   |
| 0PandaDEV/awesome-windows             | ~2.4k  | Customization        | draft ready   |
| **agarrharr/awesome-cli-apps**        | ~19.7k | Data Manipulation    | **this draft** |
| awesome-csharp                        | smaller| Tools                | draft ready   |
| awesome-dotnet                        | medium | CLI                  | draft ready   |
