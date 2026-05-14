# Show HN draft — QuickSheet

Drafted 2026-05-14. Author posts manually. Pick one title; tweak copy as you like.

---

## Title options (HN allows ~80 chars; "Show HN:" prefix recommended)

1. **Show HN: QuickSheet – my desktop wallpaper is a spreadsheet**
2. **Show HN: A terminal spreadsheet that doubles as your desktop wallpaper**
3. **Show HN: QuickSheet – transparent interactive grid replaces my wallpaper**
4. **Show HN: I replaced my wallpaper with a spreadsheet (C#, zero NuGet deps)**

Recommended: **#1**. Short, concrete, slightly weird — the wallpaper angle is the hook.

## URL

`https://github.com/cemheren/QuickSheet`

## First comment (post as soon as the submission goes live — this is the pitch)

```
Author here. QuickSheet started as an itch: I never interact with my wallpaper, and I always have a few things I want at hand — notes, todos, app launchers, frequently-visited URLs, small running totals. So I made the wallpaper itself a transparent, interactive grid.

What it does:
- Cells are plain text by default. `r: cmd` makes a cell a launcher (`r: code .`, `r: firefox github.com`). Pasting a URL auto-detects and opens it on Enter. `i: cmd` runs a subprocess and pipes output back into the cell (live-tail a log, ping a host, watch a counter).
- Persistence is CSV. No database, no JSON state file. Open it in Excel or vim, same file.
- Auto-sum per column, auto-product per row in the status bar — for the times you don't want to open a calculator.
- Extensions are git repos. `ext: github:user/repo` in a cell clones, reads the manifest, and starts a subprocess that talks JSON-lines. One I use is a Copilot extension that answers questions in-cell.
- Two run modes: `--desktop` embeds it as the wallpaper (Win32 WorkerW on Windows, `_NET_WM_WINDOW_TYPE_DESKTOP` on X11), no flag = plain terminal TUI.

Design constraints I'm holding to:
- **Zero NuGet dependencies.** All native interop is hand-written P/Invoke (X11, WinForms, ConPTY). Clone → build → run. Supply-chain surface is tiny.
- Side-project quality, written largely with AI assistance, but the zero-deps rule kept the design honest.

Cross-platform via OS-conditional TFMs in the csproj; the Windows path uses WinForms + WorkerW for wallpaper embedding, Linux uses raw X11 (Wayland gets a warning — patches welcome).

Things I want feedback on:
- The "i:" inline-process cell is the part I use most but it's the rough edge — output is capped at 200 lines and the threading model could be cleaner.
- Whether anyone has a use for sparkline-in-cell or `w: url` live-fetch — both are on the roadmap but I'd rather hear if people actually want them.

Repo: https://github.com/cemheren/QuickSheet
```

## Tips for the actual submission

- Post Tuesday–Thursday, 7–10am Pacific. Best window for `/newest` visibility.
- Do NOT title-bait. HN punishes hype.
- Reply to every comment in the first 2 hours. Engagement helps ranking.
- If someone asks "why not [other tui spreadsheet]?", answer with what's *different*, not what's better. The wallpaper-embedding + extensions angle is the differentiator.
- Have one screenshot or GIF ready to link in a reply — wallpaper mode is the visceral hook. The repo README has stills; a 10-second GIF would be even better.

## Lobsters / Reddit reuse

Title and first comment above port directly to:
- Lobsters (use `programming`, `dotnet` tags)
- r/commandline (drop "Show HN:" prefix, keep body)
- r/dotnet (lean harder on the zero-NuGet / P/Invoke angle in the body)
- r/coolgithubprojects (one-line caption + screenshot is enough)
