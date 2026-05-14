# Twitter/X thread draft — QuickSheet

Drafted 2026-05-14. Five tweets. Tweet 1 is the hook; should stand alone if no one reads further. Each tweet ≤ 280 chars.

## Tweet 1 (the hook)

```
I made my desktop wallpaper a spreadsheet.

Notes, runnable commands, live process output, hyperlinks, sparklines — all in cells, always behind every window. Click any window to focus it, click empty desktop to start editing the grid.

QuickSheet. .NET 9. Zero NuGet deps.
```

Attach: `image-4.png` (launcher screenshot) or — better — a 10-second screencap of the wallpaper-mode interaction. The visual is the whole pitch.

## Tweet 2 — the technical wedge

```
The wallpaper trick:
- Windows: WorkerW handle behind the desktop icons (SendMessageTimeout to Progman, then SetParent into the spawned WorkerW)
- Linux: raw X11 window with _NET_WM_WINDOW_TYPE_DESKTOP

Same C# code, OS-conditional TFMs in the csproj. All P/Invoke. No NuGet.
```

## Tweet 3 — the killer feature

```
Cell prefixes are the actual feature:

r: code .            ← launches VS Code on Enter
i: ping example.com  ← output streams live into the cell (ConPTY on Windows)
s: 4,7,9,3,8,12     ← renders as ▂▃▆▁▅█ sparkline
ext: github:u/r      ← installs an extension by git URL
```

## Tweet 4 — extension ecosystem

```
Extensions are separate repos. A manifest + a subprocess that talks JSON-lines on stdin/stdout. New cell prefix appears at runtime.

So far:
- weather forecast
- crypto prices
- TLS cert expiry
- inline dictionary
- mortgage calculator
- AI in a cell (Copilot)

All zero-dep.
```

## Tweet 5 — call to action

```
Side project, written largely with AI assist, but the zero-NuGet rule kept the design honest. CSV is the only persistence format — open it in Excel or vim, same data.

Repo, README, screenshots: https://github.com/cemheren/QuickSheet

Wayland support is the obvious open issue.
```

## Posting tips

- Post **Tue–Thu, 8–10am Pacific** for max US-developer reach.
- Use 2–3 hashtags max, and only on tweet 1 or 5. Candidates: `#dotnet #terminal #commandline #linux #windows`.
- Reply to your own thread with the demo GIF or screencap if you didn't attach it to tweet 1.
- Tag people only if there's a real reason (e.g. .NET community accounts that have signal-boosted similar projects). No mass-tagging.

## Variants

If you want a single tweet instead of a thread (lower engagement, lower friction):

```
My desktop wallpaper is now a spreadsheet.

Runnable commands, live subprocess output, hyperlinks, sparklines, extensions installed by git URL. Cross-platform .NET 9, zero NuGet dependencies (all P/Invoke).

https://github.com/cemheren/QuickSheet
```

## Pitfalls

- Don't post the same thread to Mastodon/Bluesky verbatim within 24h — looks copy-pasted. Adapt the opening line for each platform.
- Don't claim "production-ready" or "1000+ users". Side-project framing.
- Don't lead with the AI-assist disclosure on Twitter; it's noise that distracts from the technique. Save it for the README and HN.
