# Reddit r/commandline draft — QuickSheet

Drafted 2026-05-14. Subreddit rules: r/commandline rewards working demos and discourages "promotional" framing. Lead with the thing, not the project.

## Title

**QuickSheet — a terminal spreadsheet I built that also runs as my desktop wallpaper**

(Alternative if mods push back on self-promo language: `My terminal spreadsheet doubles as the desktop wallpaper — feedback welcome`.)

## Body

```
I wanted a "wallpaper that does something" instead of a static background, so I built a TUI spreadsheet that can also embed itself as the desktop wallpaper. Notes, runnable commands, live subprocess output, hyperlinks, sparklines — all in cells, always visible behind every window.

It's a regular terminal spreadsheet when you run it normally. The trick is the `--desktop` flag, which uses the Win32 WorkerW handle on Windows and `_NET_WM_WINDOW_TYPE_DESKTOP` on Linux/X11 to embed the same grid as the wallpaper.

Cell prefixes are the killer feature for me:

- `r: code .` → press Enter, launches VS Code
- `i: ping -c 1 example.com` → output streams back into the cell live (ConPTY on Windows, pipe redirect on Linux)
- `s: 4,7,9,3,8,12` → renders as `▂▃▆▁▅█` sparkline
- `L: A10, 5m` → loop that cell every 5 minutes
- `ext: github:user/repo` → installs an extension (subprocess, JSON-lines protocol)
- Any URL → highlighted, opens on Enter

Multi-select cells, hit Enter, and all of them fire — that's how I "open my whole dev environment" in one keystroke.

A few constraints I'm holding to:
- Zero NuGet dependencies. All native interop (X11, WinForms, ConPTY) is hand-written P/Invoke. Clone, build, run.
- CSV is the only persistence format. Open the file in Excel or vim. Same data.
- Cross-platform conditional compilation, not a shim layer.

Repo: https://github.com/cemheren/QuickSheet

Honest about the state: it's a side project, mostly written with AI assist, but the zero-deps rule keeps the surface area sane. Working examples + screenshots in the README. Wayland support is the obvious gap.

Curious whether anyone else has tried the "wallpaper does something" thing — I haven't found a prior art that's interactive.
```

## Tips for the actual post

- Post **Mon–Wed, 8–10am US Central**. r/commandline is most active mid-week mornings.
- Image flair preferred. Attach the launcher screenshot (`image-4.png`) as the post image — wallpaper mode is the visual hook.
- Have a 10s GIF link ready for the first reply. r/commandline rewards GIFs more than HN does.
- Do not cross-post to r/linux or r/dotnet within 24 hours — looks like a spam burst.
- Reply to every comment in the first 2 hours, especially "why not [other tui]?" — answer with what's *different*, not what's better.

## Pitfalls to avoid

- Don't mention HN ranking or "I posted this on HN earlier" — mods may remove as crosspost.
- Don't use marketing-y phrasing ("revolutionary", "game-changing"). Subreddit hates it.
- Don't ask for stars/upvotes anywhere in the post.
