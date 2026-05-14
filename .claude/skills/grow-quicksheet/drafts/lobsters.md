# Lobsters submission draft — QuickSheet

Drafted 2026-05-14. Lobsters rewards technical depth over UX polish — lead with the unusual technique.

## Title

**QuickSheet — A zero-NuGet .NET 9 spreadsheet that embeds as the desktop wallpaper**

(Alt: `Embedding a TUI app as the desktop wallpaper on Windows (WorkerW) and Linux (X11 _NET_WM_WINDOW_TYPE_DESKTOP)`.)

## URL

`https://github.com/cemheren/QuickSheet`

## Tags (Lobsters tag rules — pick existing tags only)

- `programming`
- `dotnet`
- `windows`
- `linux`

## Story description (the short paragraph shown on the front page)

> A side project that turns the desktop wallpaper into an interactive spreadsheet. Cell prefixes activate runnable commands, live subprocess output, hyperlinks, sparklines, and extensions installed by git URL. Cross-platform .NET 9. The notable constraint: zero NuGet dependencies — all native interop (Win32 WorkerW, X11, ConPTY) is hand-written P/Invoke.

## First comment to post immediately after submission

```
Author here. Couple of things I think Lobsters-aged readers might find interesting:

1. **WorkerW trick** (Windows): the standard wallpaper-embedding technique. Send a magic SendMessageTimeout to Progman with wParam=0x52c, it spawns a WorkerW behind the desktop icons, SetParent the form into that. Z-order needs locking because explorer.exe re-asserts itself, and Win+D (show desktop) needs detection to keep the window visible. The handling lives in `Platform/Windows/DesktopFormBase.cs` if anyone wants to compare with their own.

2. **X11 wallpaper** (Linux): I expected this to be a tarpit of compositor-specific protocols, but `_NET_WM_WINDOW_TYPE_DESKTOP` on a raw X11 window actually works on most stacking WMs and on Xfce / Cinnamon / GNOME-on-Xorg. Wayland is the open problem — see issue #3.

3. **Zero NuGet deps**: this was a deliberate constraint, not a vibes thing. All P/Invoke is hand-written against `libX11.so.6`, `libXft.so.2`, and Win32 DLLs. ConPTY for live subprocess cells on Windows; plain pipe redirect on Linux. The csproj has no `<PackageReference>` entries.

4. **The cell-prefix concept** — `r:` for runnable, `i:` for inline subprocess output piped into the cell live, `s:` for sparkline, `ext: github:user/repo` to install an extension (clones the repo, reads a manifest, spawns a subprocess that talks JSON-lines on stdin/stdout) — is the part I'd be curious to hear pushback on. Is this a sane way to extend a TUI, or am I just reinventing Emacs poorly?

The codebase is small enough to read in a sitting if you're interested in the cross-platform-conditional-compilation pattern. Highlights:
- `Program.cs` — `#if PLATFORM_WINDOWS / PLATFORM_LINUX` dispatch.
- `InlineProcessManager.cs` — ConPTY vs pipe redirect.
- `Platform/{Windows,Linux}/` — the OS-conditional folders that the csproj `<Compile Remove>`s on the wrong OS.

Repo: https://github.com/cemheren/QuickSheet
```

## Tips for posting

- Lobsters dislikes posts that read as marketing. Lead with the technique, not the UX.
- Pick tags carefully — `programming` is broad, but the OS tags signal cross-platform interest.
- Submit weekday morning EST. Avoid late nights (small community, fewer readers).
- Engage in comments. Lobsters readers will dig into the P/Invoke choices and you'll learn things.

## Pitfalls

- Do not crosspost from HN or Reddit and mention the crosspost. Lobsters mods penalize.
- Do not add affiliate links or sponsorship CTAs (auto-flag).
- Avoid "Show HN" phrasing — Lobsters has no equivalent custom; just a normal submission.
