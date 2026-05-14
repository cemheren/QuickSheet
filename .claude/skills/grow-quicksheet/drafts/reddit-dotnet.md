# Reddit r/dotnet draft — QuickSheet

Drafted 2026-05-14. r/dotnet rewards .NET-specific angles: P/Invoke, conditional compilation, zero-dependency code. Lead with the technical hook, not the UX hook.

## Title

**QuickSheet — a zero-NuGet .NET 9 terminal spreadsheet that also embeds as the desktop wallpaper**

(Alt title that leans even harder on the tech: `Building a cross-platform .NET 9 app with zero NuGet dependencies (X11 + WorkerW via hand-written P/Invoke)`.)

## Body

```
Built a side project I think r/dotnet might find interesting on a few axes:

**The thing it does:** an interactive terminal spreadsheet (cell grid, CSV persistence, runnable cells, live subprocess output) that can ALSO embed itself as the desktop wallpaper — same grid, transparent, always behind your windows. Cross-platform: Windows + Linux.

**Why r/dotnet might care:**

- **Zero NuGet dependencies.** Hard rule on the repo. The csproj has no `PackageReference`s. All native interop is hand-written P/Invoke.
- **Cross-platform conditional compilation, not abstraction layers.** The csproj defines `PLATFORM_WINDOWS` and `PLATFORM_LINUX` based on the runtime OS, swaps TFMs (`net9.0-windows` vs `net9.0`), and `<Compile Remove>`s the wrong platform's folder. Shared files use `#if PLATFORM_WINDOWS` / `#elif PLATFORM_LINUX`. No shim layer, no DI container, no plugin system.
- **Wallpaper embedding via Win32 WorkerW.** The Windows side uses WinForms and the classic WorkerW trick — send a magic SendMessageTimeout to Progman, find the spawned WorkerW, reparent the form into it. Z-order locking + Win+D detection so the window doesn't get hidden when the user "show desktop"s.
- **X11 wallpaper via raw P/Invoke.** Linux side talks directly to `libX11.so.6` and `libXft.so.2` — no Avalonia, no GTK#. Sets `_NET_WM_WINDOW_TYPE_DESKTOP` so the WM treats it as wallpaper. Wayland gets a warning (Wayland support is the obvious gap).
- **ConPTY for live subprocess cells.** `i: ping example.com` runs a child process and pipes the output back into the cell live, capped at 200 lines. Windows uses ConPTY; Linux uses pipe redirect.
- **Extension protocol.** Extensions are separate repos. `ext: github:user/repo` clones, reads a manifest, and starts a subprocess that talks JSON-lines on stdin/stdout. New cell prefix registered at runtime.

Repo: https://github.com/cemheren/QuickSheet

The codebase is small enough that the P/Invoke files are good reading if you want to see what a no-package-managers .NET 9 app actually looks like. Highlights:
- `Platform/Windows/NativeMethods.cs` — Win32 imports.
- `Platform/Linux/` — X11 + Xft imports.
- `InlineProcessManager.cs` — ConPTY vs pipe redirect.
- `Program.cs` — the conditional-compilation dispatch.

Honest about the state: it's a side project, written with significant AI assist. The zero-deps rule keeps the design honest — you can't paper over a bad idea with a package.

Curious if anyone else here ships pure-BCL .NET apps and what your favorite tricks are.
```

## Tips

- Post **Tue–Thu morning Pacific**. r/dotnet has a more US/EU mix than r/commandline.
- Lead the post with **two screenshots**: the wallpaper-mode screenshot AND the `dotnet build` output showing 0 warnings (proves the zero-deps claim).
- Linkbait the comment section by asking a real question at the bottom ("anyone else here ship pure-BCL .NET? favorite tricks?"). r/dotnet rewards engagement.
- Be ready for "why not use Spectre.Console / Terminal.Gui?" — answer: those are great libraries, the design rule here is no packages at all, and the wallpaper embedding isn't something either of them does anyway.
- Be ready for "WorkerW is fragile" — yes, that's true, the project handles Win+D and Alt+Tab edge cases in `DesktopFormBase.cs`. Link directly.

## Pitfalls

- Don't title it as "yet another TUI" — the spreadsheet-as-wallpaper is the differentiator, lead with that or the zero-NuGet angle.
- Don't cross-post to r/csharp the same day. Stagger by 3–5 days.
- Don't include affiliate links or sponsorship CTAs. r/dotnet auto-removes.
