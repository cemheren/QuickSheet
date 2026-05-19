# Contributing to QuickSheet

Thanks for your interest! QuickSheet is a zero-dependency .NET 9 spreadsheet that runs in the terminal *and* as a desktop wallpaper. Contributions are welcome — from bug reports to new extensions.

## Quick start

```bash
git clone https://github.com/cemheren/QuickSheet.git
cd QuickSheet
dotnet build ExcelConsole.csproj
dotnet run --project ExcelConsole.csproj                          # TUI mode
dotnet run --project ExcelConsole.csproj -- --desktop             # Wallpaper mode (Windows/Linux X11)
dotnet run -c Release --project ExcelConsole.csproj -- --desktop  # Release build (recommended)
```

Requires the [.NET 9 SDK](https://dotnet.microsoft.com/download/dotnet/9.0). Nothing else — zero NuGet packages.

## Ground rules

Hard rules that keep the project lean. Please follow them in PRs:

| Rule | Why |
|------|-----|
| **Zero NuGet dependencies** | Supply-chain safety. All native interop (X11, WinForms, ConPTY) is hand-written P/Invoke. If you think you need a package, ask in an issue first — there's almost always a stdlib path. |
| **CSV is the only persistence format** | No JSON state, no SQLite, no sidecars. Users can open the file in Excel or vim. |
| **Cross-platform conditional compilation** | Platform code lives under `Platform/Windows/` or `Platform/Linux/`. The `.csproj` excludes the other OS via `<Compile Remove>`. In shared files use `#if PLATFORM_WINDOWS` / `#elif PLATFORM_LINUX`. |
| **Don't break existing behavior** | Keybindings, CLI flags, CSV format, autosave path — keep them stable. Additive changes are the easiest to land. |

## What's helpful

### 1. Bug reports

Include OS, terminal emulator, and the smallest CSV that triggers the issue. Desktop mode bugs should note screen resolution and whether you're on X11 or Wayland.

### 2. New cell prefixes

The prefix system (`i:`, `r:`, `s:`, `c:color:`, `ext:`) is QuickSheet's killer feature. New ones are easy to add — see `CellPrefix.cs`. Ideas: markdown render, mini-charts, live data fetch.

### 3. Extensions (easiest way to contribute!)

Extensions are **separate repos** that communicate via JSON-lines on stdin/stdout. There are already **40+ extensions** covering weather, finance, DevOps, productivity, and more.

**To create a new extension:**

1. Create a repo named `quicksheet-<thing>` (or `quicksheet-<thing>-ext`)
2. Add a `quicksheet-extension.json` manifest:
   ```json
   {
     "name": "my-extension",
     "version": "1.0.0",
     "prefix": "myext",
     "description": "What it does",
     "entry": "dotnet run --project MyExt.csproj",
     "minProtocolVersion": 1
   }
   ```
3. Implement the protocol:
   - Receive `{"type":"init"}` → reply `{"type":"register","name":"...","prefix":"...","version":"..."}`
   - Receive `{"type":"activate","id":"...","params":[...],...}` → reply `{"type":"write","id":"<same-id>","cells":[{"r":0,"c":0,"v":"..."},...]}`
4. Users install with: `ext: github:yourname/quicksheet-myext`

See the [extensions directory](docs/extensions.md) for all existing extensions and the [weather extension](https://github.com/Deskworks/quicksheet-weather) for a minimal reference.

> **Tip:** Read from `params` array (not `cells`) in activate messages. The `anchor` field gives cell positioning. Cache network responses with appropriate TTLs.

### 4. Wayland support

The Linux desktop path uses raw X11 P/Invoke and prints a warning under Wayland. Real Wayland support via `wl_surface` is an open challenge — see [#3](https://github.com/cemheren/QuickSheet/issues/3).

### 5. Screenshots & demo GIFs

Especially for desktop mode, themes, sparklines, and extension showcases. Great for the README and [landing page](https://cemheren.github.io/QuickSheet).

## Architecture at a glance

```
GridManager          — Pure data layer (grid, CSV I/O, Σ/Π). No UI deps.
SpreadsheetApp       — Console TUI (keyboard, render, search, autosave)
Platform/IDesktopHost — Desktop mode interface: Run(csvPath) + Dispose()
Platform/Windows/    — WinForms host, Win32 WorkerW embedding
Platform/Linux/      — Raw X11 + Xft P/Invoke, _NET_WM_WINDOW_TYPE_DESKTOP
Features/IMode.cs    — Modal input interface (search, goto, help overlay)
Extensions/          — Extension manager, installer, JSON-lines protocol
CellPrefix.cs        — Prefix parser (i:, r:, s:, c:color:, ext:, URLs)
```

## PR style

- One focused change per PR.
- Run `dotnet build ExcelConsole.csproj` before pushing — 0 warnings, 0 errors.
- Conventional Commits in the title (`feat:`, `fix:`, `docs:`, `chore:`).
- If you touch shared cross-platform code, note what you tested on (Windows / Linux / both).
- Desktop mode changes should be verified in wallpaper mode, not just TUI.

## License

By contributing, you agree your contributions are licensed under the same [MIT License](LICENSE) as the project.
