# Copilot Instructions for QuickSheet

## Build & Run

```bash
dotnet build ExcelConsole.csproj          # Build
dotnet run --project ExcelConsole.csproj  # Console mode (TUI spreadsheet)
dotnet run --project ExcelConsole.csproj -- --desktop          # Desktop mode (replaces wallpaper)
dotnet run --project ExcelConsole.csproj -- --desktop data.csv # Desktop mode with a CSV file
dotnet run -c Release --project ExcelConsole.csproj -- --desktop  # Release (recommended for desktop)
```

There are no tests or linters configured.

## Architecture

QuickSheet is a .NET 9 spreadsheet that runs in two modes:

- **Console mode** — `SpreadsheetApp` drives a TUI in the terminal using `System.Console`.
- **Desktop mode** (`--desktop`) — Embeds a transparent grid as the desktop wallpaper via platform-specific hosts.

### Layer separation

- **`GridManager`** — Pure data layer. Manages the cell grid, CSV I/O, column sums, row products. Has zero platform or UI dependencies. All UI layers share this.
- **`SpreadsheetApp`** — Console/TUI layer. Handles keyboard input, rendering, search, autosave.
- **`Platform/IDesktopHost`** — Interface for desktop mode. Each platform implements `Run(csvPath)` and `Dispose()`.
- **`Platform/Windows/`** — WinForms-based desktop host. Uses Win32 interop (`NativeMethods.cs`) to embed behind desktop icons.
- **`Platform/Linux/`** — Raw X11 P/Invoke implementation (`libX11.so.6`, `libXft.so.2`). Uses `_NET_WM_WINDOW_TYPE_DESKTOP`. Requires X11 session, not Wayland.

### Cross-platform build

The `.csproj` uses OS-conditional TFMs and compiler defines:
- Windows → `net9.0-windows`, `UseWindowsForms`, `PLATFORM_WINDOWS` define
- Linux → `net9.0`, `PLATFORM_LINUX` define
- Platform-specific source files under `Platform/Windows/` and `Platform/Linux/` are excluded from compilation on the other OS via `<Compile Remove>` directives.

Use `#if PLATFORM_WINDOWS` / `#elif PLATFORM_LINUX` guards for any platform-conditional code in shared files (see `Program.cs`).

## Key Conventions

- **Zero external NuGet dependencies** — Intentional policy to avoid supply chain risk. All platform interop is done via raw P/Invoke. Do not add NuGet packages.
- **Cell prefixes drive behavior** — `r: ` prefix makes a cell a runnable command; URLs are auto-detected as hyperlinks.
- **Autosave** — Console mode autosaves to `%APPDATA%/ExcelConsole/autosave.csv`. Desktop mode autosaves every 5 seconds to its CSV.
- **CSV as persistence format** — All data is stored as CSV. `GridManager.SaveToCsv` preserves rows/columns beyond the visible grid when rewriting files.
- **Namespace** — Everything is under the `ExcelConsole` namespace (with `ExcelConsole.Platform.*` for platform code).
