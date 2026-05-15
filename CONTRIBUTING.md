# Contributing to QuickSheet

Thanks for your interest. QuickSheet is a small side project, but contributions are welcome.

## Quick start

```bash
git clone https://github.com/cemheren/QuickSheet.git
cd QuickSheet
dotnet build ExcelConsole.csproj
dotnet run --project ExcelConsole.csproj                  # TUI mode
dotnet run --project ExcelConsole.csproj -- --desktop     # Wallpaper mode
```

Requires the .NET 9 SDK. Nothing else.

## Ground rules

A few hard rules that keep the project lean. Please follow them in PRs:

- **Zero NuGet dependencies.** The repo has no package references and is meant to stay that way. All native interop (X11, WinForms, ConPTY) is hand-written P/Invoke. If you need a feature that *feels* like it needs a package, ask in an issue first — there's almost always a stdlib path.
- **CSV is the persistence format.** No JSON state files, no SQLite, no app-data sidecars. The user can open the file in Excel or vim and edit it by hand.
- **Cross-platform conditional compilation.** Platform-specific code lives under `Platform/Windows/` or `Platform/Linux/`. The `.csproj` excludes the other OS via `<Compile Remove>`. In shared files use `#if PLATFORM_WINDOWS` / `#elif PLATFORM_LINUX`.
- **Don't break existing keybindings, CLI flags, CSV format, or autosave path.** Additive changes are the easiest to land.

## What's helpful

In rough priority order:

1. **Bug reports with repro steps.** Include OS, terminal, and the smallest CSV that triggers it if relevant.
2. **New cell prefixes.** The `i:`, `r:`, `s:`, `L:`, `ext:` prefixes are the project's killer feature. New ones (Markdown render, mini-charts, etc.) are easy to add — see `CellPrefix.cs`.
3. **Extensions.** Extensions live in their own repos and register a prefix with QuickSheet over JSON-lines. See the [Extensions section](README.md#extensions--make-your-desktop-do-more) and the [weather extension](https://github.com/cemheren/quicksheet-weather) for a minimal example.
4. **Wayland support.** The Linux desktop path currently uses raw X11 and prints a warning under Wayland. Real Wayland support is open.
5. **Screenshots and demo GIFs.** Especially for less-photogenic features (multi-select launching, search, sparklines).

## Tests

There's a small headless test runner at `tests/run.sh` that exercises the CLI paths (`--version`, `--help`, `--export-md`, `--list-extensions`) end-to-end without a TTY. Run it before pushing:

```bash
bash tests/run.sh
```

The same script runs in CI on every PR (`.github/workflows/build.yml`).

## PR style

- One focused change per PR.
- Run `dotnet build ExcelConsole.csproj` and `bash tests/run.sh` before pushing — both must be 0 errors / 0 failures.
- Conventional Commits in the title (`feat:`, `fix:`, `docs:`, `chore:`).
- If you touch shared cross-platform code, please add a brief note on what you tested on (Windows / Linux / both).

## License

By contributing, you agree your contributions are licensed under the same MIT license as the project.
