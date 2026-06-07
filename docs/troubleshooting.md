# Troubleshooting

Quick fixes for the most common QuickSheet issues.

## Build & install

### `dotnet build` fails with "SDK not found" or version error

QuickSheet targets .NET 9. Check your version:

```bash
dotnet --version   # needs 9.x
```

If you have an older SDK, install .NET 9 from [dotnet.microsoft.com](https://dotnet.microsoft.com/download/dotnet/9.0).

### Build succeeds but nothing appears on screen

Make sure you're running in **desktop mode** (the default):

```bash
dotnet run -c Release --project ExcelConsole.csproj
```

On Linux, the window sits behind all other windows — minimize everything
(Super+D on most DEs) to see the grid. On Windows the grid embeds inside
the WorkerW desktop layer, also behind open windows.

## Linux / X11

### "Cannot open X11 display. Is DISPLAY set?"

QuickSheet needs an X11 session. Common causes:

| Cause | Fix |
|-------|-----|
| Running over SSH without X forwarding | `ssh -X host` or use a VNC/RDP session |
| `DISPLAY` not set | `export DISPLAY=:0` (or `:1` for secondary) |
| Running inside a container without X socket | Mount `/tmp/.X11-unix` and set `DISPLAY` |

### "Warning: QuickSheet requires an X11 session"

You're on Wayland. QuickSheet's `_NET_WM_WINDOW_TYPE_DESKTOP` hint is X11-only.

**Options:**

1. Log out, pick **"GNOME on Xorg"** (or your DE's X11 session) at the login
   screen, log back in.
2. On GNOME, switch system-wide: edit `/etc/gdm3/custom.conf` and set
   `WaylandEnable=false`, then reboot.
3. XWayland fallback — QuickSheet attempts this automatically, but the window
   may not embed as a true desktop layer on all compositors.

See [issue #3](https://github.com/cemheren/QuickSheet/issues/3) and
[docs/wayland-investigation.md](wayland-investigation.md) for the full analysis.

### Font looks wrong or boxes appear

QuickSheet asks Xft for a monospace font. If none is installed:

```bash
# Debian/Ubuntu
sudo apt install fonts-dejavu-core
# Fedora
sudo dnf install dejavu-sans-mono-fonts
```

Then restart QuickSheet.

## Windows

### Grid doesn't appear behind icons

The desktop WorkerW embedding can fail if another program (e.g. Wallpaper
Engine, TranslucentTB, Lively Wallpaper) is already controlling the WorkerW
layer. Close the conflicting program and restart QuickSheet.

### Window flashes then vanishes

QuickSheet hides its console window on startup. If it crashes immediately,
run from a terminal to see the error:

```powershell
dotnet run --project ExcelConsole.csproj
```

Check the output for missing DLLs or .NET version mismatches.

## Extensions

### Extension cell shows prefix text instead of output

1. Confirm the extension is installed: `dotnet run --project ExcelConsole.csproj -- --list-extensions`
2. Check spelling — the prefix must match the manifest exactly (e.g. `wthr:` not `weather:`).
3. Look for stderr output in the terminal: `[ext:<name>:stderr]` lines indicate the extension process is crashing.

### "ext: github:user/repo" doesn't install

- **Network**: `git clone` must succeed. Check your internet connection and proxy/firewall settings.
- **GitHub auth**: public repos need no auth; private repos need `gh auth login` first.
- **Disk**: the extension clones to `<config>/extensions/<owner>-<repo>/`. Make sure the directory is writable.

### Extension output leaking into CSV

This was fixed in v0.28.0 (issue #156). Extension cell values are stored in an
ephemeral overlay — `SaveToCsv` writes only the original `ext:` prefix text.
If you're on an older version, update to latest.

## Data & CSV

### Autosave location

| OS | Path |
|----|------|
| Windows | `%APPDATA%\ExcelConsole\autosave.csv` |
| Linux | `$XDG_CONFIG_HOME/ExcelConsole/autosave.csv` (default: `~/.config/ExcelConsole/autosave.csv`) |

To use a specific file: `dotnet run --project ExcelConsole.csproj -- mydata.csv`

### CSV looks garbled after editing in Excel

Excel may save with BOM encoding or semicolon delimiters (locale-dependent).
QuickSheet expects UTF-8 CSV with comma delimiters. Re-save as
**"CSV UTF-8 (Comma delimited)"** in Excel's Save As dialog.

### Lost data after a crash

QuickSheet autosaves every 5 seconds. Check the autosave path above — your
data is likely there. If you were using a named CSV, the last save is your file.

## Headless / export modes

### "Input CSV not found"

The CSV path must come **before** the `--export-*` flag:

```bash
# ✅ Correct
dotnet run --project ExcelConsole.csproj -- data.csv --export-md output.md

# ❌ Wrong
dotnet run --project ExcelConsole.csproj -- --export-md output.md data.csv
```

### Piping from stdin

Use `-` as the input path:

```bash
cat data.csv | dotnet run --project ExcelConsole.csproj -- - --export-md -
```

## Still stuck?

[Open an issue](https://github.com/cemheren/QuickSheet/issues) with:
- Your OS and .NET version (`dotnet --info`)
- The exact command you ran
- Any error output from the terminal
