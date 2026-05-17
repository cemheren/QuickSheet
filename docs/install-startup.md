# Run QuickSheet on startup

Once you're used to the wallpaper grid, you'll want it back every login. Two
platforms, three approaches.

## Windows

The repo ships `scripts/quicksheet-startup.ps1` — a PowerShell launcher that
runs `dotnet run -c Release --project ExcelConsole.csproj -- --desktop` from
the repo root in a detached, hidden window.

### Option A: Startup folder shortcut (simplest)

1. Press <kbd>Win</kbd>+<kbd>R</kbd> → type `shell:startup` → Enter. The
   "Startup" folder opens.
2. Right-click → **New** → **Shortcut**. Paste:
   ```
   powershell -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\path\to\QuickSheet\scripts\quicksheet-startup.ps1"
   ```
   Replace `C:\path\to\QuickSheet` with where you cloned the repo.
3. Name the shortcut "QuickSheet". Done. Reboot to verify.

### Option B: Task Scheduler "At log on" (more reliable)

1. Open Task Scheduler → **Create Basic Task...**
2. **Name:** QuickSheet · **Trigger:** When I log on · **Action:** Start a program.
3. **Program/script:** `powershell`
4. **Arguments:** `-ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\path\to\QuickSheet\scripts\quicksheet-startup.ps1"`
5. **Start in:** `C:\path\to\QuickSheet`
6. Finish. Right-click the task → **Properties** → **Conditions** tab → uncheck
   "Start the task only if the computer is on AC power" if you're on a laptop.

### Optional: auto-update on startup

Append `-Update` to the script invocation. The launcher will:

1. `git pull --rebase --autostash origin main`
2. `dotnet build -c Release`
3. Only then launch.

If either git or build fails, **the script aborts and does not launch**. This
is deliberate: a broken push to `main` would otherwise leave you with no
wallpaper at boot.

```
powershell -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\path\to\QuickSheet\scripts\quicksheet-startup.ps1" -Update
```

## Linux (X11)

No script needed — a single XDG autostart `.desktop` file is enough.

```bash
mkdir -p ~/.config/autostart
cat > ~/.config/autostart/quicksheet.desktop <<EOF
[Desktop Entry]
Type=Application
Name=QuickSheet
Comment=Interactive spreadsheet wallpaper
Exec=dotnet run -c Release --project $HOME/code/QuickSheet/ExcelConsole.csproj -- --desktop
Path=$HOME/code/QuickSheet
X-GNOME-Autostart-enabled=true
Terminal=false
EOF
```

Replace `$HOME/code/QuickSheet` with your clone path. Log out and back in to
test.

Wayland note: `--desktop` mode currently requires X11
([issue #3](https://github.com/cemheren/QuickSheet/issues/3) tracks Wayland
support). If you're on Wayland, the autostart entry will print the warning and
fall back to TUI mode in a terminal — usable but no longer a wallpaper.

### Optional: auto-update on Linux

Edit the `Exec=` line to chain the update:

```
Exec=sh -c "cd $HOME/code/QuickSheet && git pull --rebase --autostash origin main && dotnet build -c Release ExcelConsole.csproj && dotnet run -c Release --project ExcelConsole.csproj -- --desktop"
```

Same trade-off as Windows: failure of git or build aborts the chain.

## macOS

Not supported yet — `--desktop` mode uses Win32 WorkerW on Windows and X11 on
Linux. macOS would need an NSWindow-based path. PRs welcome (see
[CONTRIBUTING.md](../CONTRIBUTING.md)). For now, run the plain terminal TUI
via the standard `Login Items` flow.

## Uninstall

- **Windows / Startup folder:** delete the shortcut from `shell:startup`.
- **Windows / Task Scheduler:** open Task Scheduler → right-click the task → Delete.
- **Linux:** `rm ~/.config/autostart/quicksheet.desktop`.

## See also

- [docs/tour.md](tour.md) — 60-second tour.
- [docs/for-homelab.md](for-homelab.md) — wallpaper dashboard for selfhosters.
- [docs/for-traders.md](for-traders.md) — wallpaper dashboard for traders.
