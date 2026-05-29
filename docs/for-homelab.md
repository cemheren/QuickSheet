# QuickSheet for homelabbers

If you're running Plex, Pi-hole, Jellyfin, Nextcloud, Sonarr/Radarr, Home Assistant, or any
self-hosted stack — **your desktop wallpaper can be the dashboard.** No browser tab. No
SaaS. Your data stays on your machine.

> Already familiar with Homepage, Dashy, Heimdall, Conky, or Rainmeter? Think "those, but
> on the actual wallpaper, talking JSON-lines to extensions you can write in 50 lines of
> any language."

## Five-minute setup

```bash
git clone https://github.com/cemheren/QuickSheet.git
cd QuickSheet
dotnet run -c Release --project ExcelConsole.csproj -- --desktop examples/homelab-dashboard.csv
```

Edit cells, autosaves to CSV every 5 seconds. Add it to your startup applications and the
grid is there every boot — *behind* every window, click-through-safe, never stealing focus.

## What goes on the wallpaper

The starter sheet at `examples/homelab-dashboard.csv` gives you a frame you can edit in
place. The columns:

| Column      | What it shows                                                      | Extension used                          |
|-------------|--------------------------------------------------------------------|-----------------------------------------|
| **Status**  | Per-service health — green dot if up, red if down                  | [`k8s`](https://github.com/Deskworks/quicksheet-k8s) / [`docker`](https://github.com/Deskworks/quicksheet-docker) |
| **Cert expiry** | Days remaining on each Let's Encrypt cert                       | [`tls`](https://github.com/cemheren/quicksheet-tls-ext) |
| **Network** | Latency to upstreams + open TCP ports                              | [`ping`](https://github.com/Deskworks/quicksheet-ping-ext) / [`portck`](https://github.com/Deskworks/quicksheet-portck) |
| **System**  | CPU / RAM / disk / uptime of the box running QuickSheet            | [`sysmon`](https://github.com/Deskworks/quicksheet-sysmon) |
| **DNS**     | MX records — useful when your selfhosted mail goes weird           | [`mxck`](https://github.com/Deskworks/quicksheet-mxck-ext) |

All six extensions install with a single `ext: github:Deskworks/quicksheet-<name>` cell.
QuickSheet clones the repo, starts the subprocess, and the prefix is live.

## Make it pretty

Cycle themes with **Ctrl+T**: Dark, Light, Nord, Solarized, Matrix, Dracula, Synthwave,
Gruvbox, Monokai, HotdogStand. Gruvbox or Nord fit the r/unixporn rice aesthetic best.

For colour-coded health, use the value-driven cell colour prefix:

```
c?: >85=red, >70=yellow, *=green: 42
```

Disk usage > 85 % renders red, > 70 % yellow, else green. Pair with `i: df --output=pcent /pool | tail -1 | tr -d '% '` for live disk readings.

## Why wallpaper > tab

- **Always visible** — no `Cmd+Tab` to "go check Homepage." It's on-screen when you
  unlock the laptop.
- **Survives reboots** — autosaved CSV, opens to the same grid.
- **No SaaS** — your `services.csv` is a plain file in your folder.
- **Per-cell hot-swap** — change a single cell, the dashboard reflects it. No re-build.
- **Extensions are scripts** — write a 50-line `quicksheet-myext` for your one specific
  service. JSON-lines over stdin/stdout, zero deps, any language.

## Recipe: a colour-coded service grid

1. One row per service.
2. Status cell: use the matching extension (`docker: <name>`, `k8s: <name>`, `ping: <host>`).
3. Wrap critical cells in `c?:` rules so they go red when bad.
4. Add a [`L: A10, 5m`](../README.md#loops) loop cell anywhere to re-fire updates on an
   interval if the extension doesn't self-refresh.

## More extensions to mix in

- [`hntop`](https://github.com/Deskworks/quicksheet-hntop) — HN front page in a strip.
- [`news`](https://github.com/Deskworks/quicksheet-news) — RSS feed reader (subscribe to
  selfh.st or your own homelab RSS).
- [`apistatus`](https://github.com/Deskworks/quicksheet-apistatus) — GitHub / Cloudflare /
  npm status pages.
- [`fx`](https://github.com/Deskworks/quicksheet-fx) — currency conversion (for the
  homelab-also-a-finance-nerd subset).

See the [full extension directory](extensions.md).

## Build your own in 50 lines

If you don't see an extension for the service you're running, the protocol is two message
types ([extension-protocol.md](extension-protocol.md)). Stdlib Python or Go is enough. We
recommend zero NuGet/pip/cargo dependencies in your extension — keep the supply chain
trivial, since `ext: github:user/repo` clones it onto every install.

## Share your setup

Post your wallpaper on r/selfhosted, r/homelab, or r/unixporn (the rice angle is real).
Open an issue if you want a new prefix officially listed.
