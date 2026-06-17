# r/selfhosted — Draft Reddit Post

**Subreddit:** r/selfhosted  
**Flair:** Self-Hosted Alternatives (if available), or Tools / Software  
**Posting time:** Weekday, ~10 AM–12 PM ET (peak engagement per research)

---

## Title options (pick one)

1. `I turned my Linux desktop wallpaper into a live service-health dashboard — no browser, no Electron, just a .NET spreadsheet`
2. `Show r/selfhosted: QuickSheet — a wallpaper spreadsheet that monitors your homelab services with green/red dots`
3. `My wallpaper is now my Homepage alternative — QuickSheet puts a live health grid right on the desktop`

---

## Post body

Hey r/selfhosted,

I've been working on [QuickSheet](https://github.com/cemheren/QuickSheet) — a terminal spreadsheet (.NET 9, zero NuGet dependencies) that also runs embedded into your desktop wallpaper on Linux (X11) and Windows.

The use case for homelabbers: **your wallpaper becomes a live dashboard.** No browser tab to forget, no Electron app eating RAM. It sits behind your windows, always visible when you minimize everything, and autosaves to CSV every 5 seconds.

### The health extension

The piece I'm most excited about for this community is the [`health:` extension](https://github.com/Deskworks/quicksheet-health-ext). You point it at your services:

```
health: plex=https://plex.lan,pi-hole=http://pi.hole/admin, 4, 5
```

Or reference a `services.csv`:

```
health: ~/services.csv, 4, 10
```

And you get a row per service with ✓ / ⚠ / ✗ indicators, HTTP status codes, and latency. It's not Grafana — there's no alerting or history — but it's a quick "is everything up?" strip that's always on your screen.

### What else goes on the wallpaper

Extensions are standalone executables that speak JSON-lines over stdin/stdout. QuickSheet installs them from GitHub with a single cell:

```
ext: github:Deskworks/quicksheet-tls-ext
```

The homelab-relevant ones:

- **`tls:`** — days until each Let's Encrypt cert expires
- **`ping:`** — latency to any host
- **`docker:`** — container status (running/stopped/unhealthy)
- **`sysmon:`** — CPU, RAM, disk, uptime of the local box
- **`mxck:`** — MX record lookup (for when your self-hosted mail goes sideways)
- **`dns:`** — DNS resolution check

You can also write your own in ~50 lines of any language. The [extension protocol](https://github.com/cemheren/QuickSheet/blob/main/docs/extension-protocol.md) is just JSON-lines — no SDK needed.

### How it compares to Homepage / Dashy / Heimdall

It's **not** a replacement. Those are proper dashboards with widget ecosystems. QuickSheet is for people who want something lighter:

- No web server to run
- No YAML to configure (it's a spreadsheet — edit cells, autosave)
- No browser to open — it's on the wallpaper
- CSV is the data format — `git diff` your dashboard config

The tradeoff: no mobile access, no alerting, no pretty widgets. It's a grid with numbers and status dots.

### Setup

```bash
git clone https://github.com/cemheren/QuickSheet.git
cd QuickSheet
dotnet run -c Release --project ExcelConsole.csproj -- --desktop homelab.csv
```

Requires .NET 9 SDK. Linux needs X11 (Wayland not yet supported — [tracking issue](https://github.com/cemheren/QuickSheet/issues/3)).

There's a [homelab setup guide](https://github.com/cemheren/QuickSheet/blob/main/docs/for-homelab.md) if you want a walkthrough.

### What it's not

- Not production-grade monitoring. Use Prometheus/Grafana for that.
- Not a Homepage replacement for a household — no mobile dashboard.
- Not tested on Wayland yet.
- Not polished like commercial tools — it's a side project.

Would love feedback from people who actually run homelabs. Is a "wallpaper status strip" useful to you, or is the browser-tab approach good enough?

---

## First comment (post immediately after submission)

> **Tech details for the curious:**
>
> - .NET 9, compiles to a single binary, zero NuGet dependencies (intentional supply-chain decision)
> - Linux: raw X11 P/Invoke — sets `_NET_WM_WINDOW_TYPE_DESKTOP` to sit behind icons
> - Windows: WinForms embedded into the WorkerW wallpaper layer
> - Extensions are subprocesses — they can be written in any language (Python, Go, Rust, C#, bash script). Communication is JSON-lines over stdin/stdout
> - All config is CSV — no database, no JSON state files
> - MIT licensed
>
> The extension repo list: https://github.com/cemheren/QuickSheet#extensions--make-your-desktop-do-more
>
> Happy to answer questions about the X11 embedding approach — it was the gnarliest part of the project.

---

## Notes for posting

- r/selfhosted allows self-promo if you're upfront about being the dev. The post is transparent about tradeoffs.
- Don't oversell — the "what it's not" section is critical for this subreddit. They will call out any exaggeration.
- If someone asks about Wayland: point to issue #3, acknowledge it's a limitation, mention that wlr-layer-shell is the likely path.
- If someone asks about Docker deployment: QuickSheet runs on the desktop, not in a container. It's a different paradigm. But the `docker:` extension *talks to* Docker.
- Screenshot would massively help this post. User should capture a real wallpaper with the health grid before posting.
