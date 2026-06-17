# Reddit r/selfhosted draft — QuickSheet as homelab dashboard

Drafted 2026-06-16. Subreddit (~400k subs) rewards concrete dashboards, status pages, and self-hosted alternatives. Lead with the homelab use case, not the spreadsheet abstraction.

> **User submits this manually. Do not post to Reddit.**

## Title

**I built a desktop-wallpaper dashboard for my homelab — it checks HTTP health, TLS certs, Docker containers, and ping latency without opening a browser**

(Alternative: `My homelab dashboard lives on the desktop wallpaper — no browser, no SaaS, no Docker compose for the dashboard itself`)

## Body

```
I run a handful of self-hosted services (Jellyfin, Nextcloud, Pi-hole, a few *arr apps) and got tired of keeping a browser tab open just to check if things were up. I wanted the status visible *all the time* — on the wallpaper, behind every window.

So I built QuickSheet. It's a terminal spreadsheet (.NET 9, zero dependencies) with a `--desktop` flag that embeds the grid as the actual desktop wallpaper. On Linux it uses `_NET_WM_WINDOW_TYPE_DESKTOP` on X11; on Windows it hooks into the WorkerW handle.

The homelab-relevant part: a small extension protocol (JSON-lines over stdin/stdout) lets you write 50-line checkers in any language. Extensions I use daily:

- `health: https://jellyfin.local, 200` → green dot if 200, red otherwise. Checks every loop cycle.
- `tls: nextcloud.example.com` → days until cert expiry. Turns yellow < 14 days, red < 7.
- `docker: jellyfin` → shows container status, CPU/mem, image tag.
- `ping: 1.1.1.1` → HTTP status code + latency in ms.
- `i: curl -s http://pi.hole/admin/api.php?summary | jq .ads_blocked_today` → inline command output, refreshes live.

Each extension is a separate GitHub repo — install with `ext: github:Deskworks/quicksheet-health` in any cell and it clones + starts the subprocess automatically.

The grid autosaves to CSV every 5 seconds. Same CSV opens in Excel, vim, whatever. No database, no YAML config, no Docker compose for the dashboard itself.

**What it looks like in practice:**

A few columns on the wallpaper: service name | health dot | cert days | container status | notes. Always visible. Click a cell to edit. Press Enter on a `r: ssh user@host` cell to open a terminal. Multi-select cells and fire them all at once to restart a stack.

**What it's NOT:**
- Not a Grafana/Prometheus replacement. No time-series, no alerting, no graphs (sparklines exist but they're toy-scale).
- Not a Homepage/Dashy/Heimdall replacement in the sense of "pretty bookmark grid with icons." It's a spreadsheet — it looks like a spreadsheet.
- Not Conky. Conky draws text overlays; this is an interactive grid with editable cells, runnable commands, and a real cursor.

**What it IS:**
- A CSV grid permanently on your desktop that can run health checks.
- An alternative for people who want ambient visibility without a browser tab.
- MIT licensed, zero NuGet dependencies, ~15k lines of C#. Clone → `dotnet build` → `dotnet run -- --desktop`.

Honest caveats: Linux needs X11 (no Wayland yet), the UI is a grid not a pretty dashboard, and it's a side project not a polished product. But it does the "is my stuff up?" job well enough that I stopped opening Uptime Kuma for quick checks.

Repo: https://github.com/cemheren/QuickSheet
Homelab guide: https://github.com/cemheren/QuickSheet/blob/main/docs/for-homelab.md
Starter CSV: `examples/homelab-dashboard.csv` in the repo.

Happy to answer questions about the extension protocol if anyone wants to write a checker for their own services.
```

## Flair

Use **Dashboard** or **Self-Hosted Alternatives** flair if available. Check subreddit flair options before posting.

## Tips for the actual post

- Post **Saturday or Sunday morning US time**. r/selfhosted is most active on weekends when people are tinkering.
- Attach a screenshot showing the wallpaper with health dots, TLS days, and docker status visible. The visual hook is "that's a wallpaper, not a browser."
- r/selfhosted LOVES "no Docker for the tool itself" — emphasize the `dotnet build` simplicity vs. deploying yet another container.
- Have a follow-up comment ready: "Here's my actual CSV if anyone wants to start from it" with the contents of `examples/homelab-dashboard.csv`.
- Mention Uptime Kuma / Homepage / Heimdall by name — the sub knows them, and positioning against something familiar is better than positioning against nothing.

## Pitfalls to avoid

- Don't call it a "spreadsheet" in the title. r/selfhosted doesn't care about spreadsheets. Lead with "dashboard" / "health checker."
- Don't oversell. It's a grid, not Grafana. The sub will roast inflated claims.
- Don't mention star count or ask for stars/upvotes.
- Don't post within 24 hours of posting to r/commandline or r/linux — looks like spam.
- Don't frame it as "I replaced Homepage" unless you actually did. Frame it as "an alternative approach for ambient monitoring."

## Why r/selfhosted

- ~400k subscribers, highly engaged. Dashboard posts regularly hit 200+ upvotes.
- The "is there a self-hosted alternative to X" question is the sub's bread and butter. QuickSheet fits as a lightweight, no-Docker status page.
- The homelab extension bundle (health, tls, docker, ping, sysmon) is exactly what this audience builds.
- Cross-link opportunity: if the post does well, link from QuickSheet README to the r/selfhosted discussion as social proof.

## Comparison table for comments

If asked "how does this compare to X?", use this:

| Feature | QuickSheet | Homepage/Dashy | Uptime Kuma | Conky |
|---------|-----------|---------------|-------------|-------|
| Runs on wallpaper | ✅ | ❌ (browser) | ❌ (browser) | ✅ (overlay) |
| Interactive cells | ✅ | ❌ | ❌ | ❌ |
| Health checks | ✅ (ext) | ❌ | ✅ | ❌ |
| TLS monitoring | ✅ (ext) | ❌ | ✅ | ❌ |
| Docker status | ✅ (ext) | ✅ | ❌ | ❌ |
| Run shell commands | ✅ | ❌ | ❌ | ✅ (limited) |
| No Docker needed | ✅ | ❌ | ❌ | ✅ |
| Pretty UI | ❌ (grid) | ✅ | ✅ | ❌ |
| Alerting | ❌ | ❌ | ✅ | ❌ |
