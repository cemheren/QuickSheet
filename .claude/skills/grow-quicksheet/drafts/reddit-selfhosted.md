# r/selfhosted post draft

**Subreddit:** r/selfhosted  
**Flair:** Self-Hosted Alternatives (or "Project Share" if available)

---

## Title

**I replaced my Dashy/Homepage dashboard with a spreadsheet wallpaper — it's weirdly practical**

---

## Body

I run a small homelab (Plex, Pi-hole, Nextcloud, couple of Docker stacks). I used Homepage for a while but kept wanting something simpler — no YAML to configure, no browser tab to keep open, and something I could just *type into* when I need to jot a quick note.

So I built [QuickSheet](https://github.com/cemheren/QuickSheet): a spreadsheet that *is* the desktop wallpaper. You click the wallpaper, you're typing in a cell. Autosaves every 5 seconds to a CSV.

The homelab angle: I use it as a service dashboard that lives on the desktop permanently — no tab, no browser, no alt-tab. Here's what my layout looks like:

- **Health column** — each service gets a cell with the [`health:`](https://github.com/cemheren/quicksheet-health-ext) extension prefix. It HTTP-probes the URL and shows a green/yellow/red dot + latency. One cell per service.
- **TLS expiry column** — `tls: myserver.local` shows days remaining on the cert. Goes yellow under 30d.
- **Launcher row** — `r: ssh nas` / `r: docker compose up -d` / `r: firefox http://plex.local:32400`. Multi-select cells, hit Enter, they all fire.
- **Notes area** — just plain text cells for "reboot nas tonight" or pasting a quick IP.

Extensions are just git repos that speak JSON-lines over stdin/stdout. `ext: github:cemheren/quicksheet-health-ext` in any cell, and the `health:` prefix is live. There's also [`docker:`](https://github.com/cemheren/quicksheet-docker), [`ping:`](https://github.com/cemheren/quicksheet-ping-ext), [`tls:`](https://github.com/cemheren/quicksheet-tls-ext), [`sysmon:`](https://github.com/cemheren/quicksheet-sysmon) for CPU/RAM/disk.

**What it is:**

- .NET 9, zero NuGet dependencies (entire supply chain = .NET SDK)
- Linux (X11) and Windows
- CSV is the only persistence format — your "dashboard config" is a CSV file
- Themes (Nord, Gruvbox, Dracula, etc.) — Ctrl+T to cycle
- Cell prefixes drive everything: `r:` = run command, `i:` = live subprocess output, `s:` = sparkline, `ext:` = install extension

**What it isn't:**

- Not a Homepage/Dashy replacement if you need widget tiles, icons, or a web-accessible dashboard
- Not Wayland-compatible yet (X11 only on Linux)
- Not "pretty" in the polished-webapp sense — it's a grid with text

If you're the kind of person who has a terminal open all day anyway and just wants a persistent scratchpad + status board on the wallpaper, it might click.

**Links:**
- GitHub: https://github.com/cemheren/QuickSheet
- 60-second tour: https://github.com/cemheren/QuickSheet/blob/main/docs/tour.md
- Homelab-specific guide: https://github.com/cemheren/QuickSheet/blob/main/docs/for-homelab.md
- Health extension: https://github.com/cemheren/quicksheet-health-ext

Happy to answer questions. This is a side project — I'm aware it's niche, but for the niche it serves, it's been really useful on my own machine daily.

---

## Posting notes

- **Best time:** weekday morning EST (Tue–Thu, 8–10 AM)
- **Tone:** practical, not promotional. Lead with the problem (dashboard fatigue, YAML config files, needing a browser tab open). Show the actual use case.
- **Expectations:** r/selfhosted is friendly to small projects if you're honest about scope. "This isn't a Homepage killer, but here's why it works for me" lands better than "look at my thing."
- **Avoid:** calling it "production-grade," claiming large user counts, comparing unfavorably to established tools. Acknowledge limitations upfront (X11 only, no Wayland, not web-accessible).
- **Follow-up comment:** Prepare a top-level reply with the full extension list relevant to homelabbers (docker, health, ping, tls, sysmon, mxck) and a quick "here's my actual CSV" snippet.

---

## Prepared first comment

> For the curious, here's the extensions I use in my homelab layout:
>
> | Cell content | What it does |
> |---|---|
> | `ext: github:cemheren/quicksheet-health-ext` | Install the health-check extension |
> | `health: http://plex.local:32400/web` | Green dot if 200, red if down |
> | `tls: nas.mydomain.com` | Days until cert expires |
> | `ping: 192.168.1.1` | Latency to router |
> | `docker: plex` | Container status |
> | `sysmon: cpu` | Live CPU percentage |
>
> Each extension is ~50 lines of code in any language. If you have a weird service you want to monitor, you can write your own in an afternoon. The protocol is just JSON-lines over stdin/stdout — [docs here](https://github.com/cemheren/QuickSheet#extensions--make-your-desktop-do-more).
