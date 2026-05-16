# Persona 8 — Homelab / sysadmin hobbyists / selfhosters

Status: **done** · Last revised: 2026-05-16

## TL;DR (5 lines)

- The single most-online persona in the slate. r/homelab (~800k), r/selfhosted (~400k), r/HomeServer, r/unixporn (~400k) — they *post their setups voluntarily and frequently.*
- Already running 5–30 self-hosted services (Plex, Pi-hole, Home Assistant, Sonarr/Radarr, Nextcloud, Grafana). Already familiar with the wallpaper-dashboard concept (Conky, Rainmeter, Polybar, eww/yad widgets).
- Highest-leverage wallpaper cells: per-service health row, Pi-hole stats, Plex now-playing, Home Assistant entities, disk-usage strip, IP/external-IP cell, UPS battery state, certificate expiries.
- Where to seed: **r/selfhosted** (highest signal-to-noise), **r/homelab**, **r/unixporn** (rice screenshot), Awesome-Selfhosted GitHub list. Plus r/HomeNetworking and r/HomeAssistant.
- Direct competition: Homepage (gethomepage.dev, ~17k stars), Dashy (~17k), Heimdall, Organizr — but **all are browser dashboards in a tab.** Nobody is on the actual wallpaper. Conky/Rainmeter come closest but have no extension protocol.

## 1. Profile

- Hardware: dedicated homelab server (NAS, Proxmox cluster, mini-PC stack, Pi cluster) + a daily-driver desktop/laptop. Often dual or triple monitor at the daily-driver.
- Skill level: high. They are sysadmin-grade, write their own scripts, contribute to open-source.
- Buying brain: "Can I selfhost this? Does it have an API? MIT license?" Allergic to SaaS, telemetry, anti-features.
- Cultural identity: anti-cloud, anti-subscription. The homelab itself is the hobby.
- Time-of-day: peak posting Sunday evening (post-tinkering review threads). r/homelab "Battlestation Sunday" is a thing.

## 2. Why their desktop is wasted

What lives there now (deeply ironic for this persona):

- Homepage.io / Dashy dashboard in a browser tab they have to focus to.
- Grafana / Prometheus tab they have to focus to.
- Plex / Jellyfin web UI in another tab.
- Maybe Conky widgets (Linux) or Rainmeter skins (Windows) — but limited to text scraping, no extension protocol.

What QuickSheet replaces: the *intent* of having a Homepage.io tab open. With wallpaper-mode the dashboard is on-screen permanently — not hidden in a browser tab. **For this audience specifically, the value-prop is "your Homepage.io dashboard but on the wallpaper, and you can write your own widgets in 50 lines of any language."** Conky users will immediately understand; Rainmeter users will immediately understand. The pitch is *one sentence*.

## 3. Glanceable data they actually want behind windows

1. **Service health grid** — N rows, one per service (Plex, Pi-hole, Nextcloud, Jellyfin, …). Green dot if up, red if down. *The* homelab dashboard cell.
2. **Pi-hole stats** — queries blocked today, percent blocked, top blocked domain.
3. **Plex / Jellyfin now-playing** — username, title, playback %.
4. **Home Assistant entities** — temperature in living room, garage door state, lights on count.
5. **Disk usage row** — per-pool/per-disk % used, red if >85 %.
6. **External IP cell** — for the ISP-IP-keeps-changing crowd. Watch dynamic DNS.
7. **Certificate expiries** — Let's Encrypt domains, days remaining. (`tls:` already shipped.)
8. **UPS battery state + runtime** — APC NUT or apcupsd reader.
9. **Speedtest cell** — `i: speedtest-cli --simple` once an hour.
10. **Server uptime + load** — `sysmon:` already shipped — re-pitch for this audience.

## 4. Candidate extensions / desktop-only features (ranked)

| Rank | Item                                       | Type | Cost | Hit prob | Why                                                                            |
|------|--------------------------------------------|------|------|----------|---------------------------------------------------------------------------------|
| 1    | `health:` HTTP probe many endpoints        | ext  | low  | high     | Reads `services.csv` (name, url, expected status); returns up/down row.        |
| 2    | `pihole:` admin API stats                   | ext  | low  | high     | Pi-hole has a free /api endpoint. Every homelabber runs it.                    |
| 3    | `plex:` now-playing                         | ext  | med  | high     | Plex token auth; XML API. Massively-deployed.                                  |
| 4    | `hass:` Home Assistant entity              | ext  | med  | high     | Long-lived token; REST API. Smart-home crowd huge.                             |
| 5    | `sonarr:` / `radarr:` queue                 | ext  | low  | med-high | *arr-stack queue + downloads. Selfhosted audience overlap massive.            |
| 6    | `proxmox:` VM/CT count + cluster status     | ext  | med  | med-high | Proxmox API auth; gives "5/5 VMs running" cell.                                |
| 7    | `dyndns:` external IP watch                 | ext  | low  | med      | Single api.ipify.org call; flag when IP changes.                                |
| 8    | `ups:` apcupsd / NUT battery state          | ext  | med  | med      | Local socket on Linux; powershell on Windows. Useful but niche.                |
| 9    | `truenas:` pool + disk health               | ext  | med  | med      | TrueNAS REST API. Targets the NAS-heavy subset.                                |
| 10   | `speedtest:` already shipped via `i:`        | docs | —    | —        | Just document the pattern.                                                      |
| —    | `k8s:`, `docker:`, `portck:`, `tls:`, `sysmon:`, `mxck:` *(shipped)* | — | — | — | Reframe the cluster as the homelab bundle.                                  |

The wallpaper-only feature that matters most here: **value-driven cell colour** (7/7 personas now confirmed). Plus **cell staleness dimming** — homelabbers care intensely about "is this metric current or stale."

## 5. Where they hang out (this persona is *the* online audience)

- **r/selfhosted (400k)** — highest signal-to-noise. Tool authors regularly post here themselves; the audience welcomes new MIT/selfhosted tools.
- **r/homelab (800k)** — bigger but more noise. "Battlestation Sunday" thread is the screenshot moment.
- **r/unixporn (400k)** — rice screenshots. A homelab-themed QuickSheet wallpaper screenshot fits exactly here.
- **r/HomeNetworking (400k), r/HomeServer, r/Proxmox, r/HomeAssistant (350k), r/jellyfin (50k), r/PleX (200k)** — long tail of vertical-specific subs.
- **selfh.st newsletter / Ethan Sholly** — *the* tastemaker newsletter for selfhosted tools. A feature here = several hundred installs overnight.
- **awesome-selfhosted** GitHub list (~200k stars on the list itself) — PR adding QuickSheet is a Bucket D action.
- **Homepage (gethomepage.dev), Dashy, Heimdall** maintainers — friendly to adjacent tools.
- **Discords:** r/homelab Discord, r/selfhosted Discord, HomeAssistant Discord.
- **Podcasts:** Self-Hosted Show (Jupiter Broadcasting / Chris Fisher), Linux Unplugged, 2.5 Admins, Late Night Linux.
- **People to be visible to:** Ethan Sholly (selfh.st), Chris Fisher (Self-Hosted podcast), Wendell from Level1Techs, Jeff Geerling (homelab YouTube). High-trust signal-boosters for this niche.

## 6. Discoverability hooks

Hero image = **a Linux desktop with i3/Hyprland + a Plex window + a terminal showing `htop`. Wallpaper behind: QuickSheet grid with green/red service-health dots, Pi-hole stats ("4,837 queries blocked, 32 %"), Plex now-playing, disk usage row, external IP cell.**

Headlines that land:

- "[r/selfhosted] Replaced Homepage.io with a wallpaper-grid dashboard (zero deps, MIT, extension protocol)"
- "Conky on steroids: an interactive CSV grid that lives on your wallpaper"
- "[r/unixporn] Hyprland + wallpaper dashboard for my homelab"
- "I built a Homepage.io alternative that runs as your desktop wallpaper"

The "Homepage.io alternative that doesn't live in a tab" framing is the strongest hook. Homepage has 17k stars — *one mention by an existing Homepage.io user comparing the two is worth more than any other action.*

Avoid: terminal-only screenshots, anything in the headline that implies "TUI." This audience runs GUIs and TUIs both; we want them to picture the wallpaper.

## 7. Implications (queue these)

1. **Build `health:` HTTP-probe extension** (Bucket F, post-research). The single most-leverage extension for this persona — reads a CSV of `name,url,expected_status`, fills a row of green/red dots. The screenshot writes itself.
2. **Build `pihole:` extension** second. Lowest cost, near-universal in homelabs, immediate "wait my wallpaper shows my Pi-hole stats?" reaction.
3. **Bucket A: `docs/for-homelab.md` + `examples/homelab-dashboard.csv` + `examples/selfhosted-services.csv`.** A complete starter sheet with `health:` + `pihole:` + `tls:` + `sysmon:` + `dyndns:`. Image of this sheet = the entire r/selfhosted post.
4. **Bucket C: drafts/homelab-launch.md** — three posts: r/selfhosted, r/unixporn (Hyprland-themed rice), and an email pitch to Ethan Sholly (selfh.st). Save after `health:` ships.
5. **Bucket D: submit QuickSheet to `awesome-selfhosted` PR draft.** Save in `drafts/awesome-selfhosted.md`.

Cross-link:
- [developers-sre.md](developers-sre.md) — heavy overlap on `health`, staleness dimming, value-colour. SRE persona is "homelab at work."
- [traders.md](traders.md) — same multi-monitor + dashboard-culture overlap.
