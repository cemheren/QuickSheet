# Persona 1 — On-call SRE / DevOps (second-monitor desktop)

Status: **done** (v2, desktop-mode focus) · Last revised: 2026-05-16

## TL;DR (5 lines)

- Audience = SRE/DevOps with a second monitor that currently runs a Grafana TV, Slack, or nothing. We want to *replace that surface* with a QuickSheet wallpaper.
- They already accept "ambient dashboard behind everything" as a workflow. The pitch is "your wallpaper IS that dashboard, no TV mode, no browser, survives reboots, every cell live."
- Highest-leverage glanceable cells: incident state (PagerDuty/OpsGenie), GitHub Actions run status, k8s pod health, cert expiry, AWS spend, queue depth.
- Where to seed: **r/unixporn** (with a desktop screenshot, not a terminal), **r/homelab**, **r/sre**, **r/devops**, HN with "I replaced my Grafana TV with a wallpaper" angle.
- Direct competition for this slot = Rainmeter on Windows, Conky on Linux. Neither has an extension protocol — that is our moat.

## 1. Profile

- Titles: SRE, DevOps engineer, platform engineer, on-call developer.
- US/EN headcount: ~600k SRE/DevOps-tagged on LinkedIn; ~2M+ devs who do ≥25% ops.
- Hardware: 2–3 monitors typical. Primary = IDE. Secondary = dashboards/Slack/terminal. **The secondary monitor is our seat.**
- Buying brain: "Can I keep this glanceable while I keep typing in the IDE?" If yes → installed forever.

## 2. Why their desktop is wasted

What currently lives behind their windows on the second monitor:

- A static wallpaper (most common). Wasted entirely.
- A maximised Grafana / DataDog browser tab. Burns RAM, refresh-spins, focus-steals.
- A `tmux` session with `watch kubectl` panes. Owns a whole terminal; not glanceable while editing config.
- Slack channels. Anxiety-inducing, low signal density.
- A spotify window. Honest.

What QuickSheet replaces: the second-monitor browser-tab dashboard. **Same surface, lower friction, transparent so windows still float over it, click-anywhere scratchpad for "stop and check this" notes.**

## 3. Glanceable data they actually want behind windows

Concrete cells, in priority order:

1. **Incident state row** — N cells, one per service, red if paging.
2. **CI status grid** — N cells, one per repo, green/yellow/red for last workflow run.
3. **Cluster health** — pod count, restart count, nodes notReady. Colour-coded.
4. **Cert expiry strip** — N domains, "12d", "94d", "EXPIRED" in red.
5. **Cost row** — yesterday's AWS/GCP/Azure spend; arrow vs 7-day avg.
6. **Queue depth** — SQS / RabbitMQ / Kafka lag counters.
7. **On-call rota** — "you are on-call until Thu 18:00."
8. **A scratch column** — typed notes for "thing to remember after standup," autosaved.

All of these must update *in place* without stealing focus and *without growing the grid*.

## 4. Candidate extensions / desktop-only features (ranked)

| Rank | Item                                   | Type    | Cost | Hit prob | Why                                                                                |
|------|----------------------------------------|---------|------|----------|-------------------------------------------------------------------------------------|
| 1    | `gha:` GitHub Actions status           | ext     | low  | high     | Free API, near-zero auth, half the audience uses GH. One green/red cell per repo. |
| 2    | `incident:` PagerDuty / OpsGenie       | ext     | med  | high     | Read-only /incidents endpoint. One red cell during an incident = sold.            |
| 3    | `k8s:` already exists — promote it     | docs    | low  | high     | Wallpaper-mode demo screenshots. The ext exists; no one's seen it on a desktop.    |
| 4    | **Cell staleness dimming**             | feat    | low  | high     | Cells not refreshed in N seconds dim to 30% alpha. Trust signal for "is this live?"|
| 5    | **Value-driven cell colour**           | feat    | low  | high     | `c?:>500=red,>200=yellow,*=green: 423` — colour reacts to current cell value.     |
| 6    | `prom:` Prometheus one-shot            | ext     | low  | med-high | Curl + label parse. PromQL pastes directly.                                       |
| 7    | `aws-cost:`                            | ext     | med  | med-high | `aws ce` CLI wrap; 1h cache. High WOW for cost-paranoid teams.                    |
| 8    | **Per-row blink-on-change**            | feat    | low  | med      | A subtle 1s highlight when a cell value changes. Catches eye without sound.       |
| 9    | `loki:` / log tail                     | ext     | high | med      | Streams via `i:` infra; harder for ambient (logs scroll fast).                    |
| 10   | `gh-review:` review queue              | ext     | low  | low-med  | `ghpr:` already covers it; refine instead of duplicating.                          |

Note: features 4, 5, 8 are **desktop-only ergonomics** — they are why someone keeps the wallpaper on for more than a day.

## 5. Where they hang out (desktop-customization first)

- **r/unixporn** — their *desktop* community. Screenshot post with the grid behind a code editor + a Slack window = the right shape of post here. Title pattern: `[Linux/Hyprland] Interactive spreadsheet wallpaper with live k8s + GH Actions`.
- **r/Rainmeter** — Windows side. Pitch as "Rainmeter widgets but each widget is a CSV cell with a JSON-lines extension protocol; install in one line."
- **r/homelab** + **r/selfhosted** — overlap. They love wallpaper dashboards.
- **r/sre / r/devops / r/kubernetes** — profession side. Post comes second, after the desktop-rice post earns a screenshot.
- **HN** — "Show HN: I replaced my Grafana TV with a desktop wallpaper" or "Show HN: a wallpaper-grid for on-call engineers."
- **People to be visible to:** rice-curators on r/unixporn; Liz Fong-Jones, Charity Majors, Julia Evans for the profession side. They share interesting *desktops* readily.

## 6. Discoverability hooks

Hero image = **full screen, three windows open (VSCode, Slack, browser), QuickSheet grid visible behind them with one red incident cell and a green CI strip**. Not a terminal.

Headlines that land:

- "[Linux/Hyprland] My wallpaper is an interactive spreadsheet that shows incidents + CI"
- "I replaced my second-monitor Grafana TV with a wallpaper grid (zero deps, open source)"
- "Rainmeter is great but I wanted CSV-first: built a wallpaper-grid with an extension protocol"

Avoid: TUI screenshots, terminal-only demos, lazygit/k9s comparisons.

## 7. Implications (queue these)

1. **Build `gha:` extension next Bucket F run.** One green/red cell per workflow run. Hero screenshot material.
2. **Bucket E: cell staleness dimming + value-driven cell colour.** These are the two features that make the wallpaper *believable* for SREs. Without them, cells could be lying.
3. **Bucket C: r/unixporn screenshot post.** Drafted only after `gha:` + colour rules ship. Title with distro tag, image = full desktop, body = "here's the install one-liner; here's the protocol."

Cross-link: [homelab.md](homelab.md) — heavy overlap; share the wallpaper-rice post angle.
