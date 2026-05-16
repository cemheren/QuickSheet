# Persona 1 — Developers (refined: SRE / DevOps / platform engineers)

Status: **done** · Last revised: 2026-05-16

## TL;DR (5 lines)

- QuickSheet's current beachhead. Strongest sub-segment = SRE/DevOps/platform — on-call shifts, multi-cluster dashboards, ambient-display culture.
- They already use `--desktop` wallpaper conceptually (Grafana TVs, second-monitor `htop`/`k9s`).
- Pain points: dashboards in browsers steal attention; CLI is great per query, bad as an ambient display.
- Best-leverage extensions to build next: `incident:` (PagerDuty/OpsGenie), `gha:` (GitHub Actions runs), `prom:` (Prometheus PromQL one-shot), `aws-cost:`, `loki:`/`grep:`.
- Best community to seed: r/sre, r/devops, r/kubernetes, HackerNews "Show HN", lobste.rs, and the Console Newsletter.

## 1. Profile

- Job titles: SRE, DevOps engineer, platform engineer, infra engineer, "production engineer," cloud engineer, on-call developer.
- Headcount (US/EN): ~600k SRE/DevOps-tagged on LinkedIn; ~2M+ developers who spend ≥25 % of their day in ops.
- Day-shape: 2–4 hours in IDE, the rest split between terminal, browser tabs of dashboards, Slack/Teams, on-call paging.
- Pays attention to: tools that survive a `tmux` session, run on every machine, don't require a login, and don't add a SaaS.

## 2. Current toolchain

| Tool                    | Used for                                | Pain                                                          |
|-------------------------|------------------------------------------|----------------------------------------------------------------|
| Grafana / DataDog       | Metric dashboards in the browser         | Tab overload; not glanceable while typing in IDE.              |
| `k9s`, `lazygit`, `htop`| Terminal TUIs for a single domain        | Each owns a whole terminal; can't co-render four at once.      |
| Slack reminders         | "Check X in 30 min"                      | Noisy; not visible as a state.                                 |
| Tmux dashboards         | Pasted-together `watch` loops            | Brittle; everyone's is different; no shared dotfiles primitive.|
| 2nd monitor with Chrome | Live dashboards behind windows           | Static; tab refresh, browser RAM, breaks on sleep.             |

Big unmet need: **a low-fi, always-on, click-through, multi-domain status board that sits on the desktop and survives reboots.** That is, literally, QuickSheet's product.

## 3. Pain points QuickSheet could touch

1. **Glanceable on-call state.** Are any of my services paging? Pods crashing? Builds red?
2. **Multi-cluster context.** k8s nodes, container health, regional latency in one grid.
3. **Cost paranoia.** Daily AWS/GCP/Azure spend without opening Cost Explorer.
4. **Per-repo CI snapshot.** Which of my 12 repos has a failing pipeline right now?
5. **Personal SLA timers.** SLO burn, certificate expiry, queue depth.

## 4. Candidate extensions / features (ranked)

| Rank | Extension                  | Cost | Hit probability | Why                                                           |
|------|----------------------------|------|------------------|----------------------------------------------------------------|
| 1    | `gha:` GitHub Actions      | low  | high             | Free API; near-zero auth (PAT). Half the audience uses it.    |
| 2    | `prom:` Prometheus one-shot| low  | high             | Curl + label parse. PromQL strings paste straight in.         |
| 3    | `incident:` PagerDuty      | med  | high             | Read-only `/incidents` endpoint. Token in env var. High WOW.  |
| 4    | `aws-cost:`                | med  | high             | `aws ce` CLI wrap; cache 1 h. Lights up cost-conscious teams. |
| 5    | `loki:` / `grep:` log tail | high | med              | Streamed output (already has `i:` machinery).                  |
| 6    | `slo:` SLO burn-rate       | med  | med              | Composite of `prom:` + math; build after `prom:`.             |
| 7    | `cve:` CVE lookup (NVD)    | low  | med              | Dev-news ambient bar.                                          |
| 8    | `gh-review:` review queue  | low  | low–med          | Already covered by `ghpr`; refine instead.                    |

QuickSheet *features* (vs extensions) worth queuing for this persona:

- **Cell expiry / staleness highlight** — if `i:` or extension output hasn't refreshed in N seconds, dim the cell. SREs trust *time-stamped* dashboards, not stale ones.
- **`L:` loop visual** — already exists; add a tiny progress tick so they see "yes it just refreshed."
- **Per-cell color rules** — `if value > X color red` as a cell prefix. Status-board basic. Currently `c:color:` is static.

## 5. Where they hang out

- **Reddit:** r/sre, r/devops, r/kubernetes, r/programming, r/sysadmin, r/homelab (crossover with persona 8).
- **HN:** Show HN section consistently rewards "I built a TUI for X" if X is on-call / infra-flavoured.
- **Newsletters:** Console (console.dev), TLDR DevOps, SRE Weekly, KubeWeekly, Last Week in AWS.
- **Discords:** CNCF Slack (#k8s-novice, #sig-instrumentation), DevOps subreddit Discord.
- **Conferences:** SRECon, KubeCon (booths/swag won't apply, but talk recordings → demo blog).
- **People to be visible to (no @ spam):** Liz Fong-Jones, Charity Majors, Julia Evans, Cindy Sridharan, Will Larson — they signal-boost on-call-ergonomics tools they actually use.

## 6. Discoverability hooks

Headlines that consistently land for this segment on HN/Reddit:

- "Show HN: My desktop wallpaper is a Kubernetes dashboard (zero deps)"
- "Show HN: I replaced Grafana TVs with a 200-line TUI"
- "I track PagerDuty incidents on my wallpaper now — here's the protocol"

Lead with **screenshot first**, then the technical hook (zero deps, JSON-lines, X11/WorkerW), then the "you can install this in one cell" GIF. Avoid: enterprise-flavoured framing, ROI language, "platform for X."

## 7. Implications (queue these)

1. **Build `gha:` extension next Bucket F run.** Highest hit-prob on the ranked list. Aim for "show last 5 workflow runs across N repos with red/yellow/green."
2. **Bucket E: add cell staleness dimming.** If a cell hasn't been written-to by an extension/`L:`/`i:` in N seconds, render in a 30 %-alpha foreground. Trust signal.
3. **Bucket C draft: r/sre Show post.** Lead with k8s/PagerDuty screenshot (once that ext exists). Title: "I replaced my Grafana TV with a 200-line wallpaper grid."

Cross-link: see also [niche-communities.md](../niche-communities.md) for tone notes per subreddit.
