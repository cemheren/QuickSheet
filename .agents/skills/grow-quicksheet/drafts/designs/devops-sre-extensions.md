Agent completed. agent_id: devops-research, agent_type: research, status: completed, description: Research devops terminal tools, elapsed: 254s, total_turns: 0, duration: 250s

I now have everything I need. Let me compile the comprehensive research report.

---

# QuickSheet Extension Research Report: DevOps, SRE & Developer Productivity (2025–2026)

## Summary

**Three data-backed conclusions before the extension list:**

1. **The wallpaper angle is genuinely novel** — WTF terminal (`wtfutil/wtf`, the closest competitor) runs inside a terminal session. QuickSheet's always-visible, click-through-aware desktop layer is architecturally different and captures the **ambient monitoring** use case that no TUI tool has fully owned yet.

2. **The extension system is already production-ready** — The JSON-lines subprocess protocol (`ExtensionProtocol.cs`), `ext: github:user/repo` one-liner installs, and 20 existing extensions (`docs/extensions.md`) are exactly the right foundation. The research confirms the gaps are real: k8s, docker, GitHub PRs/Actions, HN, git status, and service status pages are all uncovered.

3. **The "viral" pattern for TUI tools is "replace a browser tab or app forever"** — k9s (33k★) replaced `kubectl get pods`. lazygit replaced SourceTree. `gh-dash` killed the GitHub notifications tab. The strongest QuickSheet extensions follow this same pattern: things you check 10-20x/day that don't need a full app.

---

## 1. Trending TUI/Terminal Tools in 2025–2026

**Verified from `github.com/topics/tui?o=desc&s=stars` and `topics/terminal?o=desc&s=stars`:**

| Tool | Stars | Language | Why it's trending |
|------|-------|----------|-------------------|
| **Textualize/rich** | ~50k | Python | Beautiful terminal formatting; framework others build on |
| **charmbracelet/bubbletea** | ~30k | Go | The dominant TUI framework; every new Go TUI uses it |
| **ratatui/ratatui** | ~24k | Rust | Rust TUI framework; exploding with the Rust wave |
| **sxyazi/yazi** | ~22k | Rust | File manager; viral for async I/O + beautiful UI |
| **Textualize/textual** | ~28k | Python | TUI apps that run in both terminal AND browser |
| **gitui-org/gitui** | ~18k | Rust | Fast git TUI (competitor to lazygit) |
| **wagoodman/dive** | ~48k | Go | Docker image layer explorer — DevOps staple |
| **dlvhdr/gh-dash** | ~8k | Go | GitHub dashboard in terminal — directly relevant |
| **darrenburns/posting** | ~10k | Python | API client in terminal (Postman killer) |
| **derailed/k9s** | **33.6k** | Go | Kubernetes TUI — most-starred K8s CLI tool |
| **wtfutil/wtf** | ~16k | Go | Terminal personal dashboard (closest QuickSheet competitor) |
| **jesseduffield/lazygit** | ~54k | Go | Git TUI — one of the all-time viral terminal tools |

**Key observation:** Rust and Go dominate new entries. Rust for raw performance and binary size (single binary = easy distribution), Go for ecosystem + quick compile. Extensions should support both since QuickSheet's protocol is language-agnostic JSON-lines.

---

## 2. What Makes Terminal Tools Go Viral — The 6 Patterns

Derived from studying k9s, lazygit, gh-dash, uptime-kuma, yazi, and WTF terminal:

### Pattern 1: "Replace a Tab" (strongest signal)
- **k9s** → replaced `kubectl get pods -A && kubectl describe pod X` (20 keystrokes → 0)
- **lazygit** → replaced opening SourceTree/GitHub Desktop
- **gh-dash** → killed visiting github.com/notifications
- **QuickSheet wallpaper angle**: `hntop:` replaces opening `news.ycombinator.com`. `ghpr:` replaces the GitHub notifications page. `apistatus:` replaces bookmarking 5 status pages.

### Pattern 2: Ambient / Always Visible > On-Demand
- Tools that are "always there" compound value over time. This is QuickSheet's moat. **Nothing in the TUI ecosystem lives on the wallpaper**. Rainmeter does, but it's Windows-only and requires XML config. QuickSheet is code-first.

### Pattern 3: Zero-Config Start
- k9s: `brew install k9s`, type namespace, done. No YAML.
- yazi: `cargo install yazi-fm`, no config.
- **QuickSheet extension equivalent**: `ext: github:cemheren/quicksheet-k8s` + `k8s: default, 3, 10` and it works with your existing `~/.kube/config`.

### Pattern 4: Beautiful Output (Screenshotable)
- People post screenshots on Reddit/Twitter. The ASCII sparklines (`s:` prefix) already work this way.
- Extensions that produce colored status indicators, progress bars, and emoji glyphs spread faster.

### Pattern 5: Local-First (Security-Sensitive Audience)
- k9s uses local kubeconfig. No cloud. SREs love this.
- QuickSheet: zero NuGet, CSV only, extensions are auditable git repos. Perfect trust story.

### Pattern 6: Solves "I Check This 20x/Day"
Git branch + status. PR count. Service health. K8s pod restarts. These are mental-overhead tasks that pollute every context switch.

---

## 3. DevOps/SRE Pain Points → QuickSheet Extension Ideas

### 🟢 Tier 1: "I Need This On My Wallpaper Right Now" (Ship These First)

---

### `quicksheet-k8s`
**Prefix:** `k8s:`  
**One-liner:** Live Kubernetes pod status table from your current kubeconfig context  
**Wallpaper angle:** SREs check `kubectl get pods -A` constantly. This makes it ambient — pod restarts and CrashLoopBackOff appear on the wallpaper without opening a terminal.  
**API/Data source:** Local `kubectl` binary + `~/.kube/config` — zero network, zero auth to set up  
**Implementation:** Shell out to `kubectl get pods -n <namespace> --no-headers -o custom-columns=...`, parse text into grid cells. Use cell color-coding: Running=green, CrashLoopBackOff=red, Pending=yellow (maps to extension `status` message).
**Example cell:**
```
k8s: default, 4, 12
```
Outputs a 4-column × 12-row grid: `NAME | STATUS | RESTARTS | AGE`  
**Feasibility: 5/5** — `kubectl` is already on every SRE's machine. Pure subprocess call.

---

### `quicksheet-docker`
**Prefix:** `docker:`  
**One-liner:** Live Docker container health grid from the local Docker socket  
**Wallpaper angle:** "Is my local postgres running? Did my dev compose stack come up?" — checked every dev session start.  
**API/Data source:** Docker Engine API via local Unix socket `/var/run/docker.sock` or named pipe `npipe:////./pipe/docker_engine`. HTTP GET `/v1.47/containers/json` — returns JSON, completely free, no credentials.  
**Implementation:** Direct socket HTTP call (no Docker CLI dependency), parse container `Names`, `State`, `Status`, `Ports`. 
**Example cell:**
```
docker: all, 4, 8
```
Outputs: `NAME | IMAGE | STATUS | PORTS`  
**Feasibility: 5/5** — Docker socket is always accessible, JSON API is stable and documented.

---

### `quicksheet-ghpr`
**Prefix:** `ghpr:`  
**One-liner:** GitHub PRs assigned to you or waiting for your review — right on the wallpaper  
**Wallpaper angle:** The #1 thing developers check on GitHub is "do I have PRs to review?" and "is my PR approved?". This eliminates the browser tab entirely.  
**API/Data source:** GitHub REST API `GET /repos/{owner}/{repo}/pulls` — 60 req/hr unauthenticated; OR delegate to `gh pr list --json number,title,reviewDecision,author` (uses existing `gh auth` token, no new credential setup).  
**Implementation:** Best pattern: call `gh pr list --json number,title,reviewDecision,isDraft,headRefName` and parse JSON. No API key management required.  
**Example cell:**
```
ghpr: open, 3, 8
```
Outputs: `#NUM | TITLE (truncated) | STATUS`  
**Feasibility: 4/5** — Requires `gh` CLI to be authenticated (already true for most devs). Fallback to unauthenticated API for public repos.

---

### `quicksheet-gitst`
**Prefix:** `gitst:`  
**One-liner:** Git status for a repo — branch, modified files, untracked count, last commit  
**Wallpaper angle:** The mental overhead of "wait, which branch am I on, are there uncommitted changes?" resolved with a glance.  
**API/Data source:** Local `git` commands — zero network, zero auth. `git status --short`, `git branch --show-current`, `git log --oneline -1`, `git stash list | wc -l`  
**Example cell:**
```
gitst: /home/user/Projects/MyApp, 3, 6
```
Outputs:
```
branch: feature/auth-refactor
modified: 4 files
untracked: 2
stashes: 1
last: abc1234 fix: remove debug log
```
**Feasibility: 5/5** — Pure local git. Works on any repo, any platform.

---

### `quicksheet-hntop`
**Prefix:** `hntop:`  
**One-liner:** Top Hacker News stories with score and comment count  
**Wallpaper angle:** "What's blowing up on HN today?" without opening a browser. For devs who check HN multiple times per day, this is a legitimate tab replacement.  
**API/Data source:** HN Firebase API — `https://hacker-news.firebaseio.com/v0/topstories.json` then `https://hacker-news.firebaseio.com/v0/item/{id}.json`. **Completely free, no API key, no rate limits documented.**  
**Example cell:**
```
hntop: 10, 3, 12
```
Outputs 10 rows: `SCORE | COMMENTS | TITLE`  
**Feasibility: 5/5** — The Firebase Realtime API is public and famously permissive. Cache for 5 minutes to be polite.  

---

### `quicksheet-apistatus`
**Prefix:** `apistatus:`  
**One-liner:** Service status aggregator — GitHub, Cloudflare, Vercel, npm, PagerDuty in one grid  
**Wallpaper angle:** When something is broken, the first question is "is it me or them?" This answers it without opening 5 browser tabs.  
**API/Data source:** All statuspage.io-compatible endpoints return `GET /api/v2/status.json`:
  - GitHub: `https://www.githubstatus.com/api/v2/status.json`
  - Cloudflare: `https://www.cloudflarestatus.com/api/v2/status.json`
  - npm: `https://status.npmjs.org/api/v2/status.json`
  - Vercel: `https://www.vercel-status.com/api/v2/status.json`
  - PagerDuty: `https://status.pagerduty.com/api/v2/status.json`
  - Fastly: `https://status.fastly.com/api/v2/status.json`
  - All return `{"status": {"indicator": "none|minor|major|critical", "description": "..."}}` — **free, no key.**  
**Example cell:**
```
apistatus: github,cloudflare,npm,vercel, 3, 6
```
Outputs: `SERVICE | STATUS | DESCRIPTION`  
**Feasibility: 5/5** — Simple HTTP GET, stable API, all services keep this working.

---

### `quicksheet-ghrun`
**Prefix:** `ghrun:`  
**One-liner:** Latest GitHub Actions workflow run status for a repo  
**Wallpaper angle:** "Did my CI pass?" is checked after every push. Instead of navigating to GitHub Actions tab, it's on the wallpaper. Also invaluable for monitoring nightly build health.  
**API/Data source:** GitHub API `GET /repos/{owner}/{repo}/actions/runs?per_page=5` — or delegate to `gh run list --json status,conclusion,name,headBranch,updatedAt`. Unauthenticated: public repos only; `gh` CLI handles private repos.  
**Example cell:**
```
ghrun: cemheren/QuickSheet, 3, 6
```
Outputs: `WORKFLOW | BRANCH | STATUS | TIME`  
**Feasibility: 4/5** — Like `ghpr:`, cleanest with `gh` CLI delegating auth. Can support public repos unauthenticated as fallback.

---

### `quicksheet-portck`
**Prefix:** `portck:`  
**One-liner:** TCP port/service health checker — is your local postgres/redis/app server up?  
**Wallpaper angle:** Dev environment management. "Which of my services are running right now?" before a focus session.  
**API/Data source:** TCP socket connect attempt — completely local, no network, no credentials. Works on both Windows and Linux with `System.Net.Sockets.TcpClient`.  
**Example cell:**
```
portck: postgres:5432,redis:6379,api:3000,nginx:80, 3, 6
```
Outputs: `SERVICE | PORT | STATUS | LATENCY`
`postgres | 5432 | ● UP | 0.3ms`  
**Feasibility: 5/5** — TCP connect is a trivial stdlib operation, works on both platforms.

---

### `quicksheet-worldtm`
**Prefix:** `worldtm:`  
**One-liner:** Multi-timezone world clock — your team's timezones in one glance  
**Wallpaper angle:** For distributed teams, "what time is it in London / SF / Tokyo?" is a constant friction. This eliminates the "is it okay to ping them right now?" mental math.  
**API/Data source:** Pure local — `System.TimeZoneInfo` (cross-platform .NET 9 stdlib). Zero network calls.  
**Example cell:**
```
worldtm: UTC,America/Los_Angeles,Europe/London,Asia/Tokyo, 2, 6
```
Outputs: `CITY | TIME | OFFSET`
`Tokyo | 09:42 Thu | +09:00`  
**Feasibility: 5/5** — Zero dependencies, pure stdlib, works offline.

---

### 🟡 Tier 2: "Strong Wallpaper Value, Slightly More Setup"

---

### `quicksheet-gitlog`
**Prefix:** `gitlog:`  
**One-liner:** Recent git commits for a repo — who landed what and when  
**Wallpaper angle:** Passive PR monitoring / team awareness. Platform engineers and tech leads want to see what just landed in main without opening GitHub.  
**API/Data source:** `git log --oneline --format="%h %an %s" -10` — local, zero network  
**Example cell:** `gitlog: /path/to/repo, 3, 12`  
**Feasibility: 5/5**

---

### `quicksheet-proc`
**Prefix:** `proc:`  
**One-liner:** Watch named processes — is nginx/postgres/redis actually running?  
**Wallpaper angle:** Local dev services that crash silently. You start working, wonder why the app is broken, and the answer is "postgres died 30 minutes ago". The wallpaper would have told you immediately.  
**API/Data source:** `/proc/[pid]/status` on Linux (parse `cmdline`), `tasklist` on Windows — local only  
**Feasibility: 5/5** — Standard OS APIs, cross-platform

---

### `quicksheet-cve`
**Prefix:** `cve:`  
**One-liner:** Recent CVEs for a package/version from the NIST NVD  
**Wallpaper angle:** Security-conscious devs running production services want to know if their OpenSSL or Nginx version has a new CVE without subscribing to mailing lists.  
**API/Data source:** NIST NVD API v2 — `https://services.nvd.nist.gov/rest/json/cves/2.0?keywordSearch=openssl+3.4` — **free, no key required.** Rate limit: 5 req/30s (use 5-minute cache TTL).  
**Example cell:** `cve: openssl 3.4, 3, 5`  
**Feasibility: 4/5** — Free API but need careful caching. Output: `CVE-ID | SEVERITY | SUMMARY`

---

### `quicksheet-dnsck`
**Prefix:** `dnsck:`  
**One-liner:** Full DNS record inspector — A, AAAA, CNAME, TXT, NS, SOA for any domain  
**Wallpaper angle:** Extends the existing `mxck:` extension (MX only). DevOps engineers managing DNS migrations check record propagation repeatedly. "Has the TTL expired? What's the A record pointing to now?"  
**API/Data source:** Cloudflare DNS-over-HTTPS `https://cloudflare-dns.com/dns-query` — **free, no key, no rate limit documented.** Same mechanism as existing `quicksheet-mxck-ext`.  
**Example cell:** `dnsck: github.com,A, 3, 4`  
**Feasibility: 5/5** — Identical pattern to existing `mxck:` extension which is already shipping.

---

### `quicksheet-jwtdec`
**Prefix:** `jwtdec:`  
**One-liner:** JWT token decoder — paste a token, see the claims instantly  
**Wallpaper angle:** Auth debugging is a constant pain. Devs currently go to `jwt.io` in a browser. This runs locally with zero network, perfect for tokens that may be sensitive.  
**API/Data source:** Pure local base64 decode — no network calls at all. Header + payload are standard base64url.  
**Example cell:** `jwtdec: eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9..., 3, 8`  
**Privacy angle:** This is specifically **better than jwt.io** because tokens never leave the machine.  
**Feasibility: 5/5** — Base64 decode is in every stdlib.

---

### `quicksheet-cronck`
**Prefix:** `cronck:`  
**One-liner:** Cron expression → next N scheduled run times  
**Wallpaper angle:** "Did I write this cron expression right? When will it actually fire?" — common pain for backend devs managing scheduled jobs. No more copy-pasting into crontab.guru.  
**API/Data source:** Pure local cron parsing — no network calls  
**Example cell:** `cronck: 0 */4 * * *, 2, 6` → shows next 5 run times  
**Feasibility: 4/5** — Cron parsing is non-trivial but well-understood. Libraries exist in Go/Python, or implement from scratch for zero-dep compliance.

---

### `quicksheet-deps`
**Prefix:** `deps:`  
**One-liner:** Outdated dependency checker for npm/pip/cargo/dotnet projects  
**Wallpaper angle:** "How behind is this project's dependency tree?" as passive ambient info. Particularly good for devs juggling multiple repos.  
**API/Data source:** CLI tools: `npm outdated --json`, `pip list --outdated --format=json`, `cargo outdated --format=json` — local, no network beyond package registry  
**Example cell:** `deps: /path/to/package.json, 3, 8`  
**Feasibility: 3/5** — Depends on ecosystem-specific CLI tools being installed. Start with `npm` only.

---

### 🔵 Tier 3: "Niche but Loved by Their Specific Audience"

---

### `quicksheet-prom`
**Prefix:** `prom:`  
**One-liner:** Active Prometheus/Alertmanager alerts for a local or remote instance  
**Wallpaper angle:** SREs with self-hosted Prometheus currently have alerts in a browser tab. This puts active alerts on the wallpaper — the always-visible "is the house on fire?" panel.  
**API/Data source:** Alertmanager API `GET http://localhost:9093/api/v2/alerts` — JSON, no auth for local instances  
**Example cell:** `prom: localhost:9093, 3, 8`  
**Feasibility: 4/5** — Only relevant to Prometheus users but they're very numerous in the SRE audience.

---

### `quicksheet-ipcalc`
**Prefix:** `ipcalc:`  
**One-liner:** IP/CIDR subnet calculator — network address, broadcast, usable hosts  
**Wallpaper angle:** Network engineers and DevOps doing VPC/subnet planning constantly calculate these. Better than remembering to open `ipcalc.io`.  
**API/Data source:** Pure local computation — IP math only  
**Example cell:** `ipcalc: 192.168.1.0/24, 2, 6`  
**Feasibility: 5/5** — Pure arithmetic, zero dependencies.

---

### `quicksheet-ctdown`
**Prefix:** `ctdown:`  
**One-liner:** Countdown to a date — sprint end, product launch, conference deadline  
**Wallpaper angle:** "How many days until the launch freeze?" permanently visible on the wallpaper creates healthy urgency without constant reminders.  
**API/Data source:** Pure local — `DateTime` arithmetic  
**Example cell:** `ctdown: 2025-12-31, Sprint Q4 end, 2, 2`  
**Feasibility: 5/5** — Trivially simple. Existing `qtr:` extension is a precedent.

---

### `quicksheet-ghstars`
**Prefix:** `ghstars:`  
**One-liner:** GitHub repo star count with recent daily delta  
**Wallpaper angle:** Indie hackers and OSS maintainers check their star count obsessively. This makes it ambient and adds the trend sparkline that github.com doesn't show on the main page.  
**API/Data source:** GitHub API `GET /repos/{owner}/{repo}` — `stargazers_count` — 60 req/hr unauthenticated. Confirmed: k9s returns 33,626 stars in a simple API call.  
**Example cell:** `ghstars: cemheren/QuickSheet, 2, 3`  
**Feasibility: 5/5** — Single REST call, stable endpoint.

---

## 4. DevOps Pain Points — Validated Research

From studying the trending DevOps repos, `devops-exercises` (38k★), `awesome-sysadmin`, `awesome-sre` (dastergon/awesome-sre), and the k9s/kubeshark/uptime-kuma ecosystems:

| Pain Point | How Devs Currently Cope | QuickSheet Extension |
|-----------|------------------------|---------------------|
| **K8s pod restarts** | `watch kubectl get pods` in tmux pane | `k8s:` — ambient pod table |
| **Local services crashing silently** | Nothing — discover when app breaks | `portck:` / `proc:` — always visible |
| **CI/CD broken on main** | Email notification or browser tab | `ghrun:` — latest run status on wallpaper |
| **PRs piling up** | GitHub notification bell | `ghpr:` — PR count on wallpaper |
| **"Is it GitHub or me?"** | Visit 5 status pages | `apistatus:` — one cell, all services |
| **DNS not propagated yet** | `dig @8.8.8.8 domain` repeatedly | `dnsck:` with auto-refresh loop |
| **New CVE for my stack** | Security mailing list / Dependabot PR | `cve:` ambient alert |
| **Cron expression wrong** | crontab.guru in browser | `cronck:` local |
| **JWT claim debugging** | jwt.io in browser (sends token out) | `jwtdec:` private, local |
| **Team timezone math** | World time website/app | `worldtm:` always visible |

---

## 5. Developer Daily Workflow — The "Check 20x/Day" Map

Verified from studying `gh-dash`, `lazygit`, and `wtf` module popularity:

```
Morning startup sequence (most devs):
1. What's on HN today?         → hntop:  (replaces browser tab)
2. What branch am I on?        → gitst:  (replaces terminal)
3. Do I have PRs to review?    → ghpr:   (replaces github.com)
4. Did nightly CI pass?        → ghrun:  (replaces Actions tab)
5. Are my services up?         → portck: (replaces terminal)
6. What's the weather?         → wthr:   (already shipped)
7. What time is it in London?  → worldtm:(replaces world clock app)

During-day checks:
- My PR got a review           → ghpr:   notification in cell
- Is the deploy pipeline done? → ghrun:  status changed
- Is prod having an incident?  → apistatus: went red
- What just landed in main?    → gitlog: recent commits
```

---

## 6. Priority Build Order (Recommended)

Based on wallpaper value, feasibility, and "replaces a tab" strength:

| Priority | Extension | Reason |
|----------|-----------|--------|
| **P0** | `quicksheet-docker` | 5/5 feasibility, solves universal dev pain, no auth needed |
| **P0** | `quicksheet-hntop` | 5/5, no API key, massive HN audience overlap with QuickSheet target |
| **P0** | `quicksheet-worldtm` | 5/5, pure local, distributed teams love it, zero friction |
| **P0** | `quicksheet-gitst` | 5/5, pure local, used every dev session |
| **P0** | `quicksheet-portck` | 5/5, pure local, solves silent-crash dev pain |
| **P1** | `quicksheet-apistatus` | 5/5, no key, directly answers "is GitHub down?" |
| **P1** | `quicksheet-k8s` | 5/5, SRE audience, kubectl is local |
| **P1** | `quicksheet-ghpr` | 4/5, needs gh CLI, high daily use |
| **P1** | `quicksheet-ghrun` | 4/5, needs gh CLI, daily dev pain |
| **P1** | `quicksheet-jwtdec` | 5/5, local-only, better than jwt.io |
| **P2** | `quicksheet-gitlog` | 5/5 feasibility, more niche than gitst |
| **P2** | `quicksheet-dnsck` | Same pattern as existing mxck:, familiar to ops devs |
| **P2** | `quicksheet-cve` | 4/5, niche but security-focused devs love it |
| **P2** | `quicksheet-cronck` | 4/5, backend devs only, but no alternatives |
| **P3** | `quicksheet-prom` | 4/5, powerful for SRE audience, Prometheus is ubiquitous |
| **P3** | `quicksheet-ghstars` | 5/5, OSS maintainer niche, viral screenshot potential |
| **P3** | `quicksheet-deps` | 3/5, needs ecosystem CLI, start npm-only |
| **P3** | `quicksheet-ipcalc` | 5/5 feasibility, network ops niche |
| **P3** | `quicksheet-ctdown` | 5/5 trivial, complements existing qtr: |

---

## 7. Architecture Notes (QuickSheet-Specific)

From reading `ExtensionProtocol.cs`, `ExtensionManager.cs`, and `ExtensionInstaller.cs`:

- **The protocol is 2 message types:** `activate` (host→ext with `Params[]` and `GridCols`/`GridRows`) and `write` (ext→host with relative `r`/`c` offset cells). Every extension above maps cleanly to this.
- **Cell reference expansion already works:** `{A1::C10}` in params gets expanded before the `activate` message is sent (`ExtensionManager.cs:306`). This means `k8s: {A2::A2}, 4, 10` can pull the namespace from another cell — very powerful for multi-cluster dashboards.
- **Refresh pattern:** Use `L: <cell>, Nm` for the extension call cell. `docker:` probably wants `L: B2, 30s` (30 second loops). `hntop:` is fine at `L: A1, 10m`.
- **All extensions can be written in Python stdlib or Go** (no external packages needed given the simple HTTP + subprocess pattern).
- **The `quicksheet-extension.json` manifest `entry` field** can be `python3 main.py` or `./qs-ext-binary` — any executable works.

---

## 8. Competitive Analysis: WTF Terminal

`wtfutil/wtf` (16k★) is the closest competitor. Key differences:
- WTF lives **inside a terminal session** (requires a terminal window to be open and in view)
- WTF uses YAML config files
- WTF has ~70 pre-built modules (Jenkins, Jira, OpsGenie, etc.)
- WTF is **not editable in-place** — you edit YAML, restart

**QuickSheet's advantages over WTF:**
1. **Wallpaper** — visible when WTF's window is buried
2. **Editable in-place** — click a cell, type, it works
3. **Composable** — `{A1::C10}` refs, sparklines, multi-cell launch sequences
4. **Extensible by the community** in any language via git clone
5. **CSV persistence** — open the data in Excel, diff it in git, pipe it to jq

WTF's advantages over QuickSheet today: more built-in modules. The extension list above closes that gap, focusing on the modules that benefit most from **ambient wallpaper visibility** rather than WTF's "I'm looking at my terminal" model.

---

## Key Gaps & Uncertainties

- **`gh` CLI availability:** `ghpr:` and `ghrun:` work best by delegating to `gh auth token` + `gh api`. If `gh` isn't installed, fallback to unauthenticated GitHub API (public repos only). Both extensions should document this clearly.
- **Docker socket on Windows:** Named pipe path `npipe:////./pipe/docker_engine` works for Docker Desktop but may need elevated permissions on some setups. Extension should surface a clear error cell value.
- **`k8s:` on multi-cluster:** Extension should read `KUBECONFIG` env var and detect active context. `kubectl config current-context` in the extension startup phase.
- **Didn't find a free, key-less GitHub Actions API for private repos** — `gh run list` is the cleanest path but requires `gh` CLI to be installed. Unauthenticated API only covers public repos.
- **`deps:` complexity:** npm, pip, cargo, and dotnet all have different `outdated` output formats. Recommend shipping npm-only first (most common), with a `--ecosystem` param pattern.
- **Awesome-tui list (rothgar/awesome-tui)** returned 404 — the repo may have moved or been deleted. The github.com/topics/tui search was used as the primary source instead.