# QuickSheet for SREs and DevOps engineers

If you're on call, managing Kubernetes clusters, chasing TLS expiry dates, or constantly
flipping between terminal tabs to check GitHub Actions, Docker health, and service status
pages — **your desktop wallpaper can hold all of it.**

> Already use k9s, lazygit, gh-dash, or Grafana dashboards on a second monitor? QuickSheet
> is the ambient layer that stays on-screen *behind* every window, always visible, never
> stealing focus.

## Five-minute setup

```bash
git clone https://github.com/cemheren/QuickSheet.git
cd QuickSheet
dotnet run -c Release --project ExcelConsole.csproj -- --desktop examples/sre-dashboard.csv
```

The starter sheet at `examples/sre-dashboard.csv` gives you a deployable frame. Edit cells
live; grid autosaves to CSV every 5 seconds. Add it to your session startup and the
situational-awareness layer is there every boot.

## What goes on the wallpaper

| Row/column             | What it shows                                           | Extension                                                                   |
|------------------------|---------------------------------------------------------|-----------------------------------------------------------------------------|
| **Service health**     | HTTP ✓/✗ + latency for each endpoint in `services.csv` | [`health`](https://github.com/cemheren/quicksheet-health)                   |
| **Vendor status**      | GitHub / Cloudflare / npm / Vercel / Discord live pages | [`apistatus`](https://github.com/cemheren/quicksheet-apistatus)             |
| **Container health**   | Running/stopped/restarting per Docker container name    | [`docker`](https://github.com/cemheren/quicksheet-docker)                   |
| **Pod status**         | Kubernetes pod state by namespace                       | [`k8s`](https://github.com/cemheren/quicksheet-k8s)                        |
| **Port probe**         | TCP open/closed for well-known local ports              | [`portck`](https://github.com/cemheren/quicksheet-portck)                  |
| **TLS expiry**         | Days remaining on each certificate                      | [`tls`](https://github.com/cemheren/quicksheet-tls-ext)                    |
| **Recent commits**     | Last N git commits: hash, author, relative time         | [`gitlog`](https://github.com/cemheren/quicksheet-gitlog)                  |
| **Open PRs**           | PR count + titles per repo                              | [`ghpr`](https://github.com/cemheren/quicksheet-ghpr)                      |
| **System metrics**     | CPU / RAM / disk / uptime of the host                   | [`sysmon`](https://github.com/cemheren/quicksheet-sysmon)                  |
| **Latency**            | Round-trip to upstreams and DNS resolvers               | [`ping`](https://github.com/cemheren/quicksheet-ping-ext)                  |

All extensions install with a single cell: `ext: github:cemheren/quicksheet-<name>`. No
package manager. No daemon config. QuickSheet clones the repo, starts the subprocess, and
the prefix is live.

## Make it useful during an incident

In an incident, you want to see:

1. **Which services are degraded?** — Put a `health:` row across all your critical endpoints.
   The extension probes in parallel; green ✓ or red ✗ refreshes every activation.

2. **Is it us or the vendor?** — One row of `apistatus: github`, `apistatus: cloudflare`,
   `apistatus: aws` etc. tells you immediately whether it's your infra or theirs.

3. **Is it the cert?** — `tls: api.example.com` shows days remaining. Expired certs cause
   deceptively generic errors.

4. **What changed recently?** — `gitlog: /path/to/repo 10` shows the last 10 commits, hashes
   and relative timestamps. Paste a hash anywhere as a runnable cell: `r: git show <hash>`.

5. **Is the pod actually running?** — `k8s: production` lists pods and their status without
   opening another terminal.

## JWT and URL debugging during on-call

Two utilities that come up in every incident involving auth or API routing:

- `jwtdec: <paste-token-here>` — decodes header + claims locally, flags expired tokens,
  annotates `iat`/`exp`/`nbf`. [Privacy-first alternative to jwt.io](https://github.com/cemheren/quicksheet-jwtdec).
- `urlenc: <url>` — URL-encodes or decodes with auto-detect. Useful for debugging redirect
  chains and malformed query strings. [quicksheet-urlenc](https://github.com/cemheren/quicksheet-urlenc)

Both run as local subprocesses — no token leaves your machine.

## Cron job visibility

If you maintain cron jobs, `cronck: 0 3 * * MON-FRI` turns any cron expression into plain
English: "At 03:00 on every day-of-week from Monday through Friday." Add a column of
`cronck:` cells next to each job row to document schedules without opening a wiki.

See [quicksheet-cronck](https://github.com/cemheren/quicksheet-cronck).

## Terminal commands as runnable cells

Prefix any cell with `r: ` to make it a one-keypress shell command:

```
r: kubectl get nodes
r: docker ps --format "table {{.Names}}\t{{.Status}}"
r: systemctl status nginx
r: journalctl -u myservice --since "5 min ago" --no-pager
```

Press **Enter** on the cell to run. Output replaces the cell content. No terminal tab
switch, no history scrolling.

## Why wallpaper > tmux panel or Grafana tab

| Pain point                           | Grafana / Datadog tab          | QuickSheet wallpaper                      |
|--------------------------------------|-------------------------------|-------------------------------------------|
| Visible during code review           | No (alt-tab)                  | **Yes — behind every window**             |
| Setup time                           | Hours (datasources, panels)   | **5 minutes (CSV + ext: cells)**          |
| Custom extension for one-off service | Not possible without plugins  | **50-line script in any language**        |
| Data leaves your machine             | Yes (SaaS)                    | **No — local subprocesses only**          |
| Works offline                        | No                            | **Yes (local probes)**                    |
| Per-team customization               | Dashboard-as-code (complex)   | **Edit CSV directly**                     |

## Make it yours

Cycle themes with **Ctrl+T**: Dark, Light, Nord, Solarized, Matrix, Dracula, Synthwave,
Gruvbox, Monokai. For an SRE aesthetic, Nord and Gruvbox keep it readable at 4 AM.

Sort any column with **Ctrl+B** — useful for sorting services by latency or TLS days
remaining. Undo any change with **Ctrl+Z**.

## Build your own in 50 lines

Don't see an extension for your specific service? The protocol is two message types
([extension-protocol.md](extension-protocol.md)). Standard-library Python or Go is enough.
Any language that can read stdin and write stdout works. The full spec with a worked
example is in [`docs/extension-protocol.md`](extension-protocol.md).

## Share your setup

Post your SRE wallpaper on r/sysadmin, r/devops, or r/unixporn. Open an issue if you
want a new prefix officially listed in [docs/extensions.md](extensions.md).
