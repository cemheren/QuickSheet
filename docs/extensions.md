# Extensions directory

A live index of QuickSheet extensions. Each one is an independent repo that registers a new cell prefix when you install it with `ext: github:user/repo`.

| Prefix    | Name              | What it does                                       | Repo |
|-----------|-------------------|----------------------------------------------------|------|
| `copilot` | Copilot           | AI in a cell. Q&A, range summarization, generation | [`quicksheet-copilot-ext`](https://github.com/Deskworks/quicksheet-copilot-ext) |
| `wthr`    | Weather forecast  | 7-day forecast for a location                      | [`quicksheet-weather`](https://github.com/Deskworks/quicksheet-weather) |
| `tls`     | TLS cert checker  | Cert expiry, issuer, CN for a host                 | [`quicksheet-tls-ext`](https://github.com/Deskworks/quicksheet-tls-ext) |
| `price`   | Crypto price      | Last trade + 24h change (CoinGecko)                | [`quicksheet-price-ext`](https://github.com/Deskworks/quicksheet-price-ext) |
| `def`     | Dictionary        | Inline definitions (dictionaryapi.dev)             | [`quicksheet-define-ext`](https://github.com/Deskworks/quicksheet-define-ext) |
| `mort`    | Mortgage calc     | Monthly payment, total interest, total cost        | [`quicksheet-mortgage-ext`](https://github.com/Deskworks/quicksheet-mortgage-ext) |
| `pomo`    | Pomodoro timer    | Live countdown with progress bar                   | [`quicksheet-pomodoro`](https://github.com/Deskworks/quicksheet-pomodoro) |
| `sys`     | System monitor    | Live CPU, RAM, disk usage with visual bars         | [`quicksheet-sysmon`](https://github.com/Deskworks/quicksheet-sysmon) |
| `mxck`    | MX record check   | MX records for a domain (DNS-over-HTTPS)           | [`quicksheet-mxck-ext`](https://github.com/Deskworks/quicksheet-mxck-ext) |
| `ping`    | HTTP ping         | Status code + latency for a URL                    | [`quicksheet-ping-ext`](https://github.com/Deskworks/quicksheet-ping-ext) |
| `cite`    | DOI citation      | Authors / year / title / venue from Crossref       | [`quicksheet-cite-ext`](https://github.com/Deskworks/quicksheet-cite-ext) |
| `thes`    | Thesaurus         | Synonyms for a word (Datamuse)                     | [`quicksheet-thes-ext`](https://github.com/Deskworks/quicksheet-thes-ext) |
| `stock`   | Stock quote       | Last close + intra-day change (Stooq)              | [`quicksheet-stock-ext`](https://github.com/Deskworks/quicksheet-stock-ext) |
| `1099`    | SE tax estimate   | US self-employment tax + quarterly estimate        | [`quicksheet-1099-ext`](https://github.com/Deskworks/quicksheet-1099-ext) |
| `qtr`     | Tax countdown     | Next IRS estimated tax deadline + days remaining   | [`quicksheet-qtr`](https://github.com/Deskworks/quicksheet-qtr) |
| `fx`      | Currency convert  | Live ECB rates, 200+ currencies, no API key        | [`quicksheet-fx`](https://github.com/Deskworks/quicksheet-fx) |
| `grav`    | Gravatar lookup   | Profile name, location, avatar URL for an email    | [`quicksheet-grav-ext`](https://github.com/Deskworks/quicksheet-grav-ext) |
| `todo`    | Todo manager      | Tasks with priorities, due dates, completion stats  | [`quicksheet-todo`](https://github.com/Deskworks/quicksheet-todo) |
| `cal`     | Calendar          | Upcoming events from .ics files, grouped by date    | [`quicksheet-cal`](https://github.com/Deskworks/quicksheet-cal) |
| `budget`  | Budget envelopes  | Track spending with visual progress bars            | [`quicksheet-budget`](https://github.com/Deskworks/quicksheet-budget) |
| `hntop`   | HN Top Stories    | Top Hacker News stories with scores & comments      | [`quicksheet-hntop`](https://github.com/Deskworks/quicksheet-hntop) |
| `apistatus` | Service Status  | Monitor GitHub/Cloudflare/npm/Discord status pages  | [`quicksheet-apistatus`](https://github.com/Deskworks/quicksheet-apistatus) |
| `portck`  | Port Checker      | TCP port/service health — which local services are up | [`quicksheet-portck`](https://github.com/Deskworks/quicksheet-portck) |
| `k8s`     | Kubernetes Pods   | Live pod status from kubeconfig — ambient CrashLoopBackOff alerts | [`quicksheet-k8s`](https://github.com/Deskworks/quicksheet-k8s) |
| `ghpr`    | GitHub PRs        | PR review dashboard — see review requests on wallpaper | [`quicksheet-ghpr`](https://github.com/Deskworks/quicksheet-ghpr) |
| `gha`     | GitHub Actions    | Live CI/CD workflow run statuses — ✅❌🔄 on wallpaper, optional token for private repos | [`quicksheet-gha`](https://github.com/Deskworks/quicksheet-gha) |
| `docker`  | Docker Health     | Container status dashboard via Docker Engine API       | [`quicksheet-docker`](https://github.com/Deskworks/quicksheet-docker) |
| `rate`    | Freelance Rate    | Min viable hourly rate for target income (taxes+benefits) | [`quicksheet-rate`](https://github.com/Deskworks/quicksheet-rate) |
| `gitst`   | Git Status        | Branch, changes, stashes, last commit for your repos   | [`quicksheet-gitst`](https://github.com/Deskworks/quicksheet-gitst) |
| `cntdn`   | Countdown Timer   | Days/hours to deadlines, launches, holidays with progress bars | [`quicksheet-cntdn`](https://github.com/Deskworks/quicksheet-cntdn) |
| `worldtm` | World Time        | Multi-timezone clock with business-hours indicators, 40+ aliases | [`quicksheet-worldtm`](https://github.com/Deskworks/quicksheet-worldtm) |
| `margin`  | Margin Calculator | Break-even point & contribution margin from price/cost/fixed     | [`quicksheet-margin-ext`](https://github.com/Deskworks/quicksheet-margin-ext) |
| `mileage` | IRS mileage       | Standard-mileage deduction (business/medical/charity, 2021-2025) | [`quicksheet-mileage-ext`](https://github.com/Deskworks/quicksheet-mileage-ext) |
| `depr`    | Depreciation      | Straight-line + MACRS half-year schedules (IRS Pub 946)          | [`quicksheet-depr-ext`](https://github.com/Deskworks/quicksheet-depr-ext) |
| `jwtdec`  | JWT decoder       | Decode JWT tokens locally — header, claims, expiry. Tokens never leave your machine. | [`quicksheet-jwtdec`](https://github.com/Deskworks/quicksheet-jwtdec) |
| `cronck`  | Cron parser       | Convert 5-field cron expressions to human-readable descriptions. Ranges, steps, named days/months. | [`quicksheet-cronck`](https://github.com/Deskworks/quicksheet-cronck) |
| `news`    | RSS feed reader   | Headlines from HN, Reddit, dev.to, BBC, TechCrunch, or any RSS/Atom URL | [`quicksheet-news`](https://github.com/Deskworks/quicksheet-news) |
| `b64`     | Base64 codec      | Encode/decode base64 with auto-detect — paste tokens, config blobs, JWTs | [`quicksheet-b64`](https://github.com/Deskworks/quicksheet-b64) |
| `guid`    | GUID generator    | Generate UUIDs on demand — standard, no-dash, braced, uppercase, batch up to 20 | [`quicksheet-guid`](https://github.com/Deskworks/quicksheet-guid) |
| `regex`   | Regex explainer   | Tokenize and explain regex patterns — anchors, classes, quantifiers, groups | [`quicksheet-regex`](https://github.com/Deskworks/quicksheet-regex) |
| `urlenc`  | URL codec         | URL encode/decode with auto-detect, component, path, and full URI modes | [`quicksheet-urlenc`](https://github.com/Deskworks/quicksheet-urlenc) |
| `curl`    | HTTP client       | cURL-style HTTP client — GET/POST/PUT/DELETE from cells with headers, JSON bodies, pretty-print | [`quicksheet-curl`](https://github.com/Deskworks/quicksheet-curl) |
| `arxiv`   | arXiv paper lookup | Look up papers by ID or keyword search — title, authors, year, abstract. No API key. | [`quicksheet-arxiv`](https://github.com/Deskworks/quicksheet-arxiv) |
| `pihole`  | Pi-hole stats     | DNS blocking stats from Pi-hole — status, block %, query counts, domains blocked | [`quicksheet-pihole`](https://github.com/Deskworks/quicksheet-pihole) |
| `health`  | HTTP health check | Probe multiple HTTP endpoints in parallel — 🟢/🔴 status with latency, self-signed cert support | [`quicksheet-health`](https://github.com/Deskworks/quicksheet-health) |
| `env`     | Env var inspector | Lookup, filter, PATH exploder — auto-masks secrets (API keys, tokens) | [`quicksheet-envck`](https://github.com/Deskworks/quicksheet-envck) |
| `roll`    | Dice roller       | Roll any dice notation — 2d6+3, d20, 4d6kh3, Fudge dice, crit/fumble detection | [`quicksheet-dice`](https://github.com/Deskworks/quicksheet-dice) |
| `lc`      | LeetCode tracker  | Daily challenge, problem lookup by number/slug, user solve stats (Easy/Medium/Hard, rank) | [`quicksheet-leetcode`](https://github.com/Deskworks/quicksheet-leetcode) |
| `ghst`    | GitHub streak     | Current streak 🔥, longest streak, total contributions, 14-day sparkline. No auth needed. | [`quicksheet-ghstreak`](https://github.com/Deskworks/quicksheet-ghstreak) |
| `iss`     | ISS tracker       | Live ISS position (lat/lon/region/altitude/speed) + all people currently in space, grouped by spacecraft | [`quicksheet-iss`](https://github.com/Deskworks/quicksheet-iss) |
| `npm`     | npm package info  | Package version, weekly downloads, license, author, last publish date. Multi-package comparison with comma-separated names. | [`quicksheet-npm`](https://github.com/Deskworks/quicksheet-npm) |
| `pypi`    | PyPI package info | Version, license, author, Python requirement, release date — single package or comparison table | [`quicksheet-pypi`](https://github.com/Deskworks/quicksheet-pypi) |
| `crates`  | Rust crate lookup | Version, all-time downloads, recent downloads, description, homepage/repo link, keywords. Search mode: `crates: search async`. Via crates.io, no API key, 30-min cache. | [`quicksheet-crates`](https://github.com/Deskworks/quicksheet-crates) |
| `ghtrend` | GitHub trending   | Today's trending repos by language — name, stars, forks, description. Type `ghtrend: python` or `ghtrend: all` | [`quicksheet-gh-trends`](https://github.com/Deskworks/quicksheet-gh-trends) |
| `co2`     | CO₂ Monitor       | Live atmospheric CO₂ from NOAA Mauna Loa (free, no key) — current ppm, pre-industrial baseline, year-over-year change, 30-day trend, annual averages | [`quicksheet-co2`](https://github.com/Deskworks/quicksheet-co2) |
| `weather` | Live Weather      | Current weather via Open-Meteo (free, no API key) — temperature, feels-like, conditions, wind, humidity, pressure. Usage: `weather: London` or `weather: 48.85,2.35` | [`quicksheet-openmeteo`](https://github.com/Deskworks/quicksheet-openmeteo) |
| `ip`      | IP Info           | IP geolocation, ISP, org, ASN, timezone, hostname, and coords via ipinfo.io. Free, no API key. Type `ip:` (own IP) or `ip: 8.8.8.8` | [`quicksheet-ipinfo`](https://github.com/Deskworks/quicksheet-ipinfo) |
| `dns`     | DNS Lookup        | DNS record lookup — A, AAAA, MX, TXT, NS, CNAME, PTR. Usage: `dns: github.com` · `dns: MX gmail.com` · `dns: PTR 8.8.8.8` · `dns: ALL example.com`. Uses system `dig`, 5-min cache. | [`quicksheet-dns`](https://github.com/Deskworks/quicksheet-dns) |
| `ssl`     | SSL Cert Checker  | SSL certificate expiry — status 🟢/🟡/🔴, days left, subject, issuer. Usage: `ssl: github.com` or `ssl: example.com:8443`. BCL SslStream, 30-min cache, zero NuGet. | [`quicksheet-ssl`](https://github.com/Deskworks/quicksheet-ssl) |
| `whois`   | WHOIS Lookup      | WHOIS domain and IP lookup — registrar, status, expiry urgency 🟡🔴, nameservers, IP org/CIDR. Usage: `whois: github.com` · `whois: 8.8.8.8`. 25+ TLD servers, 1h cache, zero NuGet. | [`quicksheet-whois`](https://github.com/Deskworks/quicksheet-whois) |
| `tracert` | Network Traceroute | Route trace to any host — each hop with RTT latency and reverse-DNS hostname. Up to 30 hops, 3 probes/hop, 5-min cache. Usage: `tracert: github.com` · `tracert: 8.8.8.8`. BCL Ping, zero NuGet. | [`quicksheet-tracert`](https://github.com/cemheren/quicksheet-tracert) |
| `mtr`     | Route Tracer (mtr-style) | Network route tracer with per-hop RTT stats and packet loss — like `mtr`/`traceroute`. Shows hop#, IP, avg RTT, loss%. `mtr: google.com` (full route trace) · `mtr: ping 8.8.8.8` (ping stats: min/avg/max RTT, loss%). BCL Ping, zero NuGet, cross-platform. | [`quicksheet-mtr`](https://github.com/Deskworks/quicksheet-mtr) |
| `speed`   | Network Speed Test | Measure download speed (Mbps) and HTTP latency (ms) from your wallpaper. `speed:` (all), `speed: download` (CDN throughput), `speed: latency` (Cloudflare/Google/GitHub/Fastly). Color-coded 🟢🟡🔴 indicators. BCL HttpClient, zero NuGet. | [`quicksheet-speedtest`](https://github.com/Deskworks/quicksheet-speedtest) |
| `epoch`   | Epoch Converter   | Unix timestamp converter — no network, pure math. `epoch: now` (current ts), `epoch: 1700000000` (epoch→date+age), `epoch: 2025-01-01` (date→epoch), `epoch: diff t1 t2` (time difference), `epoch: add <ts> 30d` (add duration: y/d/h/m/s). | [`quicksheet-epoch`](https://github.com/Deskworks/quicksheet-epoch) |
| `color`   | Color Converter   | Color code converter and palette generator — zero network, pure math. `color: #ff6600` (hex→RGB/HSL/CMYK/name/WCAG), `color: rgb(255,102,0)`, `color: hsl(24,100%,50%)`, `color: orange` (CSS name), `color: palette #3b82f6` (10-color harmony palette). ANSI swatches, 140+ CSS names. | [`quicksheet-color`](https://github.com/Deskworks/quicksheet-color) |
| `caniuse` | Browser Compat    | Can I Use browser compatibility lookup via caniuse-db. Shows support status for CSS/HTML/JS features across Chrome, Firefox, Safari, Edge, Opera, Samsung, iOS Safari. Usage: `caniuse: css-grid` · `caniuse: flexbox` · `caniuse: webassembly`. Usage %, spec link, notes. 1h cache, no API key. | [`quicksheet-caniuse`](https://github.com/Deskworks/quicksheet-caniuse) |
| `gem`     | Ruby Gem Lookup   | Ruby gem info and search via rubygems.org. Version, total downloads, authors, license, Ruby requirement, homepage, description. `gem: rails` (detail), `gem: search json parser` (top 5). No API key, 30-min cache, zero NuGet. | [`quicksheet-rubygems`](https://github.com/Deskworks/quicksheet-rubygems) |
| `hackage` | Haskell Packages  | Haskell package lookup from Hackage. `hackage: aeson` shows latest version, synopsis, author, category, license, and homepage. `hackage: search json` lists top matching packages. No API key, 30-min cache, zero NuGet. | [`quicksheet-hackage`](https://github.com/Deskworks/quicksheet-hackage) |

## Install

For any extension in the table:

```
ext: github:<owner>/<repo>
```

QuickSheet clones the repo, reads its `quicksheet-extension.json` manifest, and starts a subprocess that talks JSON-lines on stdin/stdout. The prefix in the manifest becomes a live cell prefix.

## Build your own

See **[Extension Protocol Specification](extension-protocol.md)** for the complete, strict protocol reference — message formats, coordinate system, rules, common mistakes, and working examples in C# and Python.

**Quick summary:** The protocol is intentionally tiny:

1. QuickSheet sends `{"type":"init"}` → extension replies with `{"type":"register","prefix":"xyz",...}`.
2. When a cell matching the prefix is activated, QuickSheet sends `{"type":"activate","id":"...","params":[...],"gridRows":N,"gridCols":M}` → extension replies with `{"type":"write","id":"...","cells":[[...]]}`.

Manifest format (`quicksheet-extension.json` at repo root):

```json
{
  "name": "myext",
  "version": "1.0.0",
  "prefix": "mx",
  "description": "what it does",
  "entry": "dotnet run --project MyExt.csproj",
  "minProtocolVersion": 1
}
```

Pick any language with stdin/stdout. The reference extensions are .NET 9 with zero NuGet dependencies, but Python stdlib or Go would work just as well.

## Conventions for new extensions

- Repo name: `quicksheet-<thing>-ext` (or `quicksheet-<thing>` for one-word names).
- License: MIT. Zero external dependencies where possible — `extension repos clone on install`, so a five-line subprocess is much easier to trust than a tree of packages.
- Cache network responses with a TTL appropriate to the data. Network calls are not free for the user.
- Surface failures as cell content (e.g. `err: <message>`) instead of crashing the subprocess.

## Submit yours

Open a PR adding a row to the table above. Or just open an issue on the QuickSheet repo with the link and we'll add it.
