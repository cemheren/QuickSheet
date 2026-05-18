# Extensions directory

A live index of QuickSheet extensions. Each one is an independent repo that registers a new cell prefix when you install it with `ext: github:user/repo`.

| Prefix    | Name              | What it does                                       | Repo |
|-----------|-------------------|----------------------------------------------------|------|
| `copilot` | Copilot           | AI in a cell. Q&A, range summarization, generation | [`quicksheet-copilot-ext`](https://github.com/cemheren/quicksheet-copilot-ext) |
| `wthr`    | Weather forecast  | 7-day forecast for a location                      | [`quicksheet-weather`](https://github.com/cemheren/quicksheet-weather) |
| `tls`     | TLS cert checker  | Cert expiry, issuer, CN for a host                 | [`quicksheet-tls-ext`](https://github.com/cemheren/quicksheet-tls-ext) |
| `price`   | Crypto price      | Last trade + 24h change (CoinGecko)                | [`quicksheet-price-ext`](https://github.com/cemheren/quicksheet-price-ext) |
| `def`     | Dictionary        | Inline definitions (dictionaryapi.dev)             | [`quicksheet-define-ext`](https://github.com/cemheren/quicksheet-define-ext) |
| `mort`    | Mortgage calc     | Monthly payment, total interest, total cost        | [`quicksheet-mortgage-ext`](https://github.com/cemheren/quicksheet-mortgage-ext) |
| `pomo`    | Pomodoro timer    | Live countdown with progress bar                   | [`quicksheet-pomodoro`](https://github.com/cemheren/quicksheet-pomodoro) |
| `sys`     | System monitor    | Live CPU, RAM, disk usage with visual bars         | [`quicksheet-sysmon`](https://github.com/cemheren/quicksheet-sysmon) |
| `mxck`    | MX record check   | MX records for a domain (DNS-over-HTTPS)           | [`quicksheet-mxck-ext`](https://github.com/cemheren/quicksheet-mxck-ext) |
| `ping`    | HTTP ping         | Status code + latency for a URL                    | [`quicksheet-ping-ext`](https://github.com/cemheren/quicksheet-ping-ext) |
| `cite`    | DOI citation      | Authors / year / title / venue from Crossref       | [`quicksheet-cite-ext`](https://github.com/cemheren/quicksheet-cite-ext) |
| `thes`    | Thesaurus         | Synonyms for a word (Datamuse)                     | [`quicksheet-thes-ext`](https://github.com/cemheren/quicksheet-thes-ext) |
| `stock`   | Stock quote       | Last close + intra-day change (Stooq)              | [`quicksheet-stock-ext`](https://github.com/cemheren/quicksheet-stock-ext) |
| `1099`    | SE tax estimate   | US self-employment tax + quarterly estimate        | [`quicksheet-1099-ext`](https://github.com/cemheren/quicksheet-1099-ext) |
| `qtr`     | Tax countdown     | Next IRS estimated tax deadline + days remaining   | [`quicksheet-qtr`](https://github.com/cemheren/quicksheet-qtr) |
| `fx`      | Currency convert  | Live ECB rates, 200+ currencies, no API key        | [`quicksheet-fx`](https://github.com/cemheren/quicksheet-fx) |
| `grav`    | Gravatar lookup   | Profile name, location, avatar URL for an email    | [`quicksheet-grav-ext`](https://github.com/cemheren/quicksheet-grav-ext) |
| `todo`    | Todo manager      | Tasks with priorities, due dates, completion stats  | [`quicksheet-todo`](https://github.com/cemheren/quicksheet-todo) |
| `cal`     | Calendar          | Upcoming events from .ics files, grouped by date    | [`quicksheet-cal`](https://github.com/cemheren/quicksheet-cal) |
| `budget`  | Budget envelopes  | Track spending with visual progress bars            | [`quicksheet-budget`](https://github.com/cemheren/quicksheet-budget) |
| `hntop`   | HN Top Stories    | Top Hacker News stories with scores & comments      | [`quicksheet-hntop`](https://github.com/cemheren/quicksheet-hntop) |
| `apistatus` | Service Status  | Monitor GitHub/Cloudflare/npm/Discord status pages  | [`quicksheet-apistatus`](https://github.com/cemheren/quicksheet-apistatus) |
| `portck`  | Port Checker      | TCP port/service health — which local services are up | [`quicksheet-portck`](https://github.com/cemheren/quicksheet-portck) |
| `k8s`     | Kubernetes Pods   | Live pod status from kubeconfig — ambient CrashLoopBackOff alerts | [`quicksheet-k8s`](https://github.com/cemheren/quicksheet-k8s) |
| `ghpr`    | GitHub PRs        | PR review dashboard — see review requests on wallpaper | [`quicksheet-ghpr`](https://github.com/cemheren/quicksheet-ghpr) |
| `gha`     | GitHub Actions    | Live CI/CD workflow run statuses — ✅❌🔄 on wallpaper, optional token for private repos | [`quicksheet-gha`](https://github.com/cemheren/quicksheet-gha) |
| `docker`  | Docker Health     | Container status dashboard via Docker Engine API       | [`quicksheet-docker`](https://github.com/cemheren/quicksheet-docker) |
| `rate`    | Freelance Rate    | Min viable hourly rate for target income (taxes+benefits) | [`quicksheet-rate`](https://github.com/cemheren/quicksheet-rate) |
| `gitst`   | Git Status        | Branch, changes, stashes, last commit for your repos   | [`quicksheet-gitst`](https://github.com/cemheren/quicksheet-gitst) |
| `cntdn`   | Countdown Timer   | Days/hours to deadlines, launches, holidays with progress bars | [`quicksheet-cntdn`](https://github.com/cemheren/quicksheet-cntdn) |
| `worldtm` | World Time        | Multi-timezone clock with business-hours indicators, 40+ aliases | [`quicksheet-worldtm`](https://github.com/cemheren/quicksheet-worldtm) |
| `margin`  | Margin Calculator | Break-even point & contribution margin from price/cost/fixed     | [`quicksheet-margin-ext`](https://github.com/cemheren/quicksheet-margin-ext) |
| `mileage` | IRS mileage       | Standard-mileage deduction (business/medical/charity, 2021-2025) | [`quicksheet-mileage-ext`](https://github.com/cemheren/quicksheet-mileage-ext) |
| `depr`    | Depreciation      | Straight-line + MACRS half-year schedules (IRS Pub 946)          | [`quicksheet-depr-ext`](https://github.com/cemheren/quicksheet-depr-ext) |
| `jwtdec`  | JWT decoder       | Decode JWT tokens locally — header, claims, expiry. Tokens never leave your machine. | [`quicksheet-jwtdec`](https://github.com/cemheren/quicksheet-jwtdec) |
| `cronck`  | Cron parser       | Convert 5-field cron expressions to human-readable descriptions. Ranges, steps, named days/months. | [`quicksheet-cronck`](https://github.com/cemheren/quicksheet-cronck) |
| `news`    | RSS feed reader   | Headlines from HN, Reddit, dev.to, BBC, TechCrunch, or any RSS/Atom URL | [`quicksheet-news`](https://github.com/cemheren/quicksheet-news) |
| `b64`     | Base64 codec      | Encode/decode base64 with auto-detect — paste tokens, config blobs, JWTs | [`quicksheet-b64`](https://github.com/cemheren/quicksheet-b64) |
| `guid`    | GUID generator    | Generate UUIDs on demand — standard, no-dash, braced, uppercase, batch up to 20 | [`quicksheet-guid`](https://github.com/cemheren/quicksheet-guid) |
| `regex`   | Regex explainer   | Tokenize and explain regex patterns — anchors, classes, quantifiers, groups | [`quicksheet-regex`](https://github.com/cemheren/quicksheet-regex) |
| `urlenc`  | URL codec         | URL encode/decode with auto-detect, component, path, and full URI modes | [`quicksheet-urlenc`](https://github.com/cemheren/quicksheet-urlenc) |
| `curl`    | HTTP client       | cURL-style HTTP client — GET/POST/PUT/DELETE from cells with headers, JSON bodies, pretty-print | [`quicksheet-curl`](https://github.com/cemheren/quicksheet-curl) |
| `arxiv`   | arXiv paper lookup | Look up papers by ID or keyword search — title, authors, year, abstract. No API key. | [`quicksheet-arxiv`](https://github.com/cemheren/quicksheet-arxiv) |
| `pihole`  | Pi-hole stats     | DNS blocking stats from Pi-hole — status, block %, query counts, domains blocked | [`quicksheet-pihole`](https://github.com/cemheren/quicksheet-pihole) |
| `health`  | HTTP health check | Probe multiple HTTP endpoints in parallel — 🟢/🔴 status with latency, self-signed cert support | [`quicksheet-health`](https://github.com/cemheren/quicksheet-health) |
| `env`     | Env var inspector | Lookup, filter, PATH exploder — auto-masks secrets (API keys, tokens) | [`quicksheet-envck`](https://github.com/cemheren/quicksheet-envck) |
| `roll`    | Dice roller       | Roll any dice notation — 2d6+3, d20, 4d6kh3, Fudge dice, crit/fumble detection | [`quicksheet-dice`](https://github.com/cemheren/quicksheet-dice) |
| `lc`      | LeetCode tracker  | Daily challenge, problem lookup by number/slug, user solve stats (Easy/Medium/Hard, rank) | [`quicksheet-leetcode`](https://github.com/cemheren/quicksheet-leetcode) |
| `ghst`    | GitHub streak     | Current streak 🔥, longest streak, total contributions, 14-day sparkline. No auth needed. | [`quicksheet-ghstreak`](https://github.com/cemheren/quicksheet-ghstreak) |
| `iss`     | ISS tracker       | Live ISS position (lat/lon/region/altitude/speed) + all people currently in space, grouped by spacecraft | [`quicksheet-iss`](https://github.com/cemheren/quicksheet-iss) |
| `npm`     | npm package info  | Package version, weekly downloads, license, author, last publish date. Multi-package comparison with comma-separated names. | [`quicksheet-npm`](https://github.com/cemheren/quicksheet-npm) |

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
