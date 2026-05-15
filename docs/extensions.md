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
| `ghpr`    | GitHub PRs        | PR review dashboard — see review requests on wallpaper | [`quicksheet-ghpr`](https://github.com/cemheren/quicksheet-ghpr) |
| `docker`  | Docker Health     | Container status dashboard via Docker Engine API       | [`quicksheet-docker`](https://github.com/cemheren/quicksheet-docker) |
| `rate`    | Freelance Rate    | Min viable hourly rate for target income (taxes+benefits) | [`quicksheet-rate`](https://github.com/cemheren/quicksheet-rate) |
| `gitst`   | Git Status        | Branch, changes, stashes, last commit for your repos   | [`quicksheet-gitst`](https://github.com/cemheren/quicksheet-gitst) |
| `cntdn`   | Countdown Timer   | Days/hours to deadlines, launches, holidays with progress bars | [`quicksheet-cntdn`](https://github.com/cemheren/quicksheet-cntdn) |
| `mileage` | IRS mileage       | Standard-mileage deduction (business/medical/charity, 2021-2025) | [`quicksheet-mileage-ext`](https://github.com/cemheren/quicksheet-mileage-ext) |

## Install

For any extension in the table:

```
ext: github:<owner>/<repo>
```

QuickSheet clones the repo, reads its `quicksheet-extension.json` manifest, and starts a subprocess that talks JSON-lines on stdin/stdout. The prefix in the manifest becomes a live cell prefix.

## Build your own

The protocol is intentionally tiny:

1. On startup, the extension prints a register message: `{"type":"register","prefix":"xyz","name":"...","version":"1.0.0"}`. **`version` must be a string** — `1` (int) silently fails deserialization on some hosts.
2. When a cell matching the prefix is activated, QuickSheet sends `{"type":"activate","id":"...","params":["arg1","arg2"],"gridCols":N,"gridRows":M}`. **`params` is a JSON array of strings**, not a single `arguments` string — extensions reading `arguments` will see empty input.
3. Extension replies with `{"type":"write","id":"...","cells":[...]}`. Two `cells` shapes are accepted:
   - **Object records** (preferred): `[{"r":0,"c":0,"v":"hello"}, {"r":0,"c":1,"v":"world"}]`. Explicit positions, easy to render sparse grids.
   - **Row-major grid**: `[["a","b"], ["c","d"]]`. Each inner array is a row; positions are inferred relative to the activation cell.

Status / errors are optional: `{"type":"status","id":"...","message":"..."}` or `{"type":"error","id":"...","message":"..."}`.

Manifest format (`quicksheet-extension.json` at the repo root):

```json
{
  "name": "myext",
  "version": "1.0.0",
  "entry": "dotnet run --project MyExt.csproj"
}
```

**Key gotchas:**
- The file is `quicksheet-extension.json`. The launch command is `entry`, not `entrypoint`.
- The cell `prefix` (e.g. `mx`) is declared in the **register message** at runtime, not in the manifest.
- `entry` is shelled with `bash -c` (or `cmd /c` on Windows), so pipes and redirects work but the program must read stdin and write stdout in line-delimited JSON.

Pick any language with stdin/stdout. The reference extensions are .NET 9 with zero NuGet dependencies, but Python stdlib or Go would work just as well.

## Conventions for new extensions

- Repo name: `quicksheet-<thing>-ext` (or `quicksheet-<thing>` for one-word names).
- License: MIT. Zero external dependencies where possible — `extension repos clone on install`, so a five-line subprocess is much easier to trust than a tree of packages.
- Cache network responses with a TTL appropriate to the data. Network calls are not free for the user.
- Surface failures as cell content (e.g. `err: <message>`) instead of crashing the subprocess.

## Submit yours

Open a PR adding a row to the table above. Or just open an issue on the QuickSheet repo with the link and we'll add it.
