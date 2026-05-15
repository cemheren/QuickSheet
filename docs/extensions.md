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
| `grav`    | Gravatar lookup   | Profile name, location, avatar URL for an email    | [`quicksheet-grav-ext`](https://github.com/cemheren/quicksheet-grav-ext) |
| `todo`    | Todo manager      | Tasks with priorities, due dates, completion stats  | [`quicksheet-todo`](https://github.com/cemheren/quicksheet-todo) |
| `cal`     | Calendar          | Upcoming events from .ics files, grouped by date    | [`quicksheet-cal`](https://github.com/cemheren/quicksheet-cal) |

## Install

For any extension in the table:

```
ext: github:<owner>/<repo>
```

QuickSheet clones the repo, reads its `quicksheet-extension.json` manifest, and starts a subprocess that talks JSON-lines on stdin/stdout. The prefix in the manifest becomes a live cell prefix.

## Build your own

The protocol is intentionally tiny:

1. QuickSheet sends `{"type":"init"}` → extension replies with `{"type":"register","prefix":"xyz",...}`.
2. When a cell matching the prefix is activated, QuickSheet sends `{"type":"activate","id":"...","params":[...],"gridRows":N,"gridCols":M}` → extension replies with `{"type":"write","id":"...","cells":[[...]]}`.

Manifest format:

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
