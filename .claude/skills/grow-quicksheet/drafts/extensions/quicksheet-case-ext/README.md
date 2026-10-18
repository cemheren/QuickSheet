# quicksheet-case-ext

A QuickSheet extension that looks up US case law citations and keywords via
[CourtListener](https://www.courtlistener.com/)'s free public API.

```
ext: github:cemheren/quicksheet-case-ext
case: Roe v Wade
case: 410 U.S. 113
case: "unreasonable search" fourth amendment
```

## What you get

Up to 5 results per query (configurable via grid rows). Each row shows:

| Case Name                    | Citation       | Court | Year | Snippet                                         |
|------------------------------|----------------|-------|------|--------------------------------------------------|
| Roe v. Wade                  | 410 U.S. 113  | SCOTUS| 1973 | right of privacy...broad enough to encompass...  |
| Planned Parenthood v. Casey  | 505 U.S. 833  | SCOTUS| 1992 | undue burden standard...liberty of the woman...  |

## Features

- **Citation search**: type a standard reporter citation (`410 U.S. 113`, `347 U.S. 483`) for a direct hit.
- **Keyword search**: natural language queries across millions of opinions.
- **No API key required**: CourtListener's search endpoint is free and open (rate-limited to ~100 req/hr).
- **HTML-stripped snippets**: result excerpts are cleaned of markup for readable cell display.

## Use cases

- Law students reviewing cases for class prep — keep citations on your desktop.
- Paralegals quickly verifying citation accuracy without leaving their workspace.
- Anyone who wants an ambient "legal research sticky note" on their wallpaper.

## Build

```bash
dotnet build CaseExtension.csproj
```

.NET 9, MIT, zero NuGet dependencies.

## Protocol

Standard QuickSheet JSON-lines protocol over stdin/stdout:

1. On startup, emit `{"type":"register","prefix":"case","name":"Case Law","version":"1.0.0"}`.
2. On each `{"type":"activate","id":"...","params":["<query>"],"gridRows":N}`, search CourtListener and reply with `{"type":"write","id":"...","cells":[...]}`.

See [docs/extension-protocol.md](https://github.com/cemheren/QuickSheet/blob/main/docs/extension-protocol.md) in the main repo.

## API

Uses CourtListener REST API v4: `https://www.courtlistener.com/api/rest/v4/search/?type=o&q=<query>`

CourtListener is operated by [Free Law Project](https://free.law/), a 501(c)(3) non-profit. The API is free for reasonable use. If you plan heavy usage, consider [registering for a token](https://www.courtlistener.com/sign-in/).

## License

MIT.
