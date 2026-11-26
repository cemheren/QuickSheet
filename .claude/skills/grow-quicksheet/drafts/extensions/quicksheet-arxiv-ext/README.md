# quicksheet-arxiv-ext

> **2026-05-30: DUPLICATE.** `Deskworks/quicksheet-arxiv` already covers the `arxiv:` prefix
> and is listed in the main README. This repo (`cemheren/quicksheet-arxiv-ext`) was pushed
> before discovering the existing one. User should delete `cemheren/quicksheet-arxiv-ext`
> or repurpose it. Do NOT cross-link from main README — prefix collision.

A QuickSheet extension that looks up arXiv papers by ID. Type an arXiv ID in a cell and get the title, authors, abstract snippet, and link — right on your desktop.

**Persona:** Researchers, grad students, and academics who keep a reading list on their desktop and want quick paper metadata without opening a browser.

## Install

```
ext: github:cemheren/quicksheet-arxiv-ext
```

## Usage

```
arxiv: 2301.07041
arxiv: 2301.07041v2
arxiv: hep-th/9802150
arxiv: https://arxiv.org/abs/2301.07041
```

## Output

The extension fills cells vertically:

| Cell | Content |
|------|---------|
| Row 0 | Paper title (truncated to ~100 chars) |
| Row 1 | Authors (up to 4, then "et al.") |
| Row 2 | Abstract snippet (~120 chars) |
| Row 3 | `https://arxiv.org/abs/<id>` (clickable in QuickSheet) |

## Pairs with

- **`cite:`** — use `cite:` for DOI-based citation formatting, `arxiv:` for preprint lookup.
- Together they cover the full academic workflow: discover on arXiv → cite the published version.

## API

Uses the free [arXiv API](https://info.arxiv.org/help/api/index.html) (Atom feed, no auth, no rate-limit key needed). Requests are cached in-memory for the session.

## Build

```bash
dotnet build ArxivExtension.csproj
```

.NET 9, MIT, zero NuGet dependencies.

## Protocol

Standard QuickSheet JSON-lines:

1. On `{"type":"init"}`, emits `{"type":"register","prefix":"arxiv","name":"arXiv Lookup","version":"1.0.0"}`.
2. On `{"type":"activate","id":"...","params":["2301.07041"]}`, fetches from arXiv API and replies with `{"type":"write","id":"...","cells":[...]}`.

See [docs/extension-protocol.md](https://github.com/cemheren/QuickSheet/blob/main/docs/extension-protocol.md).

## License

MIT.
