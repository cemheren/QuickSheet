# quicksheet-pubmed-ext

PubMed article lookup for [QuickSheet](https://github.com/cemheren/QuickSheet) — search papers, fetch metadata by PMID, see title/authors/journal/DOI right on your desktop wallpaper.

## Install

```
ext: github:cemheren/quicksheet-pubmed-ext
```

## Usage

| Cell value                          | What it does                              |
|-------------------------------------|-------------------------------------------|
| `pubmed: 33782455`                  | Fetch article by PMID                     |
| `pubmed: search CRISPR therapy`     | Search PubMed, show top results           |
| `pubmed: search kidney transplant`  | Any PubMed search term                    |
| `pubmed: help`                      | Show usage reference                      |

## Example output

```
pubmed: 33782455

→  CRISPR-Cas9 gene editing for sickle cell disease and β-thalassemia.
   Frangoul H, Altshuler D, Cappellini MD, et al.
   The New England Journal of Medicine (2021 Jan 21)
   doi:10.1056/NEJMoa2031054
   https://pubmed.ncbi.nlm.nih.gov/33782455/
```

```
pubmed: search mRNA vaccine delivery

→  PubMed: "mRNA vaccine delivery" (2,847 results)
   PMID:35812345 | Pardi N et al. (2023)
     mRNA vaccines — a new era in vaccinology
   PMID:34567890 | Hou X et al. (2022)
     Lipid nanoparticles for mRNA delivery
   ...
```

## Why this on a wallpaper?

- Literature monitoring without browser tabs — glance at your desktop to see latest results for a saved search.
- Pair with `cite:` (DOI → formatted citation) and `arxiv:` (preprint lookup) for a full research dashboard.
- No account needed — uses NCBI E-utilities public API (free, no API key required for low-volume use).

## Honesty

- **Rate limited.** NCBI allows ~3 requests/second without an API key. The extension caches results in-memory so repeated activations don't re-fetch.
- **No full text.** Only metadata (title, authors, journal, date, DOI). Full-text access still requires institutional subscriptions where applicable.
- **No alerts.** This is a point-in-time lookup, not a monitoring service. Re-activate the cell to refresh.

## Build

```bash
dotnet build PubMedExtension.csproj
```

.NET 9, MIT, zero NuGet dependencies.

## Protocol

Standard QuickSheet JSON-lines over stdin/stdout:

1. On `{"type":"init","version":1}`, reply with `{"type":"register","prefix":"pubmed","name":"PubMed Lookup","version":"1.0.0"}`.
2. On `{"type":"activate","id":"...","params":["<query>"],"gridRows":N}`, fetch from NCBI, reply with `{"type":"write","id":"...","cells":[...]}`.

See [docs/extension-protocol.md](https://github.com/cemheren/QuickSheet/blob/main/docs/extension-protocol.md).

## Data source

[NCBI E-utilities](https://www.ncbi.nlm.nih.gov/books/NBK25501/) — the official programmatic API for PubMed. Free for all users; no registration required for low-volume access.

## License

MIT.
