# quicksheet-cite-ext

DOI → citation lookup for [QuickSheet](https://github.com/cemheren/QuickSheet). Uses Crossref's free API (no key).

## Install

```
ext: github:cemheren/quicksheet-cite-ext
```

## Use

```
cite: 10.1145/3623476.3623525, 1, 4
```

Fills 4 rows:

```
Hassan F., Zhao Y., et al. (2024)
Toward Verifying Smart Contracts at Compile Time
Proceedings of the ACM Workshop on PL/Verification
doi:10.1145/3623476.3623525
```

Accepts plain DOIs, `doi:...` form, or full `https://doi.org/...` URLs.

## Why

Researchers, students, and writers end up shuttling DOIs between a paper, a notes file, and a citation manager all day. A `cite:` cell in your QuickSheet wallpaper gives you author, year, title, and venue at a glance — no Zotero round-trip, no browser tab.

## Build

Requires .NET 9. Zero NuGet dependencies — only BCL `System.Net.Http` and `System.Text.Json`.

```
dotnet build CiteExtension.csproj
```

## Notes

- Uses https://api.crossref.org (free, polite-pool). The extension sets a User-Agent identifying itself per Crossref's etiquette.
- Citations cached in-memory; DOIs don't change.
- Author list truncated to 3 + "et al." for readability.
- Output is a quick reference, not a full APA/MLA-formatted citation. Use a proper citation manager for paper-ready output.

## Protocol

Reads JSON lines on stdin, writes JSON lines on stdout. See QuickSheet's main README for the full extension protocol spec.

## License

MIT
