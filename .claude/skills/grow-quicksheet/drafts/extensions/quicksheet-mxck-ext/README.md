# quicksheet-mxck-ext

Inline MX record lookup for [QuickSheet](https://github.com/cemheren/QuickSheet). Uses Google's public DNS-over-HTTPS resolver — no key required.

## Install

Type into any cell:

```
ext: github:cemheren/quicksheet-mxck-ext
```

## Use

```
mxck: example.com, 1, 5
```

Fills 5 rows:

```
MX for example.com
 10 mail.example.com
 20 alt1.aspmx.example.com
 30 alt2.aspmx.example.com
```

The first column is the MX priority (lower = preferred).

## Why

If you administer email — your own domain, a customer's, or you're debugging deliverability — MX records are a thing you check ten times a day. A `mxck:` cell on the wallpaper gives you the answer without `dig`-ing or hopping to an online checker.

Combined with `L: <cell>, 60m` you can poll periodically and notice when an MX record changes unexpectedly.

## Build

Requires .NET 9. Zero NuGet dependencies — only BCL `System.Net.Http` and `System.Text.Json`.

```
dotnet build MxckExtension.csproj
```

## Notes

- Uses https://dns.google/resolve (free public DoH). No API key.
- Results cached for 1 hour in-process to be polite to the resolver.
- Trailing dots are stripped from hostnames for readable display.
- DNSSEC validation status is not reported — use a real DNS tool if you need that.

## Protocol

Reads JSON lines on stdin, writes JSON lines on stdout. See QuickSheet's main README for the full extension protocol spec.

## License

MIT
