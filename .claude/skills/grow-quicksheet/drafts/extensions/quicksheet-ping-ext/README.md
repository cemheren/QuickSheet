# quicksheet-ping-ext

HTTP ping (status + latency) for [QuickSheet](https://github.com/cemheren/QuickSheet).

## Install

Type into any cell:

```
ext: github:Deskworks/quicksheet-ping-ext
```

## Use

```
ping: https://example.com, 1, 3
```

Fills 3 rows:

```
✓ https://example.com
200 OK
142 ms
```

Indicators: `✓` 2xx/3xx, `⚠` 4xx, `✗` 5xx, `?` other. URL scheme defaults to `https://` if you omit it.

Pair it with `L: <cell>, 1m` for a one-minute uptime poll right on the wallpaper.

## Why

A grid of `ping:` cells next to friendly labels is a no-config status page — pinned to the wallpaper, never closed, always behind whatever you're doing.

## Build

Requires .NET 9. Zero NuGet dependencies — only BCL `System.Net.Http`, `System.Diagnostics`, `System.Text.Json`.

```
dotnet build PingExtension.csproj
```

## Notes

- HEAD request first; falls back to GET if the server rejects HEAD.
- 8 second timeout.
- No long-lived cache — this is meant to be an active probe. Use `L:` to schedule.
- Redirects are NOT followed; you get the real first-hop status.

## Protocol

Reads JSON lines on stdin, writes JSON lines on stdout. See QuickSheet's main README for the full extension protocol spec.

## License

MIT
