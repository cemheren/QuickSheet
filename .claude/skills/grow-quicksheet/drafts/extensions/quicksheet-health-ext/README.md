# quicksheet-health-ext

A QuickSheet extension that fills a grid of service-health rows on your wallpaper.
Aimed at homelab / self-hosted dashboards: Plex, Pi-hole, Nextcloud, *arr-stack, etc.

```
ext: github:cemheren/quicksheet-health-ext
health: plex=https://plex.lan,pi-hole=http://pi.hole/admin, 4, 5
```

Or point at a CSV next to your QuickSheet data:

```
health: ~/services.csv, 4, 10
```

`services.csv` layout (header optional, columns `name,url`):

```
plex,https://plex.lan
pi-hole,http://pi.hole/admin
nextcloud,https://cloud.example.com
sonarr,http://localhost:8989
```

## What you get

One row per service, four columns:

| name        | indicator | code | latency |
|-------------|-----------|------|---------|
| plex        | ✓         | 200  | 23 ms   |
| pi-hole     | ✓         | 200  | 5 ms    |
| nextcloud   | ⚠         | 401  | 87 ms   |
| sonarr      | ✗         | —    | 8000 ms |

Indicators:
- `✓` 2xx / 3xx
- `⚠` 4xx (often "this exists but needs auth" — still up)
- `✗` 5xx or unreachable

Latency is wall-clock from request start to first response header. HEAD first, GET fallback if the server rejects HEAD.

## Pair with the homelab guide

See [QuickSheet's audience guide for homelabbers](https://github.com/cemheren/QuickSheet/blob/main/docs/for-homelab.md) for the wider dashboard layout — combine `health:` with already-shipped `tls:` (cert expiry), `ping:` (latency probe), `sysmon:` (CPU/RAM/disk), `mxck:` (DNS), and so on.

## Honesty

- **Not a Grafana replacement.** This is a quick "is it up?" indicator strip; no alerting, no history, no metrics.
- **No HTTPS cert verification configuration.** Self-signed homelab certs will fail with `✗`; use `tls:` separately for cert expiry monitoring or expose your services over a real cert (e.g. Caddy → Let's Encrypt).
- **No auth.** Hits the bare URL with no headers. Use HEAD-friendly endpoints or expect a `⚠` on protected paths.

## Build

```bash
dotnet build HealthExtension.csproj
```

.NET 9, MIT, zero NuGet dependencies.

## Protocol

Standard QuickSheet JSON-lines protocol over stdin/stdout:

1. On startup, emit `{"type":"register","prefix":"health","name":"Service Health","version":"1.0.0"}`.
2. On each `{"type":"activate","id":"...","params":["<spec>"],"gridRows":N}`, parse `<spec>` (inline `name=url,name=url,...` or a path to a CSV), probe each service, and reply with `{"type":"write","id":"...","cells":[{"r":0,"c":0,"v":"plex"}, ...]}`.

See [docs/extension-protocol.md](https://github.com/cemheren/QuickSheet/blob/main/docs/extension-protocol.md) in the main repo.

## License

MIT.
