# quicksheet-dns-ext

DNS resolver extension for [QuickSheet](https://github.com/cemheren/QuickSheet). Resolves hostnames to IP addresses directly from your desktop spreadsheet.

## Install

From inside QuickSheet, type in any cell:

```
ext: github:cemheren/quicksheet-dns-ext
```

## Usage

```
dns: example.com           → A and AAAA records
dns: 8.8.8.8, reverse     → PTR (reverse) lookup
dns: github.com            → shows IPv4 + IPv6 addresses
```

## Output

```
⟨dns⟩ github.com
A    140.82.121.3
AAAA 2606:50c0:8000::64
```

## Pairs with

- [`ping:`](https://github.com/cemheren/quicksheet-ping-ext) — HTTP probe + latency
- [`tls:`](https://github.com/cemheren/quicksheet-tls-ext) — certificate expiry
- [`tracert:`](https://github.com/cemheren/quicksheet-tracert-ext) — network path trace

## Requirements

- .NET 9 SDK
- Zero external dependencies

## License

MIT
