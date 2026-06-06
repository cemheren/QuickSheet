# quicksheet-ip-ext

IP geolocation lookup for [QuickSheet](https://github.com/cemheren/QuickSheet). Shows country, city, ISP, and timezone for any IP address — or your own public IP.

## Install

In any QuickSheet cell:

```
ext: github:cemheren/quicksheet-ip-ext
```

## Usage

```
ip: 8.8.8.8, 1, 4
```

Parameters: `ip: <address>, <start_row>, <num_rows>`

- **address** — IPv4/IPv6 to look up. Leave blank for your own public IP.
- **start_row** — grid row to begin output (default: 1).
- **num_rows** — number of rows to fill (default: 4).

## Output

```
🌐 8.8.8.8
Mountain View, California
United States (US)
Google LLC
```

With 5+ rows, also shows timezone. With 6+, shows AS number.

## API

Uses [ip-api.com](http://ip-api.com) (free tier, no API key, 45 requests/minute limit). Responses are not cached — each activation does a fresh lookup.

## Requirements

- .NET 9 SDK
- Network access to `ip-api.com`

## License

MIT
