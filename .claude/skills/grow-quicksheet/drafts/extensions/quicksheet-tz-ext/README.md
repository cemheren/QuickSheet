# quicksheet-tz-ext

Timezone converter for [QuickSheet](https://github.com/cemheren/QuickSheet) — see any time across multiple world zones instantly on your desktop.

## Install

Type into any QuickSheet cell:

```
ext: github:cemheren/quicksheet-tz-ext
```

## Usage

```
tz: 3pm EST
tz: 14:30 Europe/London
tz: now
```

The extension converts the input time to UTC, New York, London, Berlin, Tokyo, and Sydney — giving you a quick glance at your distributed team's local time.

### Supported zone formats

| Format | Example |
|--------|---------|
| Abbreviation | `EST`, `PST`, `UTC`, `GMT`, `CET`, `IST`, `JST` |
| IANA ID | `America/New_York`, `Europe/London`, `Asia/Tokyo` |

### Time formats

| Format | Example |
|--------|---------|
| 12-hour | `3pm`, `3:00pm`, `11:30am` |
| 24-hour | `15:00`, `23:30`, `9:00` |
| Current | `now` |

## Example output

```
🕐 3:00 PM EST
  UTC: 8:00 PM
  London: 8:00 PM
  Berlin: 9:00 PM
  Tokyo: 5:00 AM
  Sydney: 6:00 AM
```

## Requirements

- .NET 9 SDK
- QuickSheet with extension support

## License

MIT
