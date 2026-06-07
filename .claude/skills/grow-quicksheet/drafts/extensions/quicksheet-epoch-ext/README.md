# quicksheet-epoch-ext

Unix epoch ↔ human date converter for [QuickSheet](https://github.com/cemheren/QuickSheet). Convert timestamps, dates, and relative offsets directly on your desktop wallpaper.

## Install

In any QuickSheet cell:

```
ext: github:cemheren/quicksheet-epoch-ext
```

## Usage

| Cell content | Output |
|---|---|
| `epoch: now` | Current epoch seconds + UTC date + milliseconds |
| `epoch: 1717740000` | `2024-06-07 07:20:00 UTC` + relative age |
| `epoch: 1717740000000` | Auto-detects milliseconds |
| `epoch: 2026-06-07` | Epoch seconds + milliseconds for that date |
| `epoch: +30d` | Date 30 days from now as epoch + readable |
| `epoch: -7d` | Date 7 days ago |
| `epoch: +2h` | 2 hours from now |
| `epoch: -30m` | 30 minutes ago |

## Why

- Debugging APIs that return epoch timestamps
- Calculating token/cert expiry without leaving your desktop
- Quick relative date math for incident timelines
- No browser tab needed for "epoch converter"

## Requirements

- .NET 9 SDK
- QuickSheet with extension support

## License

MIT
