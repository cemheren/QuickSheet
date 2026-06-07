# quicksheet-cron-ext

Cron expression parser for [QuickSheet](https://github.com/cemheren/QuickSheet). Translates cron schedules into plain English and shows the next run time — right on your desktop wallpaper.

## Install

Add to any cell in QuickSheet:

```
ext: github:cemheren/quicksheet-cron-ext
```

## Usage

```
cron: */5 * * * *        → "Every 5 min" + next run
cron: 0 9 * * 1-5        → "At 9:00 AM, on Mon–Fri"
cron: 30 2 1 * *          → "At 2:30 AM, on day 1"
cron: @daily              → "Once a day (midnight)"
cron: 0 */2 * * *        → "At :00 every 2 hrs"
```

## Supported Syntax

- Standard 5-field cron: `minute hour day-of-month month day-of-week`
- Shortcuts: `@yearly`, `@monthly`, `@weekly`, `@daily`, `@hourly`, `@reboot`
- Ranges (`1-5`), lists (`1,3,5`), steps (`*/10`, `0-30/5`)
- Day-of-week names: `mon`, `tue`, `wed`, `thu`, `fri`, `sat`, `sun`

## Output

| Row | Content |
|-----|---------|
| 1   | ⏰ `<expression>` |
| 2   | Human-readable description |
| 3   | Next occurrence with countdown |

## Why

SRE dashboards often track scheduled jobs. Instead of memorizing `0 4 * * 0` or alt-tabbing to crontab.guru, glance at your wallpaper. Pairs well with `ping:`, `tls:`, and `gha:` extensions for a complete ops dashboard.

## Requirements

- .NET 9+ runtime
- QuickSheet with extension protocol v1+

## License

MIT
