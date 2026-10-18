# quicksheet-moon-ext

A [QuickSheet](https://github.com/cemheren/QuickSheet) extension that turns a cell
into a live lunar-phase widget — phase emoji, name, illumination %, and the moon's
age in days. Great as an ambient cell on your desktop-wallpaper grid.

It is **fully offline and deterministic**: the phase is computed locally from a
standard synodic-month approximation anchored to a known new moon. No network, no
API key, no rate limits — so it never flickers or goes blank when you're offline,
which matters for a cell that lives on your wallpaper 24/7.

## Install

```
ext: github:Deskworks/quicksheet-moon-ext
```

## Usage

```
moon:                  → today's phase: emoji, name, illumination %, age in days
moon: today            → same as bare moon:
moon: 2026-12-25       → phase on a specific date (YYYY-MM-DD)
moon: next full        → date of the next full moon
moon: next new         → date of the next new moon
```

### Example output

| Cell            | Result                                          |
|-----------------|-------------------------------------------------|
| `moon:`         | `🌓 First Quarter  ·  37% lit  ·  age 6.2d`      |
| `moon: 2026-12-25` | `🌕 Full Moon  ·  99% lit  ·  age 15.7d`      |
| `moon: next full`  | `🌕 next full moon: 2026-06-29  (8d away)`    |
| `moon: next new`   | `🌑 next new moon: 2026-07-14  (23d away)`    |

The eight principal phases each map to the matching Unicode moon glyph
(🌑 🌒 🌓 🌔 🌕 🌖 🌗 🌘), so a column of dates renders as a little phase strip.

> _Screenshot placeholder — drop a `--desktop` capture of a moon: cell here._

## Why

- **Ricers / `r/unixporn`** — an always-on moon glyph that tracks the real sky is a
  clean, low-noise wallpaper widget that needs zero setup.
- **Photographers & gardeners** — quick "is it near full?" / "when's the next new
  moon?" lookups without opening an app.
- **Anyone who just likes the moon** — `=moon:` is a one-keystroke sky check.

## Accuracy

Uses the mean synodic month (29.530589 days) from a reference new moon
(2000-01-06 18:14 UTC). Phase name and illumination are accurate to well within a
day for any modern date — plenty for an ambient widget. It does not model the
small libration/perigee perturbations an ephemeris would.

## Protocol

Speaks QuickSheet's JSON-lines extension protocol: emits a `register` message on
startup, then responds to each `activate` with a `write`. See the
[extension protocol spec](https://github.com/cemheren/QuickSheet/blob/main/docs/extension-protocol.md).

## Build

```
dotnet build MoonExtension.csproj
```

Zero NuGet dependencies. .NET 9.

## License

MIT — see [LICENSE](LICENSE).
