# quicksheet-unit-ext

Instant unit conversion on your desktop wallpaper. Type a conversion in any cell and get the result without leaving your spreadsheet.

```
ext: github:cemheren/quicksheet-unit-ext
unit: 5 km to miles
unit: 100 F to C
unit: 2.5 lb to kg
unit: 1 GB to MB
unit: 60 mph to km/h
unit: 1 cup to ml
unit: 500 sqft to sqm
```

## Supported categories

| Category    | Units                                                                 |
|-------------|-----------------------------------------------------------------------|
| Length      | m, km, cm, mm, mi (miles), yd, ft, in, nmi (nautical miles)          |
| Mass        | kg, g, mg, lb, oz, t (metric ton), st (stone)                        |
| Temperature | C (Celsius), F (Fahrenheit), K (Kelvin)                              |
| Volume      | l, ml, gal, qt, pt, cup, tbsp, tsp, floz                            |
| Speed       | m/s, km/h, mph, kt (knots)                                          |
| Data        | B, KB, MB, GB, TB, KiB, MiB, GiB                                    |
| Area        | sqm (m²), sqft (ft²), acre, ha (hectare)                            |
| Time        | sec, min, hr, day, wk                                                |

## Syntax

```
unit: <number> <from-unit> to <to-unit>
```

- Numbers can use commas: `unit: 1,000 ft to m`
- Unit names are case-insensitive and accept plurals: `miles`, `Mile`, `MILES` all work
- Common abbreviations and full names accepted: `kilometer`, `km`, `kilometres`

## Install

Add to any cell in QuickSheet:

```
ext: github:cemheren/quicksheet-unit-ext
```

Then use `unit:` in any cell.

## Build

```bash
dotnet build UnitExtension.csproj
```

.NET 9, MIT, zero NuGet dependencies. Pure computation — no network calls.

## Why on a wallpaper?

- Engineering students: keep a live conversion cheat-sheet visible while studying
- Cooking: recipe scaling without alt-tabbing to Google
- International teams: quick metric ↔ imperial without context-switching
- Data work: "is 1.5 TB enough for this dataset in GiB?"

## Protocol

Standard QuickSheet JSON-lines:

1. On startup, emit `{"type":"register","prefix":"unit","name":"Unit Converter","version":"1.0.0"}`.
2. On `{"type":"activate","id":"...","params":["5 km to miles"]}`, compute and reply with `{"type":"write","id":"...","cells":[{"r":0,"c":0,"v":"5 km = 3.10686 miles"}]}`.

See [docs/extension-protocol.md](https://github.com/cemheren/QuickSheet/blob/main/docs/extension-protocol.md).

## License

MIT.
