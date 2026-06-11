# quicksheet-cal-ext

A mini text month-view calendar extension for [QuickSheet](https://github.com/cemheren/QuickSheet).
Renders a month grid directly in your desktop cells — always-visible date reference
without switching windows. Today is highlighted with `[brackets]`.

## Install

From any QuickSheet cell:

```
ext: github:cemheren/quicksheet-cal-ext
```

## Usage

| Cell content       | Result                     |
|--------------------|----------------------------|
| `mcal:`            | Current month              |
| `mcal: 2026-03`   | March 2026                 |
| `mcal: next`       | Next month                 |
| `mcal: prev`       | Previous month             |
| `mcal: 8`          | August (current year)      |

Today's date is highlighted with `[brackets]`.

## Example output

```
   June 2026
Mo Tu We Th Fr Sa Su
 1  2  3  4  5  6  7
 8  9[10]11 12 13 14
15 16 17 18 19 20 21
22 23 24 25 26 27 28
29 30
```

## Requirements

- .NET 9 SDK
- QuickSheet with extension support

## License

MIT
