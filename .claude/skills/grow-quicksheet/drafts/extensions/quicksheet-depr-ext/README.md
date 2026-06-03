# quicksheet-depr-ext

Asset depreciation schedule calculator for [QuickSheet](https://github.com/cemheren/QuickSheet). Computes straight-line and MACRS (IRS Pub 946 half-year convention) depreciation tables directly in your desktop spreadsheet.

## Install

In any QuickSheet cell:

```
ext: github:cemheren/quicksheet-depr-ext
```

## Usage

### Straight-line depreciation

```
depr: 10000, 5, straight
```

Output:
```
Yr  Depreciation  Book Value
1   $2,000.00     $8,000.00
2   $2,000.00     $6,000.00
3   $2,000.00     $4,000.00
4   $2,000.00     $2,000.00
5   $2,000.00     $0.00
    Total: $10,000.00  Salvage: $0.00
```

With salvage value:
```
depr: 10000, 5, straight, 2000
```

### MACRS depreciation (IRS half-year convention)

```
depr: 10000, 5, macrs
```

Output:
```
Yr  Rate    Depreciation  Book Value
1   20.00%  $2,000.00     $8,000.00
2   32.00%  $3,200.00     $4,800.00
3   19.20%  $1,920.00     $2,880.00
4   11.52%  $1,152.00     $1,728.00
5   11.52%  $1,152.00     $576.00
6   5.76%   $576.00       $0.00
            Total: $10,000.00  (5-yr MACRS)
```

Supported MACRS recovery periods: 3, 5, 7, 10, 15, 20 years. If your life doesn't match exactly, the nearest valid period is used.

## Parameters

```
depr: <cost>, <life_years>, [method], [salvage]
```

| Parameter | Required | Description |
|-----------|----------|-------------|
| cost | Yes | Original asset cost |
| life_years | Yes | Useful life in years |
| method | No | `straight` (default) or `macrs` |
| salvage | No | Salvage value (straight-line only, default 0) |

## Examples

| Cell | Description |
|------|-------------|
| `depr: 25000, 5, macrs` | 5-year MACRS for a $25k vehicle |
| `depr: 50000, 7, macrs` | 7-year MACRS for office furniture |
| `depr: 3000, 3, straight, 500` | 3-year straight-line, $500 salvage |
| `depr: 100000, 15, macrs` | 15-year MACRS for land improvement |

## Requirements

- .NET 9 SDK
- Zero external dependencies

## Disclaimer

This extension is for informational/educational purposes only. Not tax advice. Consult a qualified tax professional for actual depreciation elections.

## License

MIT
