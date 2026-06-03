# quicksheet-margin-ext

Break-even & contribution-margin calculator for [QuickSheet](https://github.com/cemheren/QuickSheet).

Computes contribution margin, break-even point, and margin ratio from fixed costs, variable cost per unit, and selling price. Optionally calculates profit at a target volume.

## Install

Add to any cell in QuickSheet:

```
ext: github:cemheren/quicksheet-margin-ext
```

## Usage

```
margin: <fixedCosts>, <variablePerUnit>, <sellPrice>
```

### Examples

```
margin: 50000, 12, 30
```

Output:
```
Contribution margin: $18.00/unit
Break-even: 2,778 units
Break-even revenue: $83,340.00
Margin ratio: 60.0%
```

With target volume:

```
margin: 50000, 12, 30, units=5000
```

Output adds:
```
At 5,000 units: +$40,000.00
```

### Parameters

| Position | Name | Description |
|----------|------|-------------|
| 1 | Fixed costs | Total fixed costs (rent, salaries, etc.) |
| 2 | Variable/unit | Variable cost per unit produced |
| 3 | Sell price | Selling price per unit |
| 4 (opt) | units=N | Target volume to compute profit |

## How it works

- **Contribution margin** = sell price − variable cost per unit
- **Break-even units** = fixed costs ÷ contribution margin (rounded up)
- **Margin ratio** = contribution margin ÷ sell price × 100%
- **Profit at N units** = (N × sell price) − fixed costs − (N × variable cost)

## Requirements

- .NET 9 SDK
- QuickSheet with extension support

## License

MIT
