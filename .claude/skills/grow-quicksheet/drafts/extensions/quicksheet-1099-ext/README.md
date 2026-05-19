# quicksheet-1099-ext

US self-employment tax estimator for [QuickSheet](https://github.com/cemheren/QuickSheet). Pure math, no network.

## Install

```
ext: github:Deskworks/quicksheet-1099-ext
```

## Use

```
1099: 80000, 1, 5
```

Fills 5 rows:

```
$80,000 net 1099 income
SE tax: ~$11,304 (estimate)
Quarterly: ~$2,826
+ federal income tax (varies)
Not tax advice.
```

Pass your **net** 1099 income (after deductible expenses). The extension applies the standard self-employment-tax formula:

- Multiply by 0.9235 (the deductible-half adjustment).
- 12.4% Social Security up to the annual wage base ($176,100 for 2025).
- 2.9% Medicare with no cap.
- Sum, then divide by 4 for the quarterly estimated payment.

## What this is not

- **Not tax advice.** Edge cases — additional Medicare surtax for high earners, deductions, retirement contributions, state self-employment tax, filing status — are not modeled. The number is a planning estimate to know roughly what to set aside, not what to put on a 1040-ES form.
- **Not federal income tax.** That depends on your bracket, filing status, deductions, and other income. Use a real tax tool or CPA for the income-tax portion.

## Why

If you freelance, you set aside money for taxes every time a check clears. A `1099:` cell on the wallpaper gives you the rough number to move into a separate account, instantly. It's the calculation I find myself doing in mental math.

## Build

Requires .NET 9. Zero NuGet dependencies — pure math, BCL only.

```
dotnet build Ten99Extension.csproj
```

## Tuning per year

The Social Security wage base bumps annually. Update the constant in `Program.cs`:

```csharp
private const double SocialSecurityWageBase = 176100.0;
```

If the file is out of date by a year and the wage base has gone up, the estimate will be slightly low for net incomes above the old base.

## License

MIT
