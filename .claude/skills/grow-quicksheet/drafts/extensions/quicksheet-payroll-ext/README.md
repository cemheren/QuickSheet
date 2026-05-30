# quicksheet-payroll-ext

A QuickSheet extension that estimates US payroll withholding per pay period. Type a gross pay amount in a cell and get a breakdown of federal income tax, Social Security, Medicare, and net take-home.

```
ext: github:cemheren/quicksheet-payroll-ext
payroll: 3846.15
payroll: 3846.15, biweekly, single
payroll: 8333.33, monthly, married
payroll: 100000, annual, single
```

## What it calculates

| Line | Formula |
|------|---------|
| **Federal income tax** | 2025 IRS percentage method: annualize gross → subtract standard deduction → apply progressive brackets → de-annualize. |
| **Social Security** | 6.2% of gross, capped at $176,100 wage base (2025). |
| **Medicare** | 1.45% of gross + additional 0.9% above $200k. |
| **Net pay** | Gross minus all three withholdings. |

## Parameters

```
payroll: <gross>, [frequency], [filing_status]
```

| Param | Options | Default |
|-------|---------|---------|
| `gross` | Any number (dollars per pay period) | *required* |
| `frequency` | `weekly` `biweekly` `semimonthly` `monthly` `annual` (or abbreviations: `w` `bw` `sm` `m` `a`) | `biweekly` |
| `filing_status` | `single` `married` (or `mfj` `joint`) | `single` |

## Example output

```
payroll: 3846.15, biweekly, single
```

Renders five cells vertically:

```
Gross: $3,846.15 (biweekly, single)
Fed tax: -$525.78
SS 6.2%: -$238.46
Medicare: -$55.77
Net pay: $3,026.14
```

## What this does NOT cover

- State income tax (varies by state; consider pairing with `salestax:` for state-level estimates).
- Pre-tax deductions (401k, HSA, insurance).
- Employer-side taxes (FUTA, SUTA).
- W-4 allowances / extra withholding.
- Non-US jurisdictions.

This is a quick sanity-check tool, not payroll software. **Not tax advice.**

## Build

```bash
dotnet build QuickSheetPayroll.csproj
```

.NET 9, MIT, zero NuGet dependencies.

## Protocol

Standard QuickSheet JSON-lines:

1. On startup, emit `{"type":"register","prefix":"payroll","version":"1.0.0"}`.
2. On each `{"type":"activate","id":"...","params":["3846.15, biweekly, single"]}`, compute withholding and reply with `{"type":"write","id":"...","cells":[...]}`.

See [docs/extension-protocol.md](https://github.com/cemheren/QuickSheet/blob/main/docs/extension-protocol.md).

## License

MIT.
