using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace QuickSheetPayroll;

/// <summary>
/// QuickSheet extension: payroll withholding estimator.
/// Input: gross pay per period, pay frequency, filing status.
/// Output: federal income tax withholding, Social Security, Medicare, net pay.
/// Uses 2025 IRS Publication 15-T progressive brackets (percentage method).
/// Pure math, no network, no NuGet deps. NOT tax/legal/payroll advice.
/// </summary>
class Program
{
    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    // 2025 federal income tax brackets (annual) — married filing jointly & single.
    // Source: IRS Rev. Proc. 2024-40.
    private static readonly (decimal threshold, decimal rate)[] SingleBrackets =
    {
        (0m, 0.10m),
        (11_925m, 0.12m),
        (48_475m, 0.22m),
        (103_350m, 0.24m),
        (197_300m, 0.32m),
        (250_525m, 0.35m),
        (626_350m, 0.37m)
    };

    private static readonly (decimal threshold, decimal rate)[] MarriedBrackets =
    {
        (0m, 0.10m),
        (23_850m, 0.12m),
        (96_950m, 0.22m),
        (206_700m, 0.24m),
        (394_600m, 0.32m),
        (501_050m, 0.35m),
        (751_600m, 0.37m)
    };

    // 2025 standard deduction (used in percentage method)
    private const decimal SingleStdDeduction = 15_000m;
    private const decimal MarriedStdDeduction = 30_000m;

    // FICA rates (2025)
    private const decimal SocialSecurityRate = 0.062m;
    private const decimal SocialSecurityWageCap = 176_100m;
    private const decimal MedicareRate = 0.0145m;
    private const decimal MedicareAdditionalRate = 0.009m;
    private const decimal MedicareAdditionalThreshold = 200_000m;

    // Pay periods per year
    private static readonly Dictionary<string, int> Frequencies = new(StringComparer.OrdinalIgnoreCase)
    {
        { "weekly", 52 }, { "w", 52 },
        { "biweekly", 26 }, { "bw", 26 }, { "bi-weekly", 26 },
        { "semimonthly", 24 }, { "sm", 24 }, { "semi-monthly", 24 },
        { "monthly", 12 }, { "m", 12 },
        { "annual", 1 }, { "a", 1 }, { "yearly", 1 }
    };

    static async Task Main()
    {
        using var reader = new StreamReader(Console.OpenStandardInput(), Encoding.UTF8);

        var register = new { type = "register", prefix = "payroll", version = "1.0.0" };
        Console.WriteLine(JsonSerializer.Serialize(register, JsonOpts));

        while (true)
        {
            var line = await reader.ReadLineAsync();
            if (line == null) break;
            if (string.IsNullOrWhiteSpace(line)) continue;

            try
            {
                using var doc = JsonDocument.Parse(line);
                var root = doc.RootElement;
                if (root.GetProperty("type").GetString() != "activate") continue;

                var id = root.GetProperty("id").GetString() ?? "";

                string args = "";
                if (root.TryGetProperty("params", out var p) && p.ValueKind == JsonValueKind.Array && p.GetArrayLength() > 0)
                    args = p[0].GetString() ?? "";
                else if (root.TryGetProperty("arguments", out var a))
                    args = a.GetString() ?? "";

                HandleActivate(id, args);
            }
            catch { /* malformed line — ignore */ }
        }
    }

    static void HandleActivate(string id, string args)
    {
        var parts = args.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

        if (parts.Length == 0 || !decimal.TryParse(parts[0], NumberStyles.Number, CultureInfo.InvariantCulture, out var grossPay))
        {
            WriteCells(id, new[]
            {
                new[] { "payroll: <gross>, [freq], [status]" },
                new[] { "freq: weekly|biweekly|monthly|annual" },
                new[] { "status: single|married" }
            });
            return;
        }

        // Parse frequency (default: biweekly)
        string freqKey = parts.Length > 1 ? parts[1] : "biweekly";
        if (!Frequencies.TryGetValue(freqKey, out int periodsPerYear))
            periodsPerYear = 26; // default biweekly

        // Parse filing status (default: single)
        string status = parts.Length > 2 ? parts[2].ToLowerInvariant() : "single";
        bool married = status is "married" or "mfj" or "joint";

        // Annualize
        decimal annualGross = grossPay * periodsPerYear;

        // Federal income tax (percentage method: annualize, subtract std deduction, apply brackets)
        decimal stdDeduction = married ? MarriedStdDeduction : SingleStdDeduction;
        decimal taxableIncome = Math.Max(0, annualGross - stdDeduction);
        var brackets = married ? MarriedBrackets : SingleBrackets;
        decimal annualFedTax = ComputeProgressiveTax(taxableIncome, brackets);
        decimal fedTaxPerPeriod = Math.Round(annualFedTax / periodsPerYear, 2);

        // Social Security (capped at wage base)
        decimal ssAnnual = Math.Min(annualGross, SocialSecurityWageCap) * SocialSecurityRate;
        decimal ssPerPeriod = Math.Round(ssAnnual / periodsPerYear, 2);

        // Medicare (no cap; additional 0.9% above $200k)
        decimal medAnnual = annualGross * MedicareRate;
        if (annualGross > MedicareAdditionalThreshold)
            medAnnual += (annualGross - MedicareAdditionalThreshold) * MedicareAdditionalRate;
        decimal medPerPeriod = Math.Round(medAnnual / periodsPerYear, 2);

        decimal totalWithholding = fedTaxPerPeriod + ssPerPeriod + medPerPeriod;
        decimal netPay = grossPay - totalWithholding;

        string freqLabel = freqKey.ToLowerInvariant() switch
        {
            "w" or "weekly" => "weekly",
            "bw" or "biweekly" or "bi-weekly" => "biweekly",
            "sm" or "semimonthly" or "semi-monthly" => "semimonthly",
            "m" or "monthly" => "monthly",
            "a" or "annual" or "yearly" => "annual",
            _ => "biweekly"
        };

        var rows = new[]
        {
            new[] { $"Gross: ${grossPay:N2} ({freqLabel}, {(married ? "MFJ" : "single")})" },
            new[] { $"Fed tax: -${fedTaxPerPeriod:N2}" },
            new[] { $"SS 6.2%: -${ssPerPeriod:N2}" },
            new[] { $"Medicare: -${medPerPeriod:N2}" },
            new[] { $"Net pay: ${netPay:N2}" }
        };

        WriteCells(id, rows);
    }

    static decimal ComputeProgressiveTax(decimal income, (decimal threshold, decimal rate)[] brackets)
    {
        decimal tax = 0;
        for (int i = brackets.Length - 1; i >= 0; i--)
        {
            if (income > brackets[i].threshold)
            {
                tax += (income - brackets[i].threshold) * brackets[i].rate;
                income = brackets[i].threshold;
            }
        }
        return tax;
    }

    static void WriteCells(string id, IEnumerable<string[]> rows)
    {
        var cells = new List<object>();
        int r = 0;
        foreach (var row in rows)
        {
            for (int c = 0; c < row.Length; c++)
                cells.Add(new { r, c, v = row[c] ?? "" });
            r++;
        }
        Console.WriteLine(JsonSerializer.Serialize(new { type = "write", id, cells }, JsonOpts));
    }
}
