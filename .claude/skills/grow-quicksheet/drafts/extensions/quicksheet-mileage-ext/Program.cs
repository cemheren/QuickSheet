using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace QuickSheetMileage;

/// <summary>
/// QuickSheet extension: IRS standard-mileage deduction calculator.
/// Inputs: miles, mode (business/medical/charity), optional tax-year.
/// Output (3 cells): miles, rate, deduction. Pure math, no network.
/// Not tax advice.
/// </summary>
class Program
{
    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    // IRS standard mileage rates (USD per mile) by tax year.
    // Source: irs.gov/tax-professionals/standard-mileage-rates.
    private static readonly Dictionary<int, (decimal business, decimal medical, decimal charity)> Rates = new()
    {
        { 2021, (0.560m, 0.160m, 0.140m) },
        { 2022, (0.625m, 0.220m, 0.140m) }, // mid-year change: Jan-Jun .585, Jul-Dec .625 (business). Use H2 as default.
        { 2023, (0.655m, 0.220m, 0.140m) },
        { 2024, (0.670m, 0.210m, 0.140m) },
        { 2025, (0.700m, 0.210m, 0.140m) }
    };

    private const int LatestYear = 2025;

    static async Task Main()
    {
        using var reader = new StreamReader(Console.OpenStandardInput(), Encoding.UTF8);

        var register = new { type = "register", prefix = "mileage", version = "1.0.0" };
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

        if (parts.Length == 0 || !decimal.TryParse(parts[0], NumberStyles.Number, CultureInfo.InvariantCulture, out var miles))
        {
            WriteCells(id, new[] { new[] { "mileage: <miles>, <mode>, [<year>]" } });
            return;
        }

        string mode = parts.Length > 1 ? parts[1].ToLowerInvariant() : "business";

        int year = LatestYear;
        if (parts.Length > 2 && int.TryParse(parts[2], out var y) && Rates.ContainsKey(y))
            year = y;

        if (!Rates.TryGetValue(year, out var rateSet))
            rateSet = Rates[LatestYear];

        decimal rate = mode switch
        {
            "business" or "biz" => rateSet.business,
            "medical" or "med" or "moving" => rateSet.medical,
            "charity" or "char" => rateSet.charity,
            _ => rateSet.business
        };

        decimal deduction = miles * rate;
        string modeLabel = mode switch
        {
            "medical" or "med" or "moving" => "medical",
            "charity" or "char" => "charity",
            _ => "business"
        };

        var rows = new[]
        {
            new[] { $"{miles:N0} mi {modeLabel}" },
            new[] { $"× ${rate:F3}/mi (IRS {year})" },
            new[] { $"= ${deduction:N2} deduction" }
        };

        WriteCells(id, rows);
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
