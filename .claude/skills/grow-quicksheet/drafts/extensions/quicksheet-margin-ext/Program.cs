using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace QuickSheetMargin;

/// <summary>
/// QuickSheet extension: Break-even & contribution-margin calculator.
/// Input: margin: fixedCosts, variablePerUnit, sellPrice [, units=N]
/// Output: contribution margin, break-even units, break-even revenue, margin ratio.
/// Optionally: profit at target volume.
/// </summary>
class Program
{
    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    static async Task Main()
    {
        using var reader = new StreamReader(Console.OpenStandardInput(), Encoding.UTF8);

        var register = new { type = "register", prefix = "margin", version = "1.0.0" };
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

        if (parts.Length < 3)
        {
            WriteCells(id, new[] { new[] { "margin: <fixed>, <variable/unit>, <price> [, units=N]" } });
            return;
        }

        if (!decimal.TryParse(parts[0], NumberStyles.Number, CultureInfo.InvariantCulture, out var fixedCosts) ||
            !decimal.TryParse(parts[1], NumberStyles.Number, CultureInfo.InvariantCulture, out var variablePerUnit) ||
            !decimal.TryParse(parts[2], NumberStyles.Number, CultureInfo.InvariantCulture, out var sellPrice))
        {
            WriteCells(id, new[] { new[] { "Error: all values must be numbers" } });
            return;
        }

        if (sellPrice <= 0)
        {
            WriteCells(id, new[] { new[] { "Error: sell price must be > 0" } });
            return;
        }

        decimal contribution = sellPrice - variablePerUnit;

        if (contribution <= 0)
        {
            WriteCells(id, new[] { new[] { $"CM: ${contribution:N2}/unit (negative — no break-even)" } });
            return;
        }

        decimal breakEvenUnits = Math.Ceiling(fixedCosts / contribution);
        decimal breakEvenRevenue = breakEvenUnits * sellPrice;
        decimal marginRatio = (contribution / sellPrice) * 100m;

        var rows = new List<string[]>
        {
            new[] { $"Contribution margin: ${contribution:N2}/unit" },
            new[] { $"Break-even: {breakEvenUnits:N0} units" },
            new[] { $"Break-even revenue: ${breakEvenRevenue:N2}" },
            new[] { $"Margin ratio: {marginRatio:F1}%" }
        };

        // Optional target volume: units=N
        int? targetUnits = null;
        for (int i = 3; i < parts.Length; i++)
        {
            var kv = parts[i].Split('=', 2, StringSplitOptions.TrimEntries);
            if (kv.Length == 2 && kv[0].Equals("units", StringComparison.OrdinalIgnoreCase)
                && int.TryParse(kv[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out var u))
            {
                targetUnits = u;
            }
        }

        if (targetUnits.HasValue)
        {
            decimal totalRevenue = targetUnits.Value * sellPrice;
            decimal totalCost = fixedCosts + (targetUnits.Value * variablePerUnit);
            decimal profit = totalRevenue - totalCost;
            string sign = profit >= 0 ? "+" : "";
            rows.Add(new[] { $"At {targetUnits.Value:N0} units: {sign}${profit:N2}" });
        }

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
