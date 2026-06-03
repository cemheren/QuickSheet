using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace QuickSheetDepr;

/// <summary>
/// QuickSheet extension: asset depreciation calculator.
/// Supports straight-line and MACRS (half-year convention) methods.
/// Input: depr: cost, life, method[, salvage]
/// Output: year-by-year depreciation schedule. Pure math, no network.
/// Not tax advice.
/// </summary>
class Program
{
    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    // MACRS half-year convention percentages from IRS Publication 946, Table A-1.
    // Key = recovery period in years, Value = array of annual depreciation percentages.
    private static readonly Dictionary<int, double[]> MacrsRates = new()
    {
        { 3, new[] { 33.33, 44.45, 14.81, 7.41 } },
        { 5, new[] { 20.00, 32.00, 19.20, 11.52, 11.52, 5.76 } },
        { 7, new[] { 14.29, 24.49, 17.49, 12.49, 8.93, 8.92, 8.93, 4.46 } },
        { 10, new[] { 10.00, 18.00, 14.40, 11.52, 9.22, 7.37, 6.55, 6.55, 6.56, 6.55, 3.28 } },
        { 15, new[] { 5.00, 9.50, 8.55, 7.70, 6.93, 6.23, 5.90, 5.90, 5.91, 5.90, 5.91, 5.90, 5.91, 5.90, 5.91, 2.95 } },
        { 20, new[] { 3.750, 7.219, 6.677, 6.177, 5.713, 5.285, 4.888, 4.522, 4.462, 4.461, 4.462, 4.461, 4.462, 4.461, 4.462, 4.461, 4.462, 4.461, 4.462, 4.461, 2.231 } }
    };

    static async Task Main()
    {
        using var reader = new StreamReader(Console.OpenStandardInput(), Encoding.UTF8);

        var register = new { type = "register", prefix = "depr", version = "1.0.0" };
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

        if (parts.Length < 2)
        {
            WriteCells(id, new[] { new[] { "depr: <cost>, <years>, [straight|macrs], [salvage]" } });
            return;
        }

        if (!decimal.TryParse(parts[0], NumberStyles.Number, CultureInfo.InvariantCulture, out var cost) || cost <= 0)
        {
            WriteCells(id, new[] { new[] { "Error: invalid cost" } });
            return;
        }

        if (!int.TryParse(parts[1], out var life) || life <= 0)
        {
            WriteCells(id, new[] { new[] { "Error: invalid life (years)" } });
            return;
        }

        string method = parts.Length > 2 ? parts[2].ToLowerInvariant() : "straight";

        decimal salvage = 0m;
        if (parts.Length > 3 && decimal.TryParse(parts[3], NumberStyles.Number, CultureInfo.InvariantCulture, out var s))
            salvage = s;

        if (method is "macrs" or "m")
            HandleMacrs(id, cost, life);
        else
            HandleStraightLine(id, cost, life, salvage);
    }

    static void HandleStraightLine(string id, decimal cost, int life, decimal salvage)
    {
        decimal depreciable = cost - salvage;
        decimal annual = Math.Round(depreciable / life, 2);

        var rows = new List<string[]>
        {
            new[] { "Yr", "Depreciation", "Book Value" }
        };

        decimal bookValue = cost;
        for (int yr = 1; yr <= life; yr++)
        {
            decimal dep = (yr == life) ? bookValue - salvage : annual;
            bookValue -= dep;
            rows.Add(new[] { yr.ToString(), $"${dep:N2}", $"${bookValue:N2}" });
        }

        rows.Add(new[] { "", $"Total: ${depreciable:N2}", $"Salvage: ${salvage:N2}" });
        WriteCells(id, rows);
    }

    static void HandleMacrs(string id, decimal cost, int life)
    {
        // Find closest MACRS recovery period
        int recovery = FindMacrsRecovery(life);
        if (!MacrsRates.TryGetValue(recovery, out var rates))
        {
            WriteCells(id, new[] { new[] { $"Error: no MACRS table for {life}-yr life (use 3/5/7/10/15/20)" } });
            return;
        }

        var rows = new List<string[]>
        {
            new[] { "Yr", "Rate", "Depreciation", "Book Value" }
        };

        decimal bookValue = cost;
        decimal totalDep = 0;
        for (int i = 0; i < rates.Length; i++)
        {
            decimal rate = (decimal)rates[i] / 100m;
            decimal dep = Math.Round(cost * rate, 2);
            totalDep += dep;
            bookValue = cost - totalDep;
            if (bookValue < 0) { dep += bookValue; totalDep += bookValue; bookValue = 0; }
            rows.Add(new[] { (i + 1).ToString(), $"{rates[i]:F2}%", $"${dep:N2}", $"${bookValue:N2}" });
        }

        rows.Add(new[] { "", "", $"Total: ${totalDep:N2}", $"({recovery}-yr MACRS)" });
        WriteCells(id, rows);
    }

    static int FindMacrsRecovery(int life)
    {
        int[] periods = { 3, 5, 7, 10, 15, 20 };
        foreach (var p in periods)
            if (life <= p) return p;
        return 20;
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

        var response = new { type = "result", id, cells };
        Console.WriteLine(JsonSerializer.Serialize(response, JsonOpts));
    }
}
