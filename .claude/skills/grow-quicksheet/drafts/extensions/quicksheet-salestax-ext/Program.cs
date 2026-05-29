using System.Text.Json;
using System.Text.Json.Serialization;

/// <summary>
/// QuickSheet Sales Tax Extension — US state sales tax rate lookup.
/// Reads JSON-lines from stdin, writes JSON-lines to stdout.
/// Usage: `tax: CA` → California 7.25%, `tax: CA, 100` → $7.25 on $100.
/// Supports state abbreviations and full names. Includes combined state+avg local rates.
/// </summary>
class Program
{
    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        WriteIndented = false
    };

    // State sales tax rates (state rate + average local rate combined), 2024 data.
    // Source: Tax Foundation. States with 0% have no general sales tax.
    private static readonly Dictionary<string, (string Name, decimal StateRate, decimal AvgCombined)> TaxRates = new(StringComparer.OrdinalIgnoreCase)
    {
        ["AL"] = ("Alabama", 4.00m, 9.29m),
        ["AK"] = ("Alaska", 0.00m, 1.82m),
        ["AZ"] = ("Arizona", 5.60m, 8.37m),
        ["AR"] = ("Arkansas", 6.50m, 9.46m),
        ["CA"] = ("California", 7.25m, 8.85m),
        ["CO"] = ("Colorado", 2.90m, 7.83m),
        ["CT"] = ("Connecticut", 6.35m, 6.35m),
        ["DE"] = ("Delaware", 0.00m, 0.00m),
        ["FL"] = ("Florida", 6.00m, 7.02m),
        ["GA"] = ("Georgia", 4.00m, 7.38m),
        ["HI"] = ("Hawaii", 4.00m, 4.44m),
        ["ID"] = ("Idaho", 6.00m, 6.02m),
        ["IL"] = ("Illinois", 6.25m, 8.84m),
        ["IN"] = ("Indiana", 7.00m, 7.00m),
        ["IA"] = ("Iowa", 6.00m, 6.94m),
        ["KS"] = ("Kansas", 6.50m, 8.72m),
        ["KY"] = ("Kentucky", 6.00m, 6.00m),
        ["LA"] = ("Louisiana", 4.45m, 9.56m),
        ["ME"] = ("Maine", 5.50m, 5.50m),
        ["MD"] = ("Maryland", 6.00m, 6.00m),
        ["MA"] = ("Massachusetts", 6.25m, 6.25m),
        ["MI"] = ("Michigan", 6.00m, 6.00m),
        ["MN"] = ("Minnesota", 6.875m, 7.49m),
        ["MS"] = ("Mississippi", 7.00m, 7.07m),
        ["MO"] = ("Missouri", 4.225m, 8.40m),
        ["MT"] = ("Montana", 0.00m, 0.00m),
        ["NE"] = ("Nebraska", 5.50m, 7.13m),
        ["NV"] = ("Nevada", 6.85m, 8.24m),
        ["NH"] = ("New Hampshire", 0.00m, 0.00m),
        ["NJ"] = ("New Jersey", 6.625m, 6.60m),
        ["NM"] = ("New Mexico", 4.875m, 7.72m),
        ["NY"] = ("New York", 4.00m, 8.53m),
        ["NC"] = ("North Carolina", 4.75m, 6.99m),
        ["ND"] = ("North Dakota", 5.00m, 7.01m),
        ["OH"] = ("Ohio", 5.75m, 7.25m),
        ["OK"] = ("Oklahoma", 4.50m, 8.99m),
        ["OR"] = ("Oregon", 0.00m, 0.00m),
        ["PA"] = ("Pennsylvania", 6.00m, 6.34m),
        ["RI"] = ("Rhode Island", 7.00m, 7.00m),
        ["SC"] = ("South Carolina", 6.00m, 7.47m),
        ["SD"] = ("South Dakota", 4.20m, 6.40m),
        ["TN"] = ("Tennessee", 7.00m, 9.55m),
        ["TX"] = ("Texas", 6.25m, 8.20m),
        ["UT"] = ("Utah", 6.10m, 7.19m),
        ["VT"] = ("Vermont", 6.00m, 6.36m),
        ["VA"] = ("Virginia", 5.30m, 5.75m),
        ["WA"] = ("Washington", 6.50m, 10.37m),
        ["WV"] = ("West Virginia", 6.00m, 6.57m),
        ["WI"] = ("Wisconsin", 5.00m, 5.43m),
        ["WY"] = ("Wyoming", 4.00m, 5.36m),
        ["DC"] = ("District of Columbia", 6.00m, 6.00m),
    };

    // Reverse lookup: full state name → abbreviation
    private static readonly Dictionary<string, string> NameToAbbrev;

    static Program()
    {
        NameToAbbrev = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var kv in TaxRates)
            NameToAbbrev[kv.Value.Name] = kv.Key;
    }

    static void Main()
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;
        string? line;
        while ((line = Console.ReadLine()) != null)
        {
            if (string.IsNullOrWhiteSpace(line)) continue;
            try
            {
                using var doc = JsonDocument.Parse(line);
                string? type = doc.RootElement.TryGetProperty("type", out var tp) ? tp.GetString() : null;
                switch (type)
                {
                    case "init": HandleInit(); break;
                    case "activate": HandleActivate(doc.RootElement); break;
                }
            }
            catch (Exception ex)
            {
                SendLog($"parse error: {ex.Message}");
            }
        }
    }

    static void HandleInit()
    {
        SendJson(new
        {
            type = "register",
            prefix = "tax",
            name = "US Sales Tax",
            version = "1.0.0"
        });
        SendLog("Sales Tax extension registered with prefix 'tax'");
    }

    static void HandleActivate(JsonElement root)
    {
        string id = root.TryGetProperty("id", out var idProp) ? idProp.GetString() ?? "" : "";
        int gridRows = root.TryGetProperty("gridRows", out var gr) ? gr.GetInt32() : 4;

        string[] extParams = [];
        if (root.TryGetProperty("params", out var p) && p.ValueKind == JsonValueKind.Array)
            extParams = p.EnumerateArray().Select(x => x.GetString() ?? "").ToArray();

        if (extParams.Length == 0 || string.IsNullOrWhiteSpace(extParams[0]))
        {
            WriteCells(id, new[] { new[] { "tax: <state> [, amount]" } });
            return;
        }

        string stateInput = extParams[0].Trim();
        decimal? amount = null;
        if (extParams.Length >= 2 && decimal.TryParse(extParams[1].Trim(), out var amt))
            amount = amt;

        // Resolve state abbreviation
        string? abbrev = ResolveState(stateInput);
        if (abbrev == null)
        {
            WriteCells(id, new[] { new[] { $"Unknown state: {stateInput}" } });
            return;
        }

        var (name, stateRate, avgCombined) = TaxRates[abbrev];

        var rows = new List<string[]>();
        rows.Add(new[] { $"{name} ({abbrev})" });
        rows.Add(new[] { $"State: {stateRate:0.###}%  Combined: {avgCombined:0.##}%" });

        if (amount.HasValue)
        {
            decimal stateTax = Math.Round(amount.Value * stateRate / 100m, 2);
            decimal combinedTax = Math.Round(amount.Value * avgCombined / 100m, 2);
            decimal total = amount.Value + combinedTax;
            rows.Add(new[] { $"Tax on ${amount.Value:0.##}: ${combinedTax:0.00}" });
            rows.Add(new[] { $"Total: ${total:0.00}" });
        }
        else
        {
            if (stateRate == 0m && avgCombined == 0m)
                rows.Add(new[] { "No sales tax ✓" });
            else if (stateRate == 0m)
                rows.Add(new[] { "No state tax (local only)" });
            else
                rows.Add(new[] { $"On $100: ${avgCombined:0.00}" });
        }

        while (rows.Count < gridRows) rows.Add(new[] { "" });
        if (rows.Count > gridRows) rows = rows.Take(gridRows).ToList();
        WriteCells(id, rows);
    }

    static string? ResolveState(string input)
    {
        // Try as abbreviation first
        if (input.Length <= 3 && TaxRates.ContainsKey(input))
            return input.ToUpperInvariant();

        // Try as full name
        if (NameToAbbrev.TryGetValue(input, out var abbr))
            return abbr;

        // Fuzzy: try starts-with on state names
        foreach (var kv in NameToAbbrev)
        {
            if (kv.Key.StartsWith(input, StringComparison.OrdinalIgnoreCase))
                return kv.Value;
        }

        return null;
    }

    static void WriteCells(string id, IEnumerable<string[]> rows)
    {
        SendJson(new { type = "write", id, cells = rows });
    }

    static void SendJson(object obj)
    {
        Console.WriteLine(JsonSerializer.Serialize(obj, JsonOpts));
    }

    static void SendLog(string message)
    {
        SendJson(new { type = "log", message });
    }
}
