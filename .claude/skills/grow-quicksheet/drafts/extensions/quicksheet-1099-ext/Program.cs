using System.Globalization;
using System.Text.Json;
using System.Text.Json.Serialization;

/// <summary>
/// QuickSheet 1099 Extension — reads JSON-lines from stdin, writes JSON-lines to stdout.
/// Registers the "1099" prefix. Estimates US self-employment tax (15.3% on 92.35%
/// of net earnings, with Social Security capped at the wage base) and the quarterly
/// payment. Federal income tax is NOT included — it depends on bracket, filing status,
/// deductions, etc.
/// Usage: `1099: 80000, 1, 5` → net income, SE tax, quarterly, notes.
/// </summary>
class Program
{
    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        WriteIndented = false
    };

    private static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

    // 2025 figures. Bump per year.
    private const double SocialSecurityWageBase = 176100.0;
    private const double SocialSecurityRate = 0.124;
    private const double MedicareRate = 0.029;
    private const double SelfEmploymentAdjustment = 0.9235;

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
            prefix = "1099",
            name = "US Self-Employment Tax Estimate",
            version = "1.0.0"
        });
        SendLog("1099 extension registered with prefix '1099'");
    }

    static void HandleActivate(JsonElement root)
    {
        string id = root.TryGetProperty("id", out var idProp) ? idProp.GetString() ?? "" : "";
        int gridRows = root.TryGetProperty("gridRows", out var gr) ? gr.GetInt32() : 5;

        string[] extParams = [];
        if (root.TryGetProperty("params", out var p) && p.ValueKind == JsonValueKind.Array)
            extParams = p.EnumerateArray().Select(x => x.GetString() ?? "").ToArray();

        if (extParams.Length == 0 || !double.TryParse(extParams[0], NumberStyles.Float, Inv, out double net) || net <= 0)
        {
            WriteCells(id, new[] { new[] { "1099: <net-1099-income>" } });
            return;
        }

        double seBase = net * SelfEmploymentAdjustment;
        double ssTax = Math.Min(seBase, SocialSecurityWageBase) * SocialSecurityRate;
        double medicareTax = seBase * MedicareRate;
        double seTax = ssTax + medicareTax;
        double quarterly = seTax / 4.0;

        var rows = new List<string[]>
        {
            new[] { $"${net.ToString("N0", Inv)} net 1099 income" },
            new[] { $"SE tax: ~${seTax.ToString("N0", Inv)} (estimate)" },
            new[] { $"Quarterly: ~${quarterly.ToString("N0", Inv)}" },
            new[] { "+ federal income tax (varies)" },
            new[] { "Not tax advice." }
        };
        while (rows.Count < gridRows) rows.Add(new[] { "" });
        if (rows.Count > gridRows) rows = rows.Take(gridRows).ToList();
        WriteCells(id, rows);
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
