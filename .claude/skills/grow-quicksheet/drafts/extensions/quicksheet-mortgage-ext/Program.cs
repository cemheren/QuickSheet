using System.Globalization;
using System.Text.Json;
using System.Text.Json.Serialization;

/// <summary>
/// QuickSheet Mortgage Extension — reads JSON-lines from stdin, writes JSON-lines to stdout.
/// Registers the "mort" prefix. Computes monthly payment, total interest, and total cost
/// for a fixed-rate fully-amortizing mortgage.
/// Usage: `mort: 500000, 6.5, 30, 1, 4` → 4 rows. Params: principal, annual rate (%), years.
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
            prefix = "mort",
            name = "Mortgage Calculator",
            version = "1.0.0"
        });
        SendLog("Mortgage extension registered with prefix 'mort'");
    }

    static void HandleActivate(JsonElement root)
    {
        string id = root.TryGetProperty("id", out var idProp) ? idProp.GetString() ?? "" : "";
        int gridRows = root.TryGetProperty("gridRows", out var gr) ? gr.GetInt32() : 4;

        string[] extParams = [];
        if (root.TryGetProperty("params", out var p) && p.ValueKind == JsonValueKind.Array)
            extParams = p.EnumerateArray().Select(x => x.GetString() ?? "").ToArray();

        if (extParams.Length < 3)
        {
            WriteCells(id, new[] { new[] { "mort: <principal>, <rate%>, <years>" } });
            return;
        }

        if (!double.TryParse(extParams[0], NumberStyles.Float, Inv, out double principal) || principal <= 0
            || !double.TryParse(extParams[1], NumberStyles.Float, Inv, out double annualRatePct) || annualRatePct < 0
            || !double.TryParse(extParams[2], NumberStyles.Float, Inv, out double years) || years <= 0)
        {
            WriteCells(id, new[] { new[] { "err: bad numeric input" } });
            return;
        }

        int n = (int)Math.Round(years * 12);
        double r = annualRatePct / 100.0 / 12.0;

        double monthly;
        if (r == 0) monthly = principal / n;
        else monthly = principal * r / (1.0 - Math.Pow(1.0 + r, -n));

        double total = monthly * n;
        double interest = total - principal;

        var rows = new List<string[]>
        {
            new[] { $"${principal:N0} @ {annualRatePct:0.###}% / {years:0.##}yr" },
            new[] { $"monthly: ${monthly:N2}" },
            new[] { $"total interest: ${interest:N0}" },
            new[] { $"total paid: ${total:N0}" }
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
