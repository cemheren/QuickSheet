using System.Text.Json;
using System.Text.Json.Serialization;

/// <summary>
/// QuickSheet FX Extension — live currency exchange rates.
/// Uses frankfurter.app (ECB data, free, no auth).
/// Usage: `fx: USD, EUR, 1, 3` → converts 1 USD to EUR (3 rows: pair, rate, updated).
///        `fx: GBP, JPY,CHF,CAD, 1, 4` → multi-target rates.
/// </summary>
class Program
{
    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        WriteIndented = false
    };

    private static readonly HttpClient Http = new()
    {
        Timeout = TimeSpan.FromSeconds(10),
        BaseAddress = new Uri("https://api.frankfurter.app/")
    };

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
            prefix = "fx",
            name = "Currency Exchange (ECB)",
            version = "1.0.0"
        });
        SendLog("FX extension registered with prefix 'fx'");
    }

    static void HandleActivate(JsonElement root)
    {
        string id = root.TryGetProperty("id", out var idProp) ? idProp.GetString() ?? "" : "";
        int gridRows = root.TryGetProperty("gridRows", out var gr) ? gr.GetInt32() : 3;

        string[] extParams = [];
        if (root.TryGetProperty("params", out var p) && p.ValueKind == JsonValueKind.Array)
            extParams = p.EnumerateArray().Select(x => x.GetString() ?? "").ToArray();

        if (extParams.Length < 2)
        {
            WriteCells(id, new[] { new[] { "fx: <from>, <to>[,to2,...], cols, rows" } });
            return;
        }

        string baseCurrency = extParams[0].Trim().ToUpperInvariant();

        // Parse targets — everything between the first param and the trailing numeric params
        var targets = new List<string>();
        int numericTail = extParams.Length;
        for (int i = extParams.Length - 1; i >= 1; i--)
        {
            if (int.TryParse(extParams[i].Trim(), out _))
                numericTail = i;
            else
                break;
        }
        for (int i = 1; i < numericTail; i++)
        {
            string t = extParams[i].Trim().ToUpperInvariant();
            if (!string.IsNullOrEmpty(t)) targets.Add(t);
        }

        if (targets.Count == 0)
        {
            WriteCells(id, new[] { new[] { "fx: need target currency" } });
            return;
        }

        try
        {
            string symbols = string.Join(",", targets);
            string url = $"latest?from={baseCurrency}&to={symbols}";
            string json = Http.GetStringAsync(url).GetAwaiter().GetResult();

            using var resp = JsonDocument.Parse(json);
            var rates = resp.RootElement.GetProperty("rates");
            string date = resp.RootElement.TryGetProperty("date", out var d) ? d.GetString() ?? "" : "";

            var rows = new List<string[]>();

            // Header row
            rows.Add(new[] { $"1 {baseCurrency} →" });

            // One row per target currency
            foreach (string tgt in targets)
            {
                if (rates.TryGetProperty(tgt, out var rateVal))
                {
                    decimal rate = rateVal.GetDecimal();
                    rows.Add(new[] { $"{tgt} {rate:G6}" });
                }
                else
                {
                    rows.Add(new[] { $"{tgt} n/a" });
                }
            }

            // Date row
            rows.Add(new[] { $"ECB {date}" });

            // Fit to grid
            while (rows.Count < gridRows) rows.Add(new[] { "" });
            if (rows.Count > gridRows) rows = rows.Take(gridRows).ToList();
            WriteCells(id, rows);
        }
        catch (Exception ex)
        {
            string msg = ex.Message.Length > 60 ? ex.Message[..60] + "…" : ex.Message;
            WriteCells(id, new[] { new[] { $"err: {msg}" } });
        }
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
