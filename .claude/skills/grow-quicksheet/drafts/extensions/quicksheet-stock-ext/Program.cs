using System.Collections.Concurrent;
using System.Globalization;
using System.Text.Json;
using System.Text.Json.Serialization;

/// <summary>
/// QuickSheet Stock Extension — reads JSON-lines from stdin, writes JSON-lines to stdout.
/// Registers the "stock" prefix. Fetches a daily quote from Stooq (free, CSV).
/// Usage: `stock: AAPL, 1, 3`. US tickers default to .us; LSE etc require explicit suffix
/// (`stock: bp.uk`).
/// </summary>
class Program
{
    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        WriteIndented = false
    };

    private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromSeconds(10) };
    private static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

    // Cache responses for 5 minutes. Stooq end-of-day data changes slowly.
    private static readonly ConcurrentDictionary<string, (DateTime, Quote?)> Cache = new(StringComparer.OrdinalIgnoreCase);
    private static readonly TimeSpan CacheTtl = TimeSpan.FromMinutes(5);

    record Quote(string Symbol, string Date, double Open, double Close, double High, double Low);

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
            prefix = "stock",
            name = "Stock Quote",
            version = "1.0.0"
        });
        SendLog("Stock extension registered with prefix 'stock'");
    }

    static void HandleActivate(JsonElement root)
    {
        string id = root.TryGetProperty("id", out var idProp) ? idProp.GetString() ?? "" : "";
        int gridRows = root.TryGetProperty("gridRows", out var gr) ? gr.GetInt32() : 3;

        string[] extParams = [];
        if (root.TryGetProperty("params", out var p) && p.ValueKind == JsonValueKind.Array)
            extParams = p.EnumerateArray().Select(x => x.GetString() ?? "").ToArray();

        if (extParams.Length == 0 || string.IsNullOrWhiteSpace(extParams[0]))
        {
            WriteCells(id, new[] { new[] { "stock: <ticker>" } });
            return;
        }

        string ticker = NormalizeTicker(extParams[0]);

        try
        {
            var q = FetchQuote(ticker);
            if (q == null)
            {
                WriteCells(id, new[] { new[] { $"{ticker}: not found" } });
                return;
            }
            double change = q.Close - q.Open;
            double pct = q.Open > 0 ? change / q.Open * 100.0 : 0.0;
            string arrow = change >= 0 ? "▲" : "▼";

            var rows = new List<string[]>
            {
                new[] { $"{q.Symbol} ${q.Close.ToString("N2", Inv)}" },
                new[] { $"{arrow} {pct.ToString("0.00", Inv)}% ({change.ToString("+0.00;-0.00", Inv)})" },
                new[] { $"as of {q.Date}" }
            };
            while (rows.Count < gridRows) rows.Add(new[] { "" });
            if (rows.Count > gridRows) rows = rows.Take(gridRows).ToList();
            WriteCells(id, rows);
        }
        catch (Exception ex)
        {
            WriteCells(id, new[] { new[] { $"err: {ex.Message}" } });
        }
    }

    static string NormalizeTicker(string s)
    {
        s = s.Trim().ToLowerInvariant();
        // Default to US market if no suffix.
        if (!s.Contains('.')) s += ".us";
        return s;
    }

    static Quote? FetchQuote(string ticker)
    {
        if (Cache.TryGetValue(ticker, out var cached) && DateTime.UtcNow - cached.Item1 < CacheTtl)
            return cached.Item2;

        string url = $"https://stooq.com/q/l/?s={Uri.EscapeDataString(ticker)}&f=sd2t2ohlcv&h&e=csv";
        var resp = Http.GetAsync(url).GetAwaiter().GetResult();
        resp.EnsureSuccessStatusCode();
        string csv = resp.Content.ReadAsStringAsync().GetAwaiter().GetResult();

        Quote? q = ParseCsv(csv);
        Cache[ticker] = (DateTime.UtcNow, q);
        return q;
    }

    static Quote? ParseCsv(string csv)
    {
        // Two lines: header then data.
        var lines = csv.Split('\n', StringSplitOptions.RemoveEmptyEntries);
        if (lines.Length < 2) return null;
        var fields = lines[1].Split(',');
        // Expected: Symbol,Date,Time,Open,High,Low,Close,Volume
        if (fields.Length < 8) return null;
        if (!double.TryParse(fields[3], NumberStyles.Float, Inv, out double open)) return null;
        if (!double.TryParse(fields[4], NumberStyles.Float, Inv, out double high)) return null;
        if (!double.TryParse(fields[5], NumberStyles.Float, Inv, out double low)) return null;
        if (!double.TryParse(fields[6], NumberStyles.Float, Inv, out double close)) return null;
        return new Quote(fields[0], fields[1], open, close, high, low);
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
