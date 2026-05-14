using System.Collections.Concurrent;
using System.Text.Json;
using System.Text.Json.Serialization;

/// <summary>
/// QuickSheet Price Extension — reads JSON-lines from stdin, writes JSON-lines to stdout.
/// Registers the "price" prefix. Given a CoinGecko coin id (or a common ticker),
/// returns last price in USD plus 24h change.
/// Usage: `price: bitcoin, 1, 2` → fills 2 rows: id+price, 24h change %.
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

    // Small built-in ticker → CoinGecko id alias map. Anything else passes through as-is.
    private static readonly Dictionary<string, string> TickerAliases = new(StringComparer.OrdinalIgnoreCase)
    {
        ["btc"] = "bitcoin",
        ["eth"] = "ethereum",
        ["sol"] = "solana",
        ["doge"] = "dogecoin",
        ["ada"] = "cardano",
        ["xrp"] = "ripple",
        ["dot"] = "polkadot",
        ["matic"] = "matic-network",
        ["link"] = "chainlink",
        ["ltc"] = "litecoin",
        ["bch"] = "bitcoin-cash",
        ["avax"] = "avalanche-2",
        ["atom"] = "cosmos",
        ["arb"] = "arbitrum",
        ["op"] = "optimism"
    };

    // Cache responses for 60 seconds to avoid rate-limit and reduce flicker.
    private static readonly ConcurrentDictionary<string, (DateTime, JsonDocument)> Cache = new(StringComparer.OrdinalIgnoreCase);
    private static readonly TimeSpan CacheTtl = TimeSpan.FromSeconds(60);

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
            prefix = "price",
            name = "Crypto Price Quote",
            version = "1.0.0"
        });
        SendLog("Price extension registered with prefix 'price'");
    }

    static void HandleActivate(JsonElement root)
    {
        string id = root.TryGetProperty("id", out var idProp) ? idProp.GetString() ?? "" : "";
        int gridRows = root.TryGetProperty("gridRows", out var gr) ? gr.GetInt32() : 2;

        string[] extParams = [];
        if (root.TryGetProperty("params", out var p) && p.ValueKind == JsonValueKind.Array)
            extParams = p.EnumerateArray().Select(x => x.GetString() ?? "").ToArray();

        if (extParams.Length == 0)
        {
            WriteCells(id, new[] { new[] { "price: <coin>" } });
            return;
        }

        string raw = extParams[0].Trim();
        string coinId = TickerAliases.TryGetValue(raw, out var mapped) ? mapped : raw.ToLowerInvariant();

        try
        {
            var data = FetchPrice(coinId);
            if (data == null)
            {
                WriteCells(id, new[] { new[] { $"{coinId}: not found" } });
                return;
            }
            double price = data.Value.price;
            double change = data.Value.change24h;
            string arrow = change >= 0 ? "▲" : "▼";
            string changeStr = $"{arrow} {Math.Abs(change):0.00}% 24h";

            var rows = new List<string[]>
            {
                new[] { $"{coinId} ${FormatPrice(price)}" },
                new[] { changeStr }
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

    static (double price, double change24h)? FetchPrice(string coinId)
    {
        if (Cache.TryGetValue(coinId, out var cached) && DateTime.UtcNow - cached.Item1 < CacheTtl)
        {
            return ParseResponse(cached.Item2, coinId);
        }

        string url = $"https://api.coingecko.com/api/v3/simple/price?ids={Uri.EscapeDataString(coinId)}&vs_currencies=usd&include_24hr_change=true";
        string json = Http.GetStringAsync(url).GetAwaiter().GetResult();
        var doc = JsonDocument.Parse(json);
        Cache[coinId] = (DateTime.UtcNow, doc);
        return ParseResponse(doc, coinId);
    }

    static (double price, double change24h)? ParseResponse(JsonDocument doc, string coinId)
    {
        if (!doc.RootElement.TryGetProperty(coinId, out var node)) return null;
        if (!node.TryGetProperty("usd", out var priceProp)) return null;
        double price = priceProp.GetDouble();
        double change = node.TryGetProperty("usd_24h_change", out var ch) ? ch.GetDouble() : 0.0;
        return (price, change);
    }

    static string FormatPrice(double p)
    {
        if (p >= 1000) return p.ToString("N0");
        if (p >= 1) return p.ToString("N2");
        if (p >= 0.01) return p.ToString("N4");
        return p.ToString("N6");
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
