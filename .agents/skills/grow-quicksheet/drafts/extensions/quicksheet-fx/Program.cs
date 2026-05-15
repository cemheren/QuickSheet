using System.Net.Http;
using System.Text.Json;
using System.Text.Json.Serialization;

/// <summary>
/// QuickSheet Currency Conversion Extension — live exchange rates from ECB via Frankfurter API.
/// Prefix: "fx". Usage: "fx: 1000, USD, EUR" or "fx: USD, EUR" (defaults to 1 unit).
/// No API key required. Rates update daily from the European Central Bank.
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
        Timeout = TimeSpan.FromSeconds(10)
    };

    // Cache: base+quote -> (rate, date, fetchedAt)
    private static readonly Dictionary<string, (double rate, string date, DateTime fetchedAt)> RateCache = new();
    private static readonly TimeSpan CacheTtl = TimeSpan.FromHours(1);

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
                    case "init":
                        HandleInit();
                        break;
                    case "activate":
                        HandleActivate(doc.RootElement);
                        break;
                    case "deactivate":
                        // No ongoing timers to stop — fx is request/response
                        break;
                }
            }
            catch (Exception ex)
            {
                SendJson(new { type = "error", id = "", message = $"Parse error: {ex.Message}" });
            }
        }
    }

    static void HandleInit()
    {
        SendJson(new
        {
            type = "register",
            prefix = "fx",
            name = "Currency Converter",
            version = "1.0.0"
        });
        SendLog("Currency Converter registered with prefix 'fx'. Rates from ECB via Frankfurter API.");
    }

    static void HandleActivate(JsonElement root)
    {
        string id = root.TryGetProperty("id", out var idProp) ? idProp.GetString() ?? "" : "";

        string[] extParams = [];
        if (root.TryGetProperty("params", out var paramsProp) && paramsProp.ValueKind == JsonValueKind.Array)
        {
            extParams = paramsProp.EnumerateArray()
                .Select(p => p.GetString()?.Trim() ?? "")
                .Where(p => p.Length > 0)
                .ToArray();
        }

        // Parse: "fx: 1000, USD, EUR" or "fx: USD, EUR" or "fx: USD, EUR, GBP, CAD"
        double amount = 1;
        string baseCurrency;
        string[] quoteCurrencies;

        if (extParams.Length < 2)
        {
            SendCells(id, [
                (0, 0, "⚠️ Usage: fx: [amount,] FROM, TO [, TO2, ...]"),
                (1, 0, "Example: fx: 1000, USD, EUR")
            ]);
            return;
        }

        // If first param is numeric, it's the amount
        if (double.TryParse(extParams[0], System.Globalization.NumberStyles.Any,
            System.Globalization.CultureInfo.InvariantCulture, out double parsedAmount))
        {
            amount = parsedAmount;
            baseCurrency = extParams[1].ToUpperInvariant();
            quoteCurrencies = extParams.Skip(2).Select(p => p.ToUpperInvariant()).ToArray();
        }
        else
        {
            baseCurrency = extParams[0].ToUpperInvariant();
            quoteCurrencies = extParams.Skip(1).Select(p => p.ToUpperInvariant()).ToArray();
        }

        if (quoteCurrencies.Length == 0)
        {
            SendCells(id, [(0, 0, "⚠️ Need at least one target currency"), (1, 0, "Example: fx: 1000, USD, EUR")]);
            return;
        }

        // Fetch rates and build output
        var cells = new List<(int r, int c, string v)>();
        int row = 0;

        // Header
        string amountStr = amount == 1 ? "" : $"{amount:N2} ";
        cells.Add((row, 0, $"💱 {amountStr}{baseCurrency}"));
        cells.Add((row, 1, "Rate"));
        cells.Add((row, 2, "Converted"));
        row++;

        foreach (var quote in quoteCurrencies)
        {
            var result = FetchRate(baseCurrency, quote);
            if (result.HasValue)
            {
                var (rate, date) = result.Value;
                double converted = amount * rate;
                cells.Add((row, 0, $"→ {quote}"));
                cells.Add((row, 1, $"{rate:F4}"));
                cells.Add((row, 2, $"{converted:N2} {quote}"));
            }
            else
            {
                cells.Add((row, 0, $"→ {quote}"));
                cells.Add((row, 1, "error"));
                cells.Add((row, 2, "fetch failed"));
            }
            row++;
        }

        // Footer with source and date
        var anyResult = FetchRate(baseCurrency, quoteCurrencies[0]);
        string dateStr = anyResult?.date ?? "unknown";
        cells.Add((row, 0, $"ECB · {dateStr}"));
        row++;

        SendCells(id, cells);
    }

    static (double rate, string date)? FetchRate(string baseCur, string quoteCur)
    {
        string cacheKey = $"{baseCur}/{quoteCur}";

        // Check cache
        if (RateCache.TryGetValue(cacheKey, out var cached) &&
            DateTime.UtcNow - cached.fetchedAt < CacheTtl)
        {
            return (cached.rate, cached.date);
        }

        // Fetch from Frankfurter
        try
        {
            string url = $"https://api.frankfurter.dev/v2/rate/{baseCur}/{quoteCur}";
            var response = Http.GetStringAsync(url).GetAwaiter().GetResult();
            using var doc = JsonDocument.Parse(response);

            double rate = doc.RootElement.GetProperty("rate").GetDouble();
            string date = doc.RootElement.GetProperty("date").GetString() ?? "unknown";

            RateCache[cacheKey] = (rate, date, DateTime.UtcNow);
            return (rate, date);
        }
        catch (Exception ex)
        {
            SendLog($"Rate fetch failed for {cacheKey}: {ex.Message}");
            return null;
        }
    }

    static void SendCells(string id, List<(int r, int c, string v)> cells)
    {
        var cellObjects = cells.Select(c => new { r = c.r, c = c.c, v = c.v }).ToArray();
        SendJson(new { type = "write", id, cells = cellObjects });
    }

    static void SendJson(object obj)
    {
        string json = JsonSerializer.Serialize(obj, JsonOpts);
        Console.WriteLine(json);
        Console.Out.Flush();
    }

    static void SendLog(string message)
    {
        SendJson(new { type = "log", message });
    }
}
