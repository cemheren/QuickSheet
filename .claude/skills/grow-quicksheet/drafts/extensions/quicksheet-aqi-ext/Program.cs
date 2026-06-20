using System.Globalization;
using System.Text.Json;
using System.Text.Json.Serialization;

/// <summary>
/// QuickSheet AQI Extension — reads JSON-lines from stdin, writes JSON-lines to stdout.
/// Registers the "aqi" prefix. Given a city name (or "lat,lon"), geocodes it with the
/// keyless Open-Meteo geocoding API, then fetches current air quality (US AQI, PM2.5,
/// PM10) from the keyless Open-Meteo Air Quality API.
/// Usage: `aqi: London` · `aqi: New York` · `aqi: 37.77,-122.42`.
/// </summary>
class Program
{
    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        WriteIndented = false
    };

    private static readonly HttpClient Http = new(new SocketsHttpHandler
    {
        AutomaticDecompression = System.Net.DecompressionMethods.All
    })
    { Timeout = TimeSpan.FromSeconds(10) };

    static void Main()
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;
        Http.DefaultRequestHeaders.Add("User-Agent", "quicksheet-aqi-ext/1.0");
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
            prefix = "aqi",
            name = "Air Quality",
            version = "1.0.0"
        });
        SendLog("AQI extension registered with prefix 'aqi'");
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
            WriteCells(id, new[] { new[] { "aqi: <city>" } });
            return;
        }

        string query = extParams[0].Trim();

        try
        {
            var (lat, lon, label) = Resolve(query);
            var rows = Fetch(lat, lon, label);
            while (rows.Count < gridRows) rows.Add(new[] { "" });
            if (rows.Count > gridRows) rows = rows.Take(gridRows).ToList();
            WriteCells(id, rows);
        }
        catch (Exception ex)
        {
            WriteCells(id, new[] { new[] { $"err: {ex.Message}" } });
        }
    }

    static (double lat, double lon, string label) Resolve(string query)
    {
        // Accept "lat,lon" directly.
        var parts = query.Split(',');
        if (parts.Length == 2
            && double.TryParse(parts[0].Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out double dlat)
            && double.TryParse(parts[1].Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out double dlon))
        {
            return (dlat, dlon, $"{dlat:0.##},{dlon:0.##}");
        }

        string url = "https://geocoding-api.open-meteo.com/v1/search?count=1&language=en&format=json&name="
                     + Uri.EscapeDataString(query);
        string body = Http.GetStringAsync(url).GetAwaiter().GetResult();
        using var doc = JsonDocument.Parse(body);
        if (!doc.RootElement.TryGetProperty("results", out var results)
            || results.ValueKind != JsonValueKind.Array
            || results.GetArrayLength() == 0)
        {
            throw new Exception($"no match for '{query}'");
        }

        var r = results[0];
        double lat = r.GetProperty("latitude").GetDouble();
        double lon = r.GetProperty("longitude").GetDouble();
        string name = r.TryGetProperty("name", out var n) ? n.GetString() ?? query : query;
        string country = r.TryGetProperty("country_code", out var c) ? c.GetString() ?? "" : "";
        string label = string.IsNullOrEmpty(country) ? name : $"{name}, {country}";
        return (lat, lon, label);
    }

    static List<string[]> Fetch(double lat, double lon, string label)
    {
        string url = "https://air-quality-api.open-meteo.com/v1/air-quality"
                   + $"?latitude={lat.ToString(CultureInfo.InvariantCulture)}"
                   + $"&longitude={lon.ToString(CultureInfo.InvariantCulture)}"
                   + "&current=us_aqi,pm2_5,pm10";
        string body = Http.GetStringAsync(url).GetAwaiter().GetResult();
        using var doc = JsonDocument.Parse(body);

        if (!doc.RootElement.TryGetProperty("current", out var cur))
            throw new Exception("no data");

        int? usAqi = TryInt(cur, "us_aqi");
        double? pm25 = TryDouble(cur, "pm2_5");
        double? pm10 = TryDouble(cur, "pm10");

        var (dot, category) = Categorize(usAqi);

        var rows = new List<string[]>
        {
            new[] { $"{dot} {label}" },
            new[] { usAqi.HasValue ? $"US AQI {usAqi} ({category})" : $"US AQI n/a ({category})" },
            new[] { pm25.HasValue ? $"PM2.5 {pm25:0.#} ug/m3" : "PM2.5 n/a" },
            new[] { pm10.HasValue ? $"PM10 {pm10:0.#} ug/m3" : "PM10 n/a" }
        };
        return rows;
    }

    static (string dot, string category) Categorize(int? aqi)
    {
        if (!aqi.HasValue) return ("?", "unknown");
        int v = aqi.Value;
        // US EPA AQI bands.
        if (v <= 50) return ("\u25CF", "Good");
        if (v <= 100) return ("\u25CF", "Moderate");
        if (v <= 150) return ("\u25CF", "Unhealthy (sensitive)");
        if (v <= 200) return ("\u25CF", "Unhealthy");
        if (v <= 300) return ("\u25CF", "Very Unhealthy");
        return ("\u25CF", "Hazardous");
    }

    static int? TryInt(JsonElement obj, string name)
        => obj.TryGetProperty(name, out var e) && e.ValueKind == JsonValueKind.Number
            ? (int)Math.Round(e.GetDouble()) : null;

    static double? TryDouble(JsonElement obj, string name)
        => obj.TryGetProperty(name, out var e) && e.ValueKind == JsonValueKind.Number
            ? e.GetDouble() : null;

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
