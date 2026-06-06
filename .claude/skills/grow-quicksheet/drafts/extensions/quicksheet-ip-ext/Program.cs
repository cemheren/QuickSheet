using System.Text.Json;
using System.Text.Json.Serialization;

/// <summary>
/// QuickSheet IP Extension — reads JSON-lines from stdin, writes JSON-lines to stdout.
/// Registers the "ip" prefix. Given an IP address (or blank for your own public IP),
/// returns geolocation data: country, city, ISP, timezone.
/// Uses ip-api.com (free, no API key, 45 req/min limit).
/// Usage: `ip: 8.8.8.8, 1, 4` or `ip: , 1, 4` (own IP).
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
        Timeout = TimeSpan.FromSeconds(8)
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
            prefix = "ip",
            name = "IP Geolocation",
            version = "1.0.0"
        });
        SendLog("IP extension registered with prefix 'ip'");
    }

    static void HandleActivate(JsonElement root)
    {
        string id = root.TryGetProperty("id", out var idProp) ? idProp.GetString() ?? "" : "";
        int gridRows = root.TryGetProperty("gridRows", out var gr) ? gr.GetInt32() : 4;

        string[] extParams = [];
        if (root.TryGetProperty("params", out var p) && p.ValueKind == JsonValueKind.Array)
            extParams = p.EnumerateArray().Select(x => x.GetString() ?? "").ToArray();

        string target = extParams.Length > 0 ? extParams[0].Trim() : "";

        try
        {
            var geo = Lookup(target);
            var rows = new List<string[]>();

            if (geo.Status == "fail")
            {
                rows.Add(new[] { $"✗ {geo.Message ?? "lookup failed"}" });
            }
            else
            {
                string ip = string.IsNullOrEmpty(target) ? geo.Query ?? "?" : target;
                rows.Add(new[] { $"🌐 {ip}" });
                rows.Add(new[] { $"{geo.City}, {geo.RegionName}" });
                rows.Add(new[] { $"{geo.Country} ({geo.CountryCode})" });
                rows.Add(new[] { $"{geo.Isp}" });
                if (gridRows > 4)
                    rows.Add(new[] { $"TZ: {geo.Timezone}" });
                if (gridRows > 5)
                    rows.Add(new[] { $"AS: {geo.As}" });
            }

            while (rows.Count < gridRows) rows.Add(new[] { "" });
            if (rows.Count > gridRows) rows = rows.Take(gridRows).ToList();
            WriteCells(id, rows);
        }
        catch (Exception ex)
        {
            WriteCells(id, new[] { new[] { $"err: {ex.Message}" } });
        }
    }

    static GeoResult Lookup(string ip)
    {
        string url = string.IsNullOrEmpty(ip)
            ? "http://ip-api.com/json/?fields=status,message,country,countryCode,regionName,city,timezone,isp,as,query"
            : $"http://ip-api.com/json/{ip}?fields=status,message,country,countryCode,regionName,city,timezone,isp,as,query";

        string json = Http.GetStringAsync(url).GetAwaiter().GetResult();
        return JsonSerializer.Deserialize<GeoResult>(json, new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true
        }) ?? new GeoResult { Status = "fail", Message = "empty response" };
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

class GeoResult
{
    public string? Status { get; set; }
    public string? Message { get; set; }
    public string? Country { get; set; }
    public string? CountryCode { get; set; }
    public string? RegionName { get; set; }
    public string? City { get; set; }
    public string? Timezone { get; set; }
    public string? Isp { get; set; }
    public string? As { get; set; }
    public string? Query { get; set; }
}
