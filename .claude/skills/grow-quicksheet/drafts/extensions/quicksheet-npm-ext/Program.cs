using System.Text.Json;
using System.Text.Json.Serialization;

/// <summary>
/// QuickSheet npm Extension — reads JSON-lines from stdin, writes JSON-lines to stdout.
/// Registers the "npm" prefix. Given a package name, queries the public npm registry
/// and returns: latest version, description, and weekly downloads.
/// Usage: `npm: express, 1, 3` or `npm: react, 2, 4`.
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
        DefaultRequestHeaders = { { "Accept", "application/json" } }
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
            prefix = "npm",
            name = "npm Registry Lookup",
            version = "1.0.0"
        });
        SendLog("npm extension registered with prefix 'npm'");
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
            WriteCells(id, new[] { new[] { "npm: <package>" } });
            return;
        }

        string packageName = extParams[0].Trim().ToLowerInvariant();

        try
        {
            var info = FetchPackageInfo(packageName);
            var rows = new List<string[]>
            {
                new[] { $"📦 {info.Name}" },
                new[] { $"v{info.Version}" },
                new[] { info.Description }
            };
            if (gridRows > 3)
                rows.Add(new[] { $"⬇ {FormatDownloads(info.WeeklyDownloads)}/wk" });
            if (gridRows > 4)
                rows.Add(new[] { info.License ?? "—" });

            while (rows.Count < gridRows) rows.Add(new[] { "" });
            if (rows.Count > gridRows) rows = rows.Take(gridRows).ToList();
            WriteCells(id, rows);
        }
        catch (Exception ex)
        {
            WriteCells(id, new[] { new[] { $"err: {ex.Message}" } });
        }
    }

    static PackageInfo FetchPackageInfo(string name)
    {
        // Fetch abbreviated metadata from registry
        string url = $"https://registry.npmjs.org/{Uri.EscapeDataString(name)}";
        var resp = Http.GetAsync(url).GetAwaiter().GetResult();
        if (!resp.IsSuccessStatusCode)
            throw new Exception($"{(int)resp.StatusCode} {resp.ReasonPhrase}");

        string json = resp.Content.ReadAsStringAsync().GetAwaiter().GetResult();
        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        string version = "?";
        if (root.TryGetProperty("dist-tags", out var tags) &&
            tags.TryGetProperty("latest", out var latest))
        {
            version = latest.GetString() ?? "?";
        }

        string description = "";
        if (root.TryGetProperty("description", out var desc))
            description = desc.GetString() ?? "";
        if (description.Length > 60)
            description = description[..57] + "...";

        string? license = null;
        if (root.TryGetProperty("license", out var lic))
            license = lic.GetString();

        // Fetch weekly downloads from the downloads API
        long downloads = 0;
        try
        {
            string dlUrl = $"https://api.npmjs.org/downloads/point/last-week/{Uri.EscapeDataString(name)}";
            var dlResp = Http.GetAsync(dlUrl).GetAwaiter().GetResult();
            if (dlResp.IsSuccessStatusCode)
            {
                string dlJson = dlResp.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                using var dlDoc = JsonDocument.Parse(dlJson);
                if (dlDoc.RootElement.TryGetProperty("downloads", out var dlCount))
                    downloads = dlCount.GetInt64();
            }
        }
        catch { /* non-critical */ }

        return new PackageInfo
        {
            Name = name,
            Version = version,
            Description = description,
            WeeklyDownloads = downloads,
            License = license
        };
    }

    static string FormatDownloads(long count)
    {
        if (count >= 1_000_000_000) return $"{count / 1_000_000_000.0:F1}B";
        if (count >= 1_000_000) return $"{count / 1_000_000.0:F1}M";
        if (count >= 1_000) return $"{count / 1_000.0:F1}K";
        return count.ToString();
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

class PackageInfo
{
    public string Name { get; set; } = "";
    public string Version { get; set; } = "";
    public string Description { get; set; } = "";
    public long WeeklyDownloads { get; set; }
    public string? License { get; set; }
}
