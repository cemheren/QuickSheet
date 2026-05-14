using System.Collections.Concurrent;
using System.Text.Json;
using System.Text.Json.Serialization;

/// <summary>
/// QuickSheet MX Check Extension — reads JSON-lines from stdin, writes JSON-lines to stdout.
/// Registers the "mxck" prefix. Queries Google's DNS-over-HTTPS resolver for MX records.
/// Usage: `mxck: example.com, 1, 5` → fills rows with MX hosts sorted by priority.
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

    // Cache MX answers for 1 hour. MX records change rarely.
    private static readonly ConcurrentDictionary<string, (DateTime, List<(int prio, string host)>)> Cache = new(StringComparer.OrdinalIgnoreCase);
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
            prefix = "mxck",
            name = "MX Record Check",
            version = "1.0.0"
        });
        SendLog("MX extension registered with prefix 'mxck'");
    }

    static void HandleActivate(JsonElement root)
    {
        string id = root.TryGetProperty("id", out var idProp) ? idProp.GetString() ?? "" : "";
        int gridRows = root.TryGetProperty("gridRows", out var gr) ? gr.GetInt32() : 5;

        string[] extParams = [];
        if (root.TryGetProperty("params", out var p) && p.ValueKind == JsonValueKind.Array)
            extParams = p.EnumerateArray().Select(x => x.GetString() ?? "").ToArray();

        if (extParams.Length == 0 || string.IsNullOrWhiteSpace(extParams[0]))
        {
            WriteCells(id, new[] { new[] { "mxck: <domain>" } });
            return;
        }

        string domain = extParams[0].Trim().TrimEnd('.');

        try
        {
            var mx = FetchMx(domain);
            var rows = new List<string[]> { new[] { $"MX for {domain}" } };
            if (mx.Count == 0)
                rows.Add(new[] { "(no MX records)" });
            else
                foreach (var (prio, host) in mx)
                    rows.Add(new[] { $"{prio,3} {host}" });

            while (rows.Count < gridRows) rows.Add(new[] { "" });
            if (rows.Count > gridRows) rows = rows.Take(gridRows).ToList();
            WriteCells(id, rows);
        }
        catch (Exception ex)
        {
            WriteCells(id, new[] { new[] { $"err: {ex.Message}" } });
        }
    }

    static List<(int prio, string host)> FetchMx(string domain)
    {
        if (Cache.TryGetValue(domain, out var cached) && DateTime.UtcNow - cached.Item1 < CacheTtl)
            return cached.Item2;

        string url = $"https://dns.google/resolve?name={Uri.EscapeDataString(domain)}&type=MX";
        var resp = Http.GetAsync(url).GetAwaiter().GetResult();
        resp.EnsureSuccessStatusCode();
        string json = resp.Content.ReadAsStringAsync().GetAwaiter().GetResult();

        var results = new List<(int, string)>();
        using var doc = JsonDocument.Parse(json);
        if (doc.RootElement.TryGetProperty("Answer", out var answer) && answer.ValueKind == JsonValueKind.Array)
        {
            foreach (var rec in answer.EnumerateArray())
            {
                if (rec.TryGetProperty("type", out var t) && t.GetInt32() != 15) continue; // 15 = MX
                if (!rec.TryGetProperty("data", out var d)) continue;
                string data = d.GetString() ?? "";
                // Format: "<priority> <host>"
                int space = data.IndexOf(' ');
                if (space < 0) continue;
                if (!int.TryParse(data[..space], out int prio)) continue;
                string host = data[(space + 1)..].TrimEnd('.');
                results.Add((prio, host));
            }
        }
        results.Sort((a, b) => a.Item1.CompareTo(b.Item1));
        Cache[domain] = (DateTime.UtcNow, results);
        return results;
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
