using System.Diagnostics;
using System.Text.Json;
using System.Text.Json.Serialization;

/// <summary>
/// QuickSheet Ping Extension — reads JSON-lines from stdin, writes JSON-lines to stdout.
/// Registers the "ping" prefix. Given a URL (or host), does an HTTP HEAD and reports
/// status code + latency in ms. No long-lived cache — this is intentionally a probe.
/// Usage: `ping: https://example.com, 1, 3` or `ping: example.com, 1, 3`.
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
        // Don't follow redirects — we want the real first-hop status.
        AllowAutoRedirect = false
    })
    { Timeout = TimeSpan.FromSeconds(8) };

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
            prefix = "ping",
            name = "HTTP Ping",
            version = "1.0.0"
        });
        SendLog("Ping extension registered with prefix 'ping'");
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
            WriteCells(id, new[] { new[] { "ping: <url>" } });
            return;
        }

        string target = extParams[0].Trim();
        if (!target.StartsWith("http://", StringComparison.OrdinalIgnoreCase)
            && !target.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
        {
            target = "https://" + target;
        }

        try
        {
            var (code, ms, reason) = Probe(target);
            string statusLine = reason == null
                ? $"{code}"
                : $"{code} {reason}";
            string indicator = code >= 200 && code < 400 ? "✓"
                : code >= 400 && code < 500 ? "⚠"
                : code >= 500 ? "✗"
                : "?";

            var rows = new List<string[]>
            {
                new[] { $"{indicator} {target}" },
                new[] { statusLine },
                new[] { $"{ms} ms" }
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

    static (int code, long ms, string? reason) Probe(string url)
    {
        var sw = Stopwatch.StartNew();
        using var req = new HttpRequestMessage(HttpMethod.Head, url);
        HttpResponseMessage resp;
        try
        {
            resp = Http.SendAsync(req).GetAwaiter().GetResult();
        }
        catch
        {
            // Some servers reject HEAD. Retry with GET.
            using var req2 = new HttpRequestMessage(HttpMethod.Get, url);
            resp = Http.SendAsync(req2, HttpCompletionOption.ResponseHeadersRead).GetAwaiter().GetResult();
        }
        sw.Stop();
        using (resp)
        {
            return ((int)resp.StatusCode, sw.ElapsedMilliseconds, resp.ReasonPhrase);
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
