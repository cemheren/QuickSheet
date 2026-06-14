using System.Net;
using System.Text.Json;
using System.Text.Json.Serialization;

/// <summary>
/// QuickSheet Wiki Extension — Wikipedia article summaries.
/// Uses the Wikimedia REST API (free, no auth, no rate-limit key).
/// Usage: `wiki: Pythagorean theorem, 1, 5` → title + extract in 5 rows.
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
        AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
    })
    {
        Timeout = TimeSpan.FromSeconds(10)
    };

    static Program()
    {
        Http.DefaultRequestHeaders.UserAgent.ParseAdd("QuickSheet-WikiExt/1.0 (https://github.com/cemheren/QuickSheet)");
        Http.DefaultRequestHeaders.Accept.ParseAdd("application/json");
    }

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
            prefix = "wiki",
            name = "Wikipedia Summary",
            version = "1.0.0"
        });
        SendLog("Wiki extension registered with prefix 'wiki'");
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
            WriteCells(id, new[] { new[] { "wiki: <topic>, cols, rows" } });
            return;
        }

        string topic = extParams[0].Trim();

        try
        {
            // Use Wikipedia REST API summary endpoint
            string encoded = Uri.EscapeDataString(topic.Replace(' ', '_'));
            string url = $"https://en.wikipedia.org/api/rest_v1/page/summary/{encoded}";
            string json = Http.GetStringAsync(url).GetAwaiter().GetResult();

            using var resp = JsonDocument.Parse(json);
            string title = resp.RootElement.TryGetProperty("title", out var t)
                ? t.GetString() ?? topic : topic;
            string extract = resp.RootElement.TryGetProperty("extract", out var e)
                ? e.GetString() ?? "" : "";

            if (string.IsNullOrWhiteSpace(extract))
            {
                WriteCells(id, new[] { new[] { $"wiki: no article for '{topic}'" } });
                return;
            }

            var rows = new List<string[]>();
            rows.Add(new[] { $"📖 {title}" });

            // Word-wrap extract into available rows
            int availRows = gridRows - 1; // reserve first for title
            int charsPerRow = 50; // approximate fit for typical cell width
            var words = extract.Split(' ');
            var currentLine = "";

            foreach (var word in words)
            {
                if (rows.Count > availRows) break;
                if (currentLine.Length + word.Length + 1 > charsPerRow)
                {
                    rows.Add(new[] { currentLine.TrimEnd() });
                    currentLine = word + " ";
                }
                else
                {
                    currentLine += word + " ";
                }
            }
            if (!string.IsNullOrWhiteSpace(currentLine) && rows.Count <= availRows)
                rows.Add(new[] { currentLine.TrimEnd() });

            // Fit to grid
            while (rows.Count < gridRows) rows.Add(new[] { "" });
            if (rows.Count > gridRows) rows = rows.Take(gridRows).ToList();
            WriteCells(id, rows);
        }
        catch (HttpRequestException ex) when (ex.StatusCode == HttpStatusCode.NotFound)
        {
            WriteCells(id, new[] { new[] { $"wiki: '{topic}' not found" } });
        }
        catch (Exception ex)
        {
            string msg = ex.Message.Length > 50 ? ex.Message[..50] + "…" : ex.Message;
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
