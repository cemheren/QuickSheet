using System.Collections.Concurrent;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using System.Xml.Linq;

/// <summary>
/// QuickSheet arXiv Extension — looks up papers by arXiv ID.
/// Usage: `arxiv: 2301.07041` or `arxiv: 2301.07041v2`
/// Returns: title, authors (up to 4), first ~120 chars of abstract, link.
/// Free API, no auth, no NuGet deps.
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

    static Program()
    {
        Http.DefaultRequestHeaders.UserAgent.ParseAdd(
            "quicksheet-arxiv-ext/1.0 (https://github.com/cemheren/quicksheet-arxiv-ext)");
    }

    private static readonly ConcurrentDictionary<string, List<string>> Cache = new(StringComparer.OrdinalIgnoreCase);

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
            prefix = "arxiv",
            name = "arXiv Lookup",
            version = "1.0.0"
        });
        SendLog("arXiv extension registered with prefix 'arxiv'");
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
            WriteCells(id, new[] { new[] { "arxiv: <id> (e.g. 2301.07041)" } });
            return;
        }

        string arxivId = NormalizeId(extParams[0]);
        if (string.IsNullOrEmpty(arxivId))
        {
            WriteCells(id, new[] { new[] { "err: invalid arXiv ID" } });
            return;
        }

        try
        {
            var lines = FetchPaper(arxivId);
            if (lines.Count == 0)
            {
                WriteCells(id, new[] { new[] { $"{arxivId}: not found" } });
                return;
            }
            var rows = lines.Select(l => new[] { l }).ToList();
            while (rows.Count < gridRows) rows.Add(new[] { "" });
            if (rows.Count > gridRows) rows = rows.Take(gridRows).ToList();
            WriteCells(id, rows);
        }
        catch (Exception ex)
        {
            WriteCells(id, new[] { new[] { $"err: {ex.Message}" } });
        }
    }

    static string NormalizeId(string s)
    {
        s = s.Trim();
        // Strip common URL prefixes
        if (s.StartsWith("https://arxiv.org/abs/", StringComparison.OrdinalIgnoreCase))
            s = s[22..];
        else if (s.StartsWith("http://arxiv.org/abs/", StringComparison.OrdinalIgnoreCase))
            s = s[21..];
        else if (s.StartsWith("arxiv:", StringComparison.OrdinalIgnoreCase))
            s = s[6..];
        s = s.Trim();
        // Validate: new-style (YYMM.NNNNN[vN]) or old-style (category/YYMMNNN)
        if (Regex.IsMatch(s, @"^\d{4}\.\d{4,5}(v\d+)?$")) return s;
        if (Regex.IsMatch(s, @"^[a-z\-]+/\d{7}(v\d+)?$")) return s;
        return "";
    }

    static List<string> FetchPaper(string arxivId)
    {
        if (Cache.TryGetValue(arxivId, out var cached)) return cached;

        string url = $"https://export.arxiv.org/api/query?id_list={Uri.EscapeDataString(arxivId)}";
        var resp = Http.GetAsync(url).GetAwaiter().GetResult();
        if (!resp.IsSuccessStatusCode) { Cache[arxivId] = []; return []; }
        string xml = resp.Content.ReadAsStringAsync().GetAwaiter().GetResult();

        var lines = new List<string>();
        XNamespace atom = "http://www.w3.org/2005/Atom";
        var feed = XDocument.Parse(xml);
        var entry = feed.Root?.Element(atom + "entry");
        if (entry == null) { Cache[arxivId] = lines; return lines; }

        // Check for error (arXiv returns an entry with id containing "api/errors")
        string? entryId = entry.Element(atom + "id")?.Value;
        if (entryId != null && entryId.Contains("api/errors"))
        {
            Cache[arxivId] = lines;
            return lines;
        }

        // Title
        string title = entry.Element(atom + "title")?.Value?.Trim() ?? "";
        title = Regex.Replace(title, @"\s+", " ");
        if (!string.IsNullOrEmpty(title))
            lines.Add(title.Length > 100 ? title[..97] + "..." : title);

        // Authors (up to 4)
        var authors = entry.Elements(atom + "author")
            .Select(a => a.Element(atom + "name")?.Value ?? "")
            .Where(n => !string.IsNullOrEmpty(n))
            .ToList();
        if (authors.Count > 0)
        {
            string authStr = authors.Count <= 4
                ? string.Join(", ", authors)
                : string.Join(", ", authors.Take(4)) + ", et al.";
            lines.Add(authStr);
        }

        // Abstract snippet
        string abs = entry.Element(atom + "summary")?.Value?.Trim() ?? "";
        abs = Regex.Replace(abs, @"\s+", " ");
        if (!string.IsNullOrEmpty(abs))
            lines.Add(abs.Length > 120 ? abs[..117] + "..." : abs);

        // Link
        lines.Add($"https://arxiv.org/abs/{arxivId}");

        Cache[arxivId] = lines;
        return lines;
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
