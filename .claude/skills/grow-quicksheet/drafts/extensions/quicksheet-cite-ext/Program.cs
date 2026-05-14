using System.Collections.Concurrent;
using System.Text.Json;
using System.Text.Json.Serialization;

/// <summary>
/// QuickSheet Cite Extension — reads JSON-lines from stdin, writes JSON-lines to stdout.
/// Registers the "cite" prefix. Looks up a DOI on the Crossref API and formats a short
/// citation: authors, year, title, container/journal.
/// Usage: `cite: 10.1145/3623476.3623525, 1, 4`.
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
        // Crossref asks for a User-Agent identifying the consumer.
        Http.DefaultRequestHeaders.UserAgent.ParseAdd("quicksheet-cite-ext/1.0 (https://github.com/cemheren/quicksheet-cite-ext)");
    }

    // Cache forever-ish — DOIs don't change.
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
            prefix = "cite",
            name = "DOI Citation",
            version = "1.0.0"
        });
        SendLog("Cite extension registered with prefix 'cite'");
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
            WriteCells(id, new[] { new[] { "cite: <doi>" } });
            return;
        }

        string doi = NormalizeDoi(extParams[0]);
        try
        {
            var lines = FetchCitation(doi);
            if (lines.Count == 0)
            {
                WriteCells(id, new[] { new[] { $"{doi}: not found" } });
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

    static string NormalizeDoi(string s)
    {
        s = s.Trim();
        if (s.StartsWith("https://doi.org/", StringComparison.OrdinalIgnoreCase)) s = s[16..];
        else if (s.StartsWith("http://doi.org/", StringComparison.OrdinalIgnoreCase)) s = s[15..];
        else if (s.StartsWith("doi:", StringComparison.OrdinalIgnoreCase)) s = s[4..];
        return s.Trim();
    }

    static List<string> FetchCitation(string doi)
    {
        if (Cache.TryGetValue(doi, out var cached)) return cached;

        string url = $"https://api.crossref.org/works/{Uri.EscapeDataString(doi)}";
        var resp = Http.GetAsync(url).GetAwaiter().GetResult();
        if (!resp.IsSuccessStatusCode) { Cache[doi] = []; return []; }
        string json = resp.Content.ReadAsStringAsync().GetAwaiter().GetResult();

        var lines = new List<string>();
        using var doc = JsonDocument.Parse(json);
        if (!doc.RootElement.TryGetProperty("message", out var msg)) return lines;

        // Authors
        string authors = "";
        if (msg.TryGetProperty("author", out var authProp) && authProp.ValueKind == JsonValueKind.Array)
        {
            var names = new List<string>();
            foreach (var a in authProp.EnumerateArray())
            {
                string fam = a.TryGetProperty("family", out var fp) ? fp.GetString() ?? "" : "";
                string giv = a.TryGetProperty("given", out var gp) ? gp.GetString() ?? "" : "";
                string init = string.IsNullOrEmpty(giv) ? "" : $" {giv[0]}.";
                if (!string.IsNullOrEmpty(fam)) names.Add($"{fam}{init}");
                if (names.Count >= 3) break;
            }
            if (names.Count > 0)
            {
                authors = string.Join(", ", names);
                if (authProp.GetArrayLength() > names.Count) authors += ", et al.";
            }
        }

        // Year
        int year = 0;
        if (msg.TryGetProperty("issued", out var issued)
            && issued.TryGetProperty("date-parts", out var dp)
            && dp.ValueKind == JsonValueKind.Array
            && dp.GetArrayLength() > 0)
        {
            var first = dp.EnumerateArray().First();
            if (first.ValueKind == JsonValueKind.Array && first.GetArrayLength() > 0)
                year = first.EnumerateArray().First().GetInt32();
        }

        // Title (array, take first)
        string title = "";
        if (msg.TryGetProperty("title", out var tp) && tp.ValueKind == JsonValueKind.Array && tp.GetArrayLength() > 0)
            title = tp.EnumerateArray().First().GetString() ?? "";

        // Container / journal
        string container = "";
        if (msg.TryGetProperty("container-title", out var cp) && cp.ValueKind == JsonValueKind.Array && cp.GetArrayLength() > 0)
            container = cp.EnumerateArray().First().GetString() ?? "";

        if (!string.IsNullOrEmpty(authors))
            lines.Add(year > 0 ? $"{authors} ({year})" : authors);
        if (!string.IsNullOrEmpty(title)) lines.Add(title);
        if (!string.IsNullOrEmpty(container)) lines.Add(container);
        lines.Add($"doi:{doi}");

        Cache[doi] = lines;
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
