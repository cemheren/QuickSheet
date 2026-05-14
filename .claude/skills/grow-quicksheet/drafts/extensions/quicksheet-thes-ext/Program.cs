using System.Collections.Concurrent;
using System.Text.Json;
using System.Text.Json.Serialization;

/// <summary>
/// QuickSheet Thesaurus Extension — reads JSON-lines from stdin, writes JSON-lines to stdout.
/// Registers the "thes" prefix. Fetches synonyms from the free Datamuse API.
/// Usage: `thes: laconic, 1, 6` → up to 5 synonyms grouped on lines.
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

    // Cache forever — synonyms don't change.
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
            prefix = "thes",
            name = "Thesaurus Lookup",
            version = "1.0.0"
        });
        SendLog("Thes extension registered with prefix 'thes'");
    }

    static void HandleActivate(JsonElement root)
    {
        string id = root.TryGetProperty("id", out var idProp) ? idProp.GetString() ?? "" : "";
        int gridRows = root.TryGetProperty("gridRows", out var gr) ? gr.GetInt32() : 6;

        string[] extParams = [];
        if (root.TryGetProperty("params", out var p) && p.ValueKind == JsonValueKind.Array)
            extParams = p.EnumerateArray().Select(x => x.GetString() ?? "").ToArray();

        if (extParams.Length == 0 || string.IsNullOrWhiteSpace(extParams[0]))
        {
            WriteCells(id, new[] { new[] { "thes: <word>" } });
            return;
        }

        string word = extParams[0].Trim();

        try
        {
            var syns = FetchSynonyms(word);
            var rows = new List<string[]> { new[] { word } };
            if (syns.Count == 0) rows.Add(new[] { "(no synonyms)" });
            else foreach (string s in syns) rows.Add(new[] { s });

            while (rows.Count < gridRows) rows.Add(new[] { "" });
            if (rows.Count > gridRows) rows = rows.Take(gridRows).ToList();
            WriteCells(id, rows);
        }
        catch (Exception ex)
        {
            WriteCells(id, new[] { new[] { $"err: {ex.Message}" } });
        }
    }

    static List<string> FetchSynonyms(string word)
    {
        if (Cache.TryGetValue(word, out var cached)) return cached;

        // rel_syn = synonyms specifically (vs ml= "means like" which is broader).
        string url = $"https://api.datamuse.com/words?rel_syn={Uri.EscapeDataString(word)}&max=8";
        var resp = Http.GetAsync(url).GetAwaiter().GetResult();
        resp.EnsureSuccessStatusCode();
        string json = resp.Content.ReadAsStringAsync().GetAwaiter().GetResult();

        var results = new List<string>();
        using var doc = JsonDocument.Parse(json);
        if (doc.RootElement.ValueKind != JsonValueKind.Array) { Cache[word] = results; return results; }
        foreach (var el in doc.RootElement.EnumerateArray())
        {
            if (el.TryGetProperty("word", out var w)) results.Add(w.GetString() ?? "");
            if (results.Count >= 8) break;
        }
        Cache[word] = results;
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
