using System.Collections.Concurrent;
using System.Text.Json;
using System.Text.Json.Serialization;

/// <summary>
/// QuickSheet Define Extension — reads JSON-lines from stdin, writes JSON-lines to stdout.
/// Registers the "def" prefix. Looks up an English word using the free dictionaryapi.dev
/// service (no key required). Returns part-of-speech and first definition for each sense.
/// Usage: `def: laconic, 1, 4` → 4 rows with definitions.
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

    // Cache lookups for 24h — definitions don't change.
    private static readonly ConcurrentDictionary<string, (DateTime, List<string>)> Cache = new(StringComparer.OrdinalIgnoreCase);
    private static readonly TimeSpan CacheTtl = TimeSpan.FromHours(24);

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
            prefix = "def",
            name = "Dictionary Lookup",
            version = "1.0.0"
        });
        SendLog("Define extension registered with prefix 'def'");
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
            WriteCells(id, new[] { new[] { "def: <word>" } });
            return;
        }

        string word = extParams[0].Trim();

        try
        {
            var senses = FetchDefinitions(word);
            if (senses.Count == 0)
            {
                WriteCells(id, new[] { new[] { $"{word}: not found" } });
                return;
            }

            var rows = new List<string[]> { new[] { word } };
            foreach (string s in senses) rows.Add(new[] { s });
            while (rows.Count < gridRows) rows.Add(new[] { "" });
            if (rows.Count > gridRows) rows = rows.Take(gridRows).ToList();
            WriteCells(id, rows);
        }
        catch (Exception ex)
        {
            WriteCells(id, new[] { new[] { $"err: {ex.Message}" } });
        }
    }

    static List<string> FetchDefinitions(string word)
    {
        if (Cache.TryGetValue(word, out var cached) && DateTime.UtcNow - cached.Item1 < CacheTtl)
            return cached.Item2;

        string url = $"https://api.dictionaryapi.dev/api/v2/entries/en/{Uri.EscapeDataString(word.ToLowerInvariant())}";
        var resp = Http.GetAsync(url).GetAwaiter().GetResult();
        if (!resp.IsSuccessStatusCode) { Cache[word] = (DateTime.UtcNow, []); return []; }
        string json = resp.Content.ReadAsStringAsync().GetAwaiter().GetResult();

        var senses = new List<string>();
        using var doc = JsonDocument.Parse(json);
        if (doc.RootElement.ValueKind != JsonValueKind.Array) return senses;

        foreach (var entry in doc.RootElement.EnumerateArray())
        {
            if (!entry.TryGetProperty("meanings", out var meanings)) continue;
            foreach (var meaning in meanings.EnumerateArray())
            {
                string pos = meaning.TryGetProperty("partOfSpeech", out var p) ? p.GetString() ?? "" : "";
                if (!meaning.TryGetProperty("definitions", out var defs)) continue;
                foreach (var d in defs.EnumerateArray())
                {
                    string def = d.TryGetProperty("definition", out var dt) ? dt.GetString() ?? "" : "";
                    if (string.IsNullOrWhiteSpace(def)) continue;
                    string posTag = string.IsNullOrEmpty(pos) ? "" : $"({pos}) ";
                    senses.Add($"{posTag}{def}");
                    break; // first def per part-of-speech is enough
                }
            }
        }

        Cache[word] = (DateTime.UtcNow, senses);
        return senses;
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
