using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

/// <summary>
/// QuickSheet TLDR Extension — show tldr-pages command cheatsheets inline.
/// Fetches from the tldr-pages GitHub raw content (no API key needed).
/// Usage: tldr: tar | tldr: git-rebase | tldr: ffmpeg
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

    // Cache to avoid repeated fetches for the same command
    private static readonly Dictionary<string, (DateTime fetched, string[] lines)> Cache = new();
    private static readonly TimeSpan CacheTtl = TimeSpan.FromMinutes(30);

    // tldr-pages platforms to search (in order)
    private static readonly string[] Platforms = { "common", "linux", "osx", "windows" };

    static void Main()
    {
        Console.OutputEncoding = Encoding.UTF8;
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
            prefix = "tldr",
            name = "TLDR Pages",
            version = "1.0.0"
        });
        SendLog("TLDR extension registered with prefix 'tldr'");
    }

    static void HandleActivate(JsonElement root)
    {
        string id = root.TryGetProperty("id", out var idProp) ? idProp.GetString() ?? "" : "";
        int gridRows = root.TryGetProperty("gridRows", out var gr) ? gr.GetInt32() : 12;

        string[] extParams = [];
        if (root.TryGetProperty("params", out var p) && p.ValueKind == JsonValueKind.Array)
            extParams = p.EnumerateArray().Select(x => x.GetString() ?? "").ToArray();

        if (extParams.Length == 0 || string.IsNullOrWhiteSpace(extParams[0]))
        {
            WriteCells(id, new[] { new[] { "tldr: <command>" }, new[] { "e.g. tldr: tar" } });
            return;
        }

        string command = extParams[0].Trim().ToLowerInvariant().Replace(' ', '-');

        try
        {
            string[] lines = FetchTldr(command);
            if (lines.Length == 0)
            {
                WriteCells(id, new[] { new[] { $"⚠ No tldr page for '{command}'" } });
                return;
            }

            var rows = FormatForGrid(lines, gridRows);
            WriteCells(id, rows);
        }
        catch (Exception ex)
        {
            WriteCells(id, new[] { new[] { $"⚠ {ex.Message}" } });
        }
    }

    static string[] FetchTldr(string command)
    {
        // Check cache
        if (Cache.TryGetValue(command, out var cached) && DateTime.UtcNow - cached.fetched < CacheTtl)
            return cached.lines;

        foreach (var platform in Platforms)
        {
            string url = $"https://raw.githubusercontent.com/tldr-pages/tldr/main/pages/{platform}/{command}.md";
            try
            {
                var response = Http.GetAsync(url).GetAwaiter().GetResult();
                if (response.IsSuccessStatusCode)
                {
                    string content = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                    var lines = content.Split('\n', StringSplitOptions.None);
                    Cache[command] = (DateTime.UtcNow, lines);
                    return lines;
                }
            }
            catch
            {
                // Try next platform
            }
        }

        Cache[command] = (DateTime.UtcNow, Array.Empty<string>());
        return Array.Empty<string>();
    }

    static List<string[]> FormatForGrid(string[] mdLines, int maxRows)
    {
        var rows = new List<string[]>();

        foreach (var rawLine in mdLines)
        {
            string line = rawLine.TrimEnd();

            // Skip empty lines at the top and the title line (# command)
            if (rows.Count == 0 && (string.IsNullOrWhiteSpace(line) || line.StartsWith("# ")))
                continue;

            if (line.StartsWith("> "))
            {
                // Description line
                rows.Add(new[] { line[2..] });
            }
            else if (line.StartsWith("- "))
            {
                // Example description
                rows.Add(new[] { $"▸ {line[2..]}" });
            }
            else if (line.StartsWith("`") && line.EndsWith("`"))
            {
                // Code example
                rows.Add(new[] { $"  {line[1..^1]}" });
            }
            else if (!string.IsNullOrWhiteSpace(line))
            {
                rows.Add(new[] { line });
            }

            if (rows.Count >= maxRows) break;
        }

        // Pad to fill grid
        while (rows.Count < maxRows) rows.Add(new[] { "" });
        return rows;
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
