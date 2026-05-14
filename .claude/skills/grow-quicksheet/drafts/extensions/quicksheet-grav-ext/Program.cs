using System.Collections.Concurrent;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

/// <summary>
/// QuickSheet Gravatar Extension — reads JSON-lines from stdin, writes JSON-lines to stdout.
/// Registers the "grav" prefix. Given an email, computes the MD5 hash and reports:
///   - Whether a Gravatar profile exists (en.gravatar.com/<hash>.json).
///   - Display name, location (if profile public).
///   - Avatar URL.
/// Usage: `grav: jane@example.com, 1, 4`.
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

    // 24h cache — profiles change slowly.
    private static readonly ConcurrentDictionary<string, (DateTime, Result)> Cache = new(StringComparer.OrdinalIgnoreCase);
    private static readonly TimeSpan CacheTtl = TimeSpan.FromHours(24);

    record Result(string Hash, string AvatarUrl, string? DisplayName, string? Location, bool ProfileFound);

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
            prefix = "grav",
            name = "Gravatar Lookup",
            version = "1.0.0"
        });
        SendLog("Grav extension registered with prefix 'grav'");
    }

    static void HandleActivate(JsonElement root)
    {
        string id = root.TryGetProperty("id", out var idProp) ? idProp.GetString() ?? "" : "";
        int gridRows = root.TryGetProperty("gridRows", out var gr) ? gr.GetInt32() : 4;

        string[] extParams = [];
        if (root.TryGetProperty("params", out var p) && p.ValueKind == JsonValueKind.Array)
            extParams = p.EnumerateArray().Select(x => x.GetString() ?? "").ToArray();

        if (extParams.Length == 0 || !extParams[0].Contains('@'))
        {
            WriteCells(id, new[] { new[] { "grav: <email>" } });
            return;
        }

        string email = extParams[0].Trim().ToLowerInvariant();

        try
        {
            var r = Lookup(email);
            var rows = new List<string[]>
            {
                new[] { email },
                new[] { r.ProfileFound ? (r.DisplayName ?? "(no name)") : "(no Gravatar profile)" },
                new[] { r.Location ?? "" },
                new[] { r.AvatarUrl }
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

    static Result Lookup(string email)
    {
        if (Cache.TryGetValue(email, out var cached) && DateTime.UtcNow - cached.Item1 < CacheTtl)
            return cached.Item2;

        string hash = Md5Hex(email);
        string avatarUrl = $"https://www.gravatar.com/avatar/{hash}?d=identicon";

        string profileUrl = $"https://en.gravatar.com/{hash}.json";
        string? name = null, location = null;
        bool found = false;
        try
        {
            var resp = Http.GetAsync(profileUrl).GetAwaiter().GetResult();
            if (resp.IsSuccessStatusCode)
            {
                string json = resp.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                using var doc = JsonDocument.Parse(json);
                if (doc.RootElement.TryGetProperty("entry", out var entry) && entry.ValueKind == JsonValueKind.Array && entry.GetArrayLength() > 0)
                {
                    var e = entry.EnumerateArray().First();
                    if (e.TryGetProperty("displayName", out var dn)) name = dn.GetString();
                    if (e.TryGetProperty("currentLocation", out var cl)) location = cl.GetString();
                    found = true;
                }
            }
        }
        catch { /* leave found=false */ }

        var result = new Result(hash, avatarUrl, name, location, found);
        Cache[email] = (DateTime.UtcNow, result);
        return result;
    }

    static string Md5Hex(string s)
    {
        byte[] bytes = MD5.HashData(Encoding.UTF8.GetBytes(s));
        var sb = new StringBuilder(32);
        foreach (var b in bytes) sb.Append(b.ToString("x2"));
        return sb.ToString();
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
