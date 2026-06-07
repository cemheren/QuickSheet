using System.Globalization;
using System.Text.Json;
using System.Text.Json.Serialization;

/// <summary>
/// QuickSheet Epoch Extension — converts between Unix timestamps and human-readable dates.
/// Protocol: JSON-lines on stdin/stdout.
/// Usage:
///   epoch: now         → current epoch + ISO date
///   epoch: 1717740000  → human-readable UTC date
///   epoch: 2026-06-07  → epoch seconds
///   epoch: +30d        → epoch + date 30 days from now
///   epoch: -7d         → epoch + date 7 days ago
/// </summary>
class Program
{
    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        WriteIndented = false
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
            prefix = "epoch",
            name = "Epoch Converter",
            version = "1.0.0"
        });
        SendLog("Epoch extension registered with prefix 'epoch'");
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
            WriteCells(id, BuildRows(new[] { "epoch: now | <timestamp> | <date> | +Nd | -Nd" }, gridRows));
            return;
        }

        string input = extParams[0].Trim();
        try
        {
            var result = Convert(input);
            WriteCells(id, BuildRows(result, gridRows));
        }
        catch (Exception ex)
        {
            WriteCells(id, BuildRows(new[] { $"err: {ex.Message}" }, gridRows));
        }
    }

    static string[] Convert(string input)
    {
        // "now" — show current epoch + ISO date
        if (input.Equals("now", StringComparison.OrdinalIgnoreCase))
        {
            var now = DateTimeOffset.UtcNow;
            return new[]
            {
                $"⏱ {now.ToUnixTimeSeconds()}",
                now.ToString("yyyy-MM-dd HH:mm:ss UTC"),
                $"ms: {now.ToUnixTimeMilliseconds()}"
            };
        }

        // Relative offset: +30d, -7d, +2h, -30m
        if ((input.StartsWith('+') || input.StartsWith('-')) && input.Length >= 2)
        {
            var offset = ParseRelative(input);
            if (offset.HasValue)
            {
                var target = DateTimeOffset.UtcNow.Add(offset.Value);
                return new[]
                {
                    $"⏱ {target.ToUnixTimeSeconds()}",
                    target.ToString("yyyy-MM-dd HH:mm:ss UTC"),
                    $"({input} from now)"
                };
            }
        }

        // Pure numeric — treat as epoch seconds
        if (long.TryParse(input, out long epochSec))
        {
            // Heuristic: if > 1e12 it's probably milliseconds
            DateTimeOffset dto;
            string suffix;
            if (epochSec > 1_000_000_000_000L)
            {
                dto = DateTimeOffset.FromUnixTimeMilliseconds(epochSec);
                suffix = "(ms)";
            }
            else
            {
                dto = DateTimeOffset.FromUnixTimeSeconds(epochSec);
                suffix = "(s)";
            }
            var age = DateTimeOffset.UtcNow - dto;
            string agoStr = FormatAge(age);
            return new[]
            {
                dto.ToString("yyyy-MM-dd HH:mm:ss UTC"),
                $"{suffix} {agoStr}",
                dto.ToLocalTime().ToString("ddd, dd MMM yyyy")
            };
        }

        // Try parse as date string
        if (DateTimeOffset.TryParse(input, CultureInfo.InvariantCulture,
                DateTimeStyles.AssumeUniversal, out var parsed))
        {
            return new[]
            {
                $"⏱ {parsed.ToUnixTimeSeconds()}",
                $"ms: {parsed.ToUnixTimeMilliseconds()}",
                parsed.ToString("ddd, dd MMM yyyy HH:mm UTC")
            };
        }

        return new[] { $"Cannot parse: {input}", "Try: now, 1717740000, 2026-06-07, +30d" };
    }

    static TimeSpan? ParseRelative(string input)
    {
        char unit = char.ToLower(input[^1]);
        if (!int.TryParse(input[..^1], out int amount)) return null;
        return unit switch
        {
            'd' => TimeSpan.FromDays(amount),
            'h' => TimeSpan.FromHours(amount),
            'm' => TimeSpan.FromMinutes(amount),
            's' => TimeSpan.FromSeconds(amount),
            _ => null
        };
    }

    static string FormatAge(TimeSpan age)
    {
        if (age.TotalSeconds < 0) age = age.Negate();
        bool future = age.TotalSeconds < 0;
        string label = future ? "from now" : "ago";
        if (age.TotalDays >= 365) return $"{(int)(age.TotalDays / 365)}y {label}";
        if (age.TotalDays >= 30) return $"{(int)(age.TotalDays / 30)}mo {label}";
        if (age.TotalDays >= 1) return $"{(int)age.TotalDays}d {label}";
        if (age.TotalHours >= 1) return $"{(int)age.TotalHours}h {label}";
        if (age.TotalMinutes >= 1) return $"{(int)age.TotalMinutes}min {label}";
        return $"{(int)age.TotalSeconds}s {label}";
    }

    static List<string[]> BuildRows(string[] lines, int gridRows)
    {
        var rows = lines.Select(l => new[] { l }).ToList();
        while (rows.Count < gridRows) rows.Add(new[] { "" });
        if (rows.Count > gridRows) rows = rows.Take(gridRows).ToList();
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
