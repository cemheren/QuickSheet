using System.Text.Json;
using System.Text.Json.Serialization;

/// <summary>
/// QuickSheet Timezone Extension — converts a time from one zone to multiple target zones.
/// Usage: `tz: 3:00pm EST` or `tz: 14:30 Europe/London` or `tz: now, 1, 6`
/// Outputs the equivalent time in a curated set of major zones.
/// </summary>
class Program
{
    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        WriteIndented = false
    };

    // Common abbreviation → IANA mapping (most popular zones)
    private static readonly Dictionary<string, string> Abbreviations = new(StringComparer.OrdinalIgnoreCase)
    {
        ["EST"] = "America/New_York",
        ["EDT"] = "America/New_York",
        ["CST"] = "America/Chicago",
        ["CDT"] = "America/Chicago",
        ["MST"] = "America/Denver",
        ["MDT"] = "America/Denver",
        ["PST"] = "America/Los_Angeles",
        ["PDT"] = "America/Los_Angeles",
        ["UTC"] = "Etc/UTC",
        ["GMT"] = "Etc/GMT",
        ["CET"] = "Europe/Berlin",
        ["CEST"] = "Europe/Berlin",
        ["BST"] = "Europe/London",
        ["IST"] = "Asia/Kolkata",
        ["JST"] = "Asia/Tokyo",
        ["KST"] = "Asia/Seoul",
        ["CST_CN"] = "Asia/Shanghai",
        ["AEST"] = "Australia/Sydney",
        ["AEDT"] = "Australia/Sydney",
        ["NZST"] = "Pacific/Auckland",
        ["NZDT"] = "Pacific/Auckland",
    };

    // Default target zones for output
    private static readonly (string Label, string IanaId)[] DefaultTargets =
    {
        ("UTC", "Etc/UTC"),
        ("New York", "America/New_York"),
        ("London", "Europe/London"),
        ("Berlin", "Europe/Berlin"),
        ("Tokyo", "Asia/Tokyo"),
        ("Sydney", "Australia/Sydney"),
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
            prefix = "tz",
            name = "Timezone Converter",
            version = "1.0.0"
        });
        SendLog("Timezone extension registered with prefix 'tz'");
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
            WriteCells(id, new[] { new[] { "tz: <time> <zone>" }, new[] { "e.g. tz: 3pm EST" } });
            return;
        }

        string input = extParams[0].Trim();

        try
        {
            var (time, sourceZone) = ParseInput(input);
            var rows = new List<string[]>();

            rows.Add(new[] { $"🕐 {FormatTime(time)} {GetZoneAbbr(sourceZone)}" });

            foreach (var (label, ianaId) in DefaultTargets)
            {
                try
                {
                    var tz = TimeZoneInfo.FindSystemTimeZoneById(ianaId);
                    if (tz.Id == sourceZone.Id) continue;
                    var converted = ConvertTime(time, sourceZone, tz);
                    rows.Add(new[] { $"  {label}: {FormatTime(converted)}" });
                }
                catch { /* skip unavailable zones */ }
            }

            while (rows.Count < gridRows) rows.Add(new[] { "" });
            if (rows.Count > gridRows) rows = rows.Take(gridRows).ToList();
            WriteCells(id, rows);
        }
        catch (Exception ex)
        {
            WriteCells(id, new[] { new[] { $"err: {ex.Message}" } });
        }
    }

    static (DateTimeOffset time, TimeZoneInfo zone) ParseInput(string input)
    {
        // Handle "now"
        if (input.Equals("now", StringComparison.OrdinalIgnoreCase))
        {
            return (DateTimeOffset.UtcNow, TimeZoneInfo.Utc);
        }

        // Try to split into time part and zone part
        // Expected formats: "3:00pm EST", "14:30 Europe/London", "3pm PST"
        var parts = input.Split(' ', StringSplitOptions.RemoveEmptyEntries);

        string timePart;
        string zonePart;

        if (parts.Length == 1)
        {
            // Just a time, assume UTC
            timePart = parts[0];
            zonePart = "UTC";
        }
        else
        {
            timePart = parts[0];
            zonePart = string.Join(" ", parts.Skip(1));
        }

        // Parse the time
        var time = ParseTime(timePart);

        // Resolve zone
        var zone = ResolveZone(zonePart);

        // Combine: create a DateTimeOffset for today in the given zone
        var today = DateTimeOffset.UtcNow.Date;
        var dt = new DateTime(today.Year, today.Month, today.Day, time.Hour, time.Minute, 0, DateTimeKind.Unspecified);
        var offset = zone.GetUtcOffset(dt);
        var dto = new DateTimeOffset(dt, offset);

        return (dto, zone);
    }

    static DateTimeOffset ConvertTime(DateTimeOffset source, TimeZoneInfo sourceZone, TimeZoneInfo targetZone)
    {
        var utc = source.UtcDateTime;
        var targetOffset = targetZone.GetUtcOffset(utc);
        return new DateTimeOffset(utc + targetOffset, targetOffset);
    }

    static TimeOnly ParseTime(string s)
    {
        // Try standard formats
        string[] formats = {
            "h:mmtt", "hh:mmtt", "h:mm tt", "hh:mm tt",
            "htt", "hhtt", "h tt", "hh tt",
            "H:mm", "HH:mm", "H", "HH"
        };

        if (TimeOnly.TryParseExact(s, formats, System.Globalization.CultureInfo.InvariantCulture,
            System.Globalization.DateTimeStyles.None, out var result))
        {
            return result;
        }

        // Fallback: try general parse
        if (TimeOnly.TryParse(s, System.Globalization.CultureInfo.InvariantCulture, out result))
        {
            return result;
        }

        throw new FormatException($"Cannot parse time: '{s}'. Use formats like 3pm, 3:00pm, 15:00");
    }

    static TimeZoneInfo ResolveZone(string zone)
    {
        // Check abbreviation map
        if (Abbreviations.TryGetValue(zone, out var ianaId))
        {
            return TimeZoneInfo.FindSystemTimeZoneById(ianaId);
        }

        // Try direct IANA or Windows ID
        try
        {
            return TimeZoneInfo.FindSystemTimeZoneById(zone);
        }
        catch
        {
            throw new ArgumentException($"Unknown timezone: '{zone}'. Use EST, PST, UTC, Europe/London, etc.");
        }
    }

    static string GetZoneAbbr(TimeZoneInfo tz)
    {
        // Return a short display name
        if (tz.Id == "Etc/UTC" || tz.Id == "Etc/GMT") return "UTC";
        var abbr = tz.IsDaylightSavingTime(DateTimeOffset.UtcNow)
            ? tz.DaylightName : tz.StandardName;
        // On Linux, IANA names are returned — use the ID directly if name is long
        return abbr.Length > 5 ? tz.Id.Split('/').Last().Replace("_", " ") : abbr;
    }

    static string FormatTime(DateTimeOffset dto)
    {
        return dto.ToString("h:mm tt", System.Globalization.CultureInfo.InvariantCulture);
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
