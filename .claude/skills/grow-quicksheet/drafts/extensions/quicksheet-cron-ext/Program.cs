using System.Text.Json;
using System.Text.Json.Serialization;

/// <summary>
/// QuickSheet Cron Extension — parses cron expressions into human-readable descriptions.
/// Supports standard 5-field cron (minute, hour, day-of-month, month, day-of-week)
/// and common shortcuts (@hourly, @daily, @weekly, @monthly, @yearly, @reboot).
/// Usage: `cron: */5 * * * *` → "Every 5 minutes"
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
            prefix = "cron",
            name = "Cron Parser",
            version = "1.0.0"
        });
        SendLog("Cron extension registered with prefix 'cron'");
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
            WriteCells(id, new[] { new[] { "cron: <expression>" }, new[] { "e.g. */5 * * * *" } }, gridRows);
            return;
        }

        // Rejoin params since cron expressions contain spaces that get split on commas
        string expression = string.Join(",", extParams).Trim();

        try
        {
            string description = CronParser.Describe(expression);
            string nextRun = CronParser.NextOccurrence(expression);

            var rows = new List<string[]>
            {
                new[] { $"⏰ {expression}" },
                new[] { description },
                new[] { nextRun != "" ? $"Next: {nextRun}" : "" }
            };
            WriteCells(id, rows, gridRows);
        }
        catch (Exception ex)
        {
            WriteCells(id, new[] { new[] { $"err: {ex.Message}" } }, gridRows);
        }
    }

    static void WriteCells(string id, IEnumerable<string[]> rows, int gridRows)
    {
        var list = rows.ToList();
        while (list.Count < gridRows) list.Add(new[] { "" });
        if (list.Count > gridRows) list = list.Take(gridRows).ToList();
        SendJson(new { type = "write", id, cells = list });
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

/// <summary>
/// Pure cron expression parser. No dependencies. Handles 5-field standard cron
/// and common shortcut aliases.
/// </summary>
static class CronParser
{
    private static readonly string[] MonthNames =
        { "", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };
    private static readonly string[] DayNames =
        { "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat" };

    public static string Describe(string expression)
    {
        expression = expression.Trim();

        // Handle shortcuts
        switch (expression.ToLowerInvariant())
        {
            case "@yearly" or "@annually": return "Once a year (Jan 1, midnight)";
            case "@monthly": return "Once a month (1st, midnight)";
            case "@weekly": return "Once a week (Sun, midnight)";
            case "@daily" or "@midnight": return "Once a day (midnight)";
            case "@hourly": return "Once an hour (at :00)";
            case "@reboot": return "Once at startup";
        }

        var parts = expression.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length < 5)
            return "Invalid: need 5 fields (min hr dom mon dow)";
        if (parts.Length > 5)
            // Take first 5 only (ignore optional 6th "year" field some systems support)
            parts = parts[..5];

        string minute = parts[0];
        string hour = parts[1];
        string dom = parts[2];
        string month = parts[3];
        string dow = parts[4];

        var segments = new List<string>();

        // Frequency / minute description
        if (minute == "*" && hour == "*")
            segments.Add("Every minute");
        else if (minute.StartsWith("*/"))
            segments.Add($"Every {minute[2..]} min");
        else if (hour == "*")
            segments.Add($"At :{minute.PadLeft(2, '0')} every hour");
        else if (hour.StartsWith("*/"))
            segments.Add($"At :{minute.PadLeft(2, '0')} every {hour[2..]} hrs");
        else
            segments.Add($"At {FormatTime(hour, minute)}");

        // Day-of-month
        if (dom != "*" && dom != "?")
        {
            if (dom.Contains(','))
                segments.Add($"on days {dom}");
            else if (dom.Contains('-'))
                segments.Add($"on days {dom}");
            else if (dom.StartsWith("*/"))
                segments.Add($"every {dom[2..]} days");
            else
                segments.Add($"on day {dom}");
        }

        // Month
        if (month != "*")
        {
            if (int.TryParse(month, out int m) && m >= 1 && m <= 12)
                segments.Add($"in {MonthNames[m]}");
            else if (month.Contains(','))
                segments.Add($"in months {month}");
            else
                segments.Add($"in {month}");
        }

        // Day-of-week
        if (dow != "*" && dow != "?")
        {
            string dowDesc = FormatDow(dow);
            segments.Add($"on {dowDesc}");
        }

        return string.Join(", ", segments);
    }

    public static string NextOccurrence(string expression)
    {
        expression = expression.Trim();

        // Handle shortcuts
        switch (expression.ToLowerInvariant())
        {
            case "@reboot": return "next reboot";
        }

        var parts = expression.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length < 5) return "";

        // Simple next-run calculation for common patterns
        try
        {
            var now = DateTime.Now;
            var next = FindNext(parts[0], parts[1], parts[2], parts[3], parts[4], now);
            if (next.HasValue)
            {
                var diff = next.Value - now;
                if (diff.TotalMinutes < 60)
                    return $"{next.Value:HH:mm} (in {(int)diff.TotalMinutes}m)";
                if (diff.TotalHours < 24)
                    return $"{next.Value:HH:mm} (in {(int)diff.TotalHours}h {diff.Minutes}m)";
                return $"{next.Value:ddd MMM d HH:mm}";
            }
        }
        catch { }
        return "";
    }

    private static DateTime? FindNext(string minF, string hrF, string domF, string monF, string dowF, DateTime from)
    {
        // Brute-force: check each minute in the next 48 hours
        var candidate = new DateTime(from.Year, from.Month, from.Day, from.Hour, from.Minute, 0).AddMinutes(1);
        var limit = from.AddHours(48);

        while (candidate < limit)
        {
            if (MatchesField(minF, candidate.Minute, 0, 59)
                && MatchesField(hrF, candidate.Hour, 0, 23)
                && MatchesField(domF, candidate.Day, 1, 31)
                && MatchesField(monF, candidate.Month, 1, 12)
                && MatchesDow(dowF, candidate.DayOfWeek))
            {
                return candidate;
            }
            candidate = candidate.AddMinutes(1);
        }
        return null;
    }

    private static bool MatchesField(string field, int value, int min, int max)
    {
        if (field == "*" || field == "?") return true;
        if (field.StartsWith("*/"))
        {
            if (int.TryParse(field[2..], out int step) && step > 0)
                return (value - min) % step == 0;
        }
        if (field.Contains(','))
        {
            return field.Split(',').Any(p => MatchesSingle(p.Trim(), value, min));
        }
        return MatchesSingle(field, value, min);
    }

    private static bool MatchesSingle(string part, int value, int min)
    {
        if (part.Contains('-'))
        {
            var range = part.Split('-');
            if (range.Length == 2 && int.TryParse(range[0], out int lo) && int.TryParse(range[1], out int hi))
                return value >= lo && value <= hi;
        }
        if (part.Contains('/'))
        {
            var stepParts = part.Split('/');
            if (stepParts.Length == 2 && int.TryParse(stepParts[1], out int step) && step > 0)
            {
                int start = stepParts[0] == "*" ? min : (int.TryParse(stepParts[0], out int s) ? s : min);
                return value >= start && (value - start) % step == 0;
            }
        }
        if (int.TryParse(part, out int exact))
            return value == exact;
        return false;
    }

    private static bool MatchesDow(string field, DayOfWeek dayOfWeek)
    {
        if (field == "*" || field == "?") return true;
        int dow = (int)dayOfWeek; // 0=Sun

        if (field.StartsWith("*/"))
        {
            if (int.TryParse(field[2..], out int step) && step > 0)
                return dow % step == 0;
        }
        if (field.Contains(','))
            return field.Split(',').Any(p => MatchesDowSingle(p.Trim(), dow));
        return MatchesDowSingle(field, dow);
    }

    private static bool MatchesDowSingle(string part, int dow)
    {
        if (part.Contains('-'))
        {
            var range = part.Split('-');
            if (range.Length == 2)
            {
                int lo = ParseDow(range[0]);
                int hi = ParseDow(range[1]);
                if (lo >= 0 && hi >= 0)
                    return lo <= hi ? (dow >= lo && dow <= hi) : (dow >= lo || dow <= hi);
            }
        }
        int val = ParseDow(part);
        return val >= 0 && val == dow;
    }

    private static int ParseDow(string s)
    {
        if (int.TryParse(s, out int n)) return n % 7;
        return s.ToLowerInvariant() switch
        {
            "sun" => 0, "mon" => 1, "tue" => 2, "wed" => 3,
            "thu" => 4, "fri" => 5, "sat" => 6, _ => -1
        };
    }

    private static string FormatTime(string hour, string minute)
    {
        if (int.TryParse(hour, out int h) && int.TryParse(minute, out int m))
        {
            string period = h >= 12 ? "PM" : "AM";
            int h12 = h == 0 ? 12 : h > 12 ? h - 12 : h;
            return $"{h12}:{m:D2} {period}";
        }
        return $"{hour}:{minute.PadLeft(2, '0')}";
    }

    private static string FormatDow(string dow)
    {
        if (dow.Contains(','))
        {
            var days = dow.Split(',').Select(d =>
            {
                int val = ParseDow(d.Trim());
                return val >= 0 && val < 7 ? DayNames[val] : d.Trim();
            });
            return string.Join(", ", days);
        }
        if (dow.Contains('-'))
        {
            var range = dow.Split('-');
            if (range.Length == 2)
            {
                int lo = ParseDow(range[0]);
                int hi = ParseDow(range[1]);
                if (lo >= 0 && lo < 7 && hi >= 0 && hi < 7)
                    return $"{DayNames[lo]}–{DayNames[hi]}";
            }
        }
        int v = ParseDow(dow);
        return v >= 0 && v < 7 ? DayNames[v] : dow;
    }
}
