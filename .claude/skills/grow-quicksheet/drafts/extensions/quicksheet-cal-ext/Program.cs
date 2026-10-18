using System.Globalization;
using System.Text.Json;
using System.Text.Json.Serialization;

/// <summary>
/// QuickSheet Calendar Extension — renders a mini text calendar in cells.
/// Prefix: "mcal". Usage:
///   mcal:             → current month
///   mcal: 2026-03    → March 2026
///   mcal: next       → next month
///   mcal: prev       → previous month
/// Highlights today with [DD] brackets.
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
            prefix = "mcal",
            name = "Month Calendar",
            version = "1.0.0"
        });
        SendLog("Month Calendar extension registered with prefix 'mcal'");
    }

    static void HandleActivate(JsonElement root)
    {
        string id = root.TryGetProperty("id", out var idProp) ? idProp.GetString() ?? "" : "";
        int gridRows = root.TryGetProperty("gridRows", out var gr) ? gr.GetInt32() : 8;

        string[] extParams = [];
        if (root.TryGetProperty("params", out var p) && p.ValueKind == JsonValueKind.Array)
            extParams = p.EnumerateArray().Select(x => x.GetString() ?? "").ToArray();

        var today = DateTime.Today;
        int year = today.Year;
        int month = today.Month;

        if (extParams.Length > 0 && !string.IsNullOrWhiteSpace(extParams[0]))
        {
            string arg = extParams[0].Trim().ToLowerInvariant();
            if (arg == "next")
            {
                var next = today.AddMonths(1);
                year = next.Year;
                month = next.Month;
            }
            else if (arg == "prev")
            {
                var prev = today.AddMonths(-1);
                year = prev.Year;
                month = prev.Month;
            }
            else if (DateTime.TryParseExact(arg, "yyyy-MM", CultureInfo.InvariantCulture,
                         DateTimeStyles.None, out var parsed))
            {
                year = parsed.Year;
                month = parsed.Month;
            }
            else if (int.TryParse(arg, out int monthNum) && monthNum >= 1 && monthNum <= 12)
            {
                month = monthNum;
            }
        }

        var rows = BuildCalendar(year, month, today);

        // Trim or pad to gridRows
        while (rows.Count < gridRows) rows.Add(new[] { "" });
        if (rows.Count > gridRows) rows = rows.Take(gridRows).ToList();

        WriteCells(id, rows);
    }

    static List<string[]> BuildCalendar(int year, int month, DateTime today)
    {
        var rows = new List<string[]>();
        string header = CultureInfo.InvariantCulture.DateTimeFormat.GetMonthName(month) + " " + year;
        rows.Add(new[] { header });
        rows.Add(new[] { "Mo Tu We Th Fr Sa Su" });

        int daysInMonth = DateTime.DaysInMonth(year, month);
        var firstDay = new DateTime(year, month, 1);
        // Monday=0 .. Sunday=6
        int startOffset = ((int)firstDay.DayOfWeek + 6) % 7;

        var weekCells = new List<string>();
        // Each day slot is 3 chars: " DD". Leading offset pads with spaces.
        string week = new string(' ', startOffset * 3);

        for (int day = 1; day <= daysInMonth; day++)
        {
            bool isToday = (year == today.Year && month == today.Month && day == today.Day);
            // Each slot is 3 chars wide ("DD "). Today uses "[DD]" but we
            // consume the trailing space of the previous slot visually.
            if (isToday)
                week += $"[{day,2}]";
            else
                week += $" {day,2}";

            int dayOfWeek = (startOffset + day - 1) % 7;
            if (dayOfWeek == 6 || day == daysInMonth)
            {
                rows.Add(new[] { week.TrimEnd() });
                week = "";
            }
        }

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
