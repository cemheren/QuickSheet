using System.Globalization;
using System.Text.Json;

// QuickSheet moon extension — ambient lunar-phase widget for your desktop grid.
//
// Cell usage examples:
//   moon:                  → today's phase: emoji, name, illumination %, moon age
//   moon: today            → same as bare "moon:"
//   moon: 2026-12-25       → phase on a specific date (YYYY-MM-DD)
//   moon: next full        → date of the next full moon from today
//   moon: next new         → date of the next new moon from today
//
// Everything is computed locally from a standard synodic-month approximation
// anchored to a known new moon (2000-01-06 18:14 UTC). No network, no API key,
// no state — perfect as an always-on wallpaper cell that never flickers offline.

const double SynodicMonth = 29.530588853;          // mean length of a lunar cycle, days
const double KnownNewMoonJd = 2451550.1;            // 2000-01-06 18:14 UTC, a reference new moon

Console.OutputEncoding = System.Text.Encoding.UTF8;
Console.WriteLine(JsonSerializer.Serialize(new
{
    type = "register",
    prefix = "moon",
    name = "Lunar Phase",
    version = "1.0.0",
}));
Console.Out.Flush();

string? line;
while ((line = Console.ReadLine()) != null)
{
    if (string.IsNullOrWhiteSpace(line)) continue;
    try
    {
        using var doc = JsonDocument.Parse(line);
        if (!doc.RootElement.TryGetProperty("type", out var t)) continue;
        if (t.GetString() != "activate") continue;
        HandleActivate(doc.RootElement);
    }
    catch { /* swallow malformed input lines */ }
}

void HandleActivate(JsonElement root)
{
    string id = root.TryGetProperty("id", out var idEl) ? idEl.GetString() ?? "" : "";

    string spec = "";
    if (root.TryGetProperty("params", out var p) && p.ValueKind == JsonValueKind.Array)
    {
        var en = p.EnumerateArray();
        if (en.MoveNext()) spec = en.Current.GetString() ?? "";
    }
    spec = spec.Trim();

    if (spec.Length == 0 || spec.Equals("today", StringComparison.OrdinalIgnoreCase))
    {
        Emit(id, DescribePhase(DateTime.UtcNow));
        return;
    }

    if (spec.StartsWith("next", StringComparison.OrdinalIgnoreCase))
    {
        string which = spec[4..].Trim().ToLowerInvariant();
        if (which is "full")
        {
            var d = NextPhase(DateTime.UtcNow, 0.5);
            Emit(id, $"🌕 next full moon: {d:yyyy-MM-dd}  ({(d - DateTime.UtcNow).Days}d away)");
            return;
        }
        if (which is "new" or "")
        {
            var d = NextPhase(DateTime.UtcNow, 0.0);
            Emit(id, $"🌑 next new moon: {d:yyyy-MM-dd}  ({(d - DateTime.UtcNow).Days}d away)");
            return;
        }
        Emit(id, "moon: next full | next new");
        return;
    }

    if (DateTime.TryParse(spec, CultureInfo.InvariantCulture,
            DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out var when))
    {
        Emit(id, DescribePhase(when));
        return;
    }

    Emit(id, "moon: [today | YYYY-MM-DD | next full | next new]");
}

string DescribePhase(DateTime utc)
{
    double age = MoonAgeDays(utc);                       // 0..SynodicMonth
    double frac = age / SynodicMonth;                    // 0..1 through the cycle
    double illum = (1 - Math.Cos(2 * Math.PI * frac)) / 2.0; // 0=new, 1=full
    var (emoji, name) = PhaseLabel(frac);
    return $"{emoji} {name}  ·  {illum * 100:0}% lit  ·  age {age:0.0}d";
}

(string emoji, string name) PhaseLabel(double frac)
{
    // Eight principal phases, each a 1/8 slice of the cycle centred on its phase point.
    double f = ((frac % 1) + 1) % 1;
    int slice = (int)Math.Floor(f * 8 + 0.5) % 8;
    return slice switch
    {
        0 => ("🌑", "New Moon"),
        1 => ("🌒", "Waxing Crescent"),
        2 => ("🌓", "First Quarter"),
        3 => ("🌔", "Waxing Gibbous"),
        4 => ("🌕", "Full Moon"),
        5 => ("🌖", "Waning Gibbous"),
        6 => ("🌗", "Last Quarter"),
        _ => ("🌘", "Waning Crescent"),
    };
}

double MoonAgeDays(DateTime utc)
{
    double jd = ToJulianDate(utc);
    double age = (jd - KnownNewMoonJd) % SynodicMonth;
    if (age < 0) age += SynodicMonth;
    return age;
}

// Find the next datetime (after `from`) at which the cycle fraction equals `targetFrac`
// (0.0 = new moon, 0.5 = full moon).
DateTime NextPhase(DateTime from, double targetFrac)
{
    double age = MoonAgeDays(from);
    double curFrac = age / SynodicMonth;
    double delta = targetFrac - curFrac;
    if (delta <= 0) delta += 1.0;
    double daysAhead = delta * SynodicMonth;
    return from.AddDays(daysAhead);
}

static double ToJulianDate(DateTime utc)
{
    int year = utc.Year, month = utc.Month;
    double day = utc.Day + (utc.Hour + utc.Minute / 60.0 + utc.Second / 3600.0) / 24.0;
    if (month <= 2) { year -= 1; month += 12; }
    int a = year / 100;
    int b = 2 - a + a / 4;
    return Math.Floor(365.25 * (year + 4716))
         + Math.Floor(30.6001 * (month + 1))
         + day + b - 1524.5;
}

void Emit(string id, string text)
{
    var cells = new[] { new { r = 0, c = 0, v = text } };
    Console.WriteLine(JsonSerializer.Serialize(new { type = "write", id, cells }));
    Console.Out.Flush();
}
