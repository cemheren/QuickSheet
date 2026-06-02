using System.Text.Json;
using System.Text.RegularExpressions;

// QuickSheet unit-conversion extension.
// Cell usage: unit: 5 km to miles | unit: 100 F to C | unit: 1 GB to MB
// Supports: length, mass, temperature, volume, speed, data, area, time.
// Pure computation — no network, no NuGet deps, no state.

Console.OutputEncoding = System.Text.Encoding.UTF8;
Console.WriteLine(JsonSerializer.Serialize(new
{
    type = "register",
    prefix = "unit",
    name = "Unit Converter",
    version = "1.0.0",
}));
Console.Out.Flush();

// --- Data tables (must precede usage in local functions) ---

var Aliases = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
{
    ["meter"] = "m", ["metre"] = "m", ["meters"] = "m", ["metres"] = "m",
    ["kilometer"] = "km", ["kilometre"] = "km", ["kilometers"] = "km",
    ["centimeter"] = "cm", ["centimetre"] = "cm",
    ["millimeter"] = "mm", ["millimetre"] = "mm",
    ["mile"] = "mi", ["miles"] = "mi",
    ["yard"] = "yd", ["yards"] = "yd",
    ["foot"] = "ft", ["feet"] = "ft",
    ["inch"] = "in", ["inches"] = "in",
    ["nautical mile"] = "nmi", ["nm"] = "nmi",
    ["kilogram"] = "kg", ["kilograms"] = "kg", ["kilo"] = "kg",
    ["gram"] = "g", ["grams"] = "g",
    ["milligram"] = "mg", ["milligrams"] = "mg",
    ["pound"] = "lb", ["pounds"] = "lb", ["lbs"] = "lb",
    ["ounce"] = "oz", ["ounces"] = "oz",
    ["ton"] = "t", ["tonne"] = "t", ["metric ton"] = "t",
    ["stone"] = "st",
    ["liter"] = "l", ["litre"] = "l", ["liters"] = "l", ["litres"] = "l",
    ["milliliter"] = "ml", ["millilitre"] = "ml",
    ["gallon"] = "gal", ["gallons"] = "gal",
    ["quart"] = "qt", ["quarts"] = "qt",
    ["pint"] = "pt", ["pints"] = "pt",
    ["cups"] = "cup", ["tablespoon"] = "tbsp", ["teaspoon"] = "tsp",
    ["fluid ounce"] = "floz", ["fl oz"] = "floz",
    ["km/h"] = "kmh", ["kph"] = "kmh", ["kmph"] = "kmh",
    ["mph"] = "mph", ["mi/h"] = "mph",
    ["m/s"] = "ms", ["meters/second"] = "ms",
    ["knot"] = "kt", ["knots"] = "kt", ["kn"] = "kt",
    ["celsius"] = "c", ["centigrade"] = "c", ["°c"] = "c",
    ["fahrenheit"] = "f", ["°f"] = "f",
    ["kelvin"] = "k",
    ["byte"] = "b", ["bytes"] = "b",
    ["kilobyte"] = "kb", ["kilobytes"] = "kb",
    ["megabyte"] = "mb", ["megabytes"] = "mb",
    ["gigabyte"] = "gb", ["gigabytes"] = "gb",
    ["terabyte"] = "tb", ["terabytes"] = "tb",
    ["kibibyte"] = "kib", ["mebibyte"] = "mib", ["gibibyte"] = "gib",
    ["square meter"] = "sqm", ["sq m"] = "sqm", ["m2"] = "sqm", ["m²"] = "sqm",
    ["square foot"] = "sqft", ["sq ft"] = "sqft", ["ft2"] = "sqft", ["ft²"] = "sqft",
    ["acre"] = "acre", ["acres"] = "acre",
    ["hectare"] = "ha", ["hectares"] = "ha",
    ["second"] = "sec", ["seconds"] = "sec",
    ["minute"] = "min", ["minutes"] = "min",
    ["hour"] = "hr", ["hours"] = "hr",
    ["day"] = "day", ["days"] = "day",
    ["week"] = "wk", ["weeks"] = "wk",
};

// Factor = how many base units one of this unit equals.
// Base: m (length), kg (mass), l (volume), m/s (speed), byte (data), sqm (area), sec (time).
var ToBase = new Dictionary<string, UnitInfo>(StringComparer.OrdinalIgnoreCase)
{
    ["m"] = new("length", 1), ["km"] = new("length", 1000),
    ["cm"] = new("length", 0.01), ["mm"] = new("length", 0.001),
    ["mi"] = new("length", 1609.344), ["yd"] = new("length", 0.9144),
    ["ft"] = new("length", 0.3048), ["in"] = new("length", 0.0254),
    ["nmi"] = new("length", 1852),
    ["kg"] = new("mass", 1), ["g"] = new("mass", 0.001),
    ["mg"] = new("mass", 0.000001), ["lb"] = new("mass", 0.453592),
    ["oz"] = new("mass", 0.0283495), ["t"] = new("mass", 1000),
    ["st"] = new("mass", 6.35029),
    ["l"] = new("volume", 1), ["ml"] = new("volume", 0.001),
    ["gal"] = new("volume", 3.78541), ["qt"] = new("volume", 0.946353),
    ["pt"] = new("volume", 0.473176), ["cup"] = new("volume", 0.236588),
    ["tbsp"] = new("volume", 0.014787), ["tsp"] = new("volume", 0.004929),
    ["floz"] = new("volume", 0.029574),
    ["ms"] = new("speed", 1), ["kmh"] = new("speed", 0.277778),
    ["mph"] = new("speed", 0.44704), ["kt"] = new("speed", 0.514444),
    ["b"] = new("data", 1), ["kb"] = new("data", 1000),
    ["mb"] = new("data", 1_000_000), ["gb"] = new("data", 1_000_000_000),
    ["tb"] = new("data", 1_000_000_000_000),
    ["kib"] = new("data", 1024), ["mib"] = new("data", 1_048_576),
    ["gib"] = new("data", 1_073_741_824),
    ["sqm"] = new("area", 1), ["sqft"] = new("area", 0.092903),
    ["acre"] = new("area", 4046.86), ["ha"] = new("area", 10000),
    ["sec"] = new("time", 1), ["min"] = new("time", 60),
    ["hr"] = new("time", 3600), ["day"] = new("time", 86400),
    ["wk"] = new("time", 604800),
};

// --- Event loop ---

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
    catch { /* swallow bad input */ }
}

// --- Local functions ---

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

    if (spec.Length == 0)
    {
        Emit(id, "unit: <value> <from> to <to>  (e.g. 5 km to miles)");
        return;
    }

    var match = Regex.Match(spec, @"^([\d.,]+)\s+(.+?)\s+to\s+(.+)$", RegexOptions.IgnoreCase);
    if (!match.Success)
    {
        Emit(id, $"err: expected '<value> <from> to <to>', got '{spec}'");
        return;
    }

    if (!double.TryParse(match.Groups[1].Value.Replace(",", ""), out double value))
    {
        Emit(id, $"err: cannot parse number '{match.Groups[1].Value}'");
        return;
    }

    string from = Normalize(match.Groups[2].Value.Trim());
    string to = Normalize(match.Groups[3].Value.Trim());

    if (IsTemp(from) && IsTemp(to))
    {
        double result = ConvertTemp(value, from, to);
        Emit(id, $"{value} {match.Groups[2].Value} = {result:G6} {match.Groups[3].Value}");
        return;
    }

    if (!ToBase.TryGetValue(from, out var fromInfo))
    {
        Emit(id, $"err: unknown unit '{match.Groups[2].Value}'");
        return;
    }
    if (!ToBase.TryGetValue(to, out var toInfo))
    {
        Emit(id, $"err: unknown unit '{match.Groups[3].Value}'");
        return;
    }
    if (fromInfo.Category != toInfo.Category)
    {
        Emit(id, $"err: cannot convert {fromInfo.Category} to {toInfo.Category}");
        return;
    }

    double baseValue = value * fromInfo.Factor;
    double result2 = baseValue / toInfo.Factor;
    Emit(id, $"{value} {match.Groups[2].Value} = {result2:G6} {match.Groups[3].Value}");
}

void Emit(string id, string text)
{
    var cells = new[] { new { r = 0, c = 0, v = text } };
    Console.WriteLine(JsonSerializer.Serialize(new { type = "write", id, cells }));
    Console.Out.Flush();
}

string Normalize(string u)
{
    u = u.ToLowerInvariant().Trim().TrimEnd('.');
    if (u.EndsWith("es") && !u.EndsWith("ches") && !u.EndsWith("ses"))
        u = u[..^2];
    else if (u.EndsWith("s") && !u.EndsWith("ius") && !u.EndsWith("ss"))
        u = u[..^1];
    return Aliases.TryGetValue(u, out var canonical) ? canonical : u;
}

static bool IsTemp(string u) => u is "c" or "f" or "k";

static double ConvertTemp(double v, string from, string to)
{
    double c = from switch
    {
        "c" => v,
        "f" => (v - 32) * 5.0 / 9.0,
        "k" => v - 273.15,
        _ => v
    };
    return to switch
    {
        "c" => c,
        "f" => c * 9.0 / 5.0 + 32,
        "k" => c + 273.15,
        _ => c
    };
}

// --- Types ---

record UnitInfo(string Category, double Factor);
