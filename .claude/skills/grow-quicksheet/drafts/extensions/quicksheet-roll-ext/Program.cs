using System.Text.Json;
using System.Text.RegularExpressions;

// QuickSheet roll extension — dice roller for tabletop GMs.
//
// Cell usage examples:
//   roll: 1d20
//   roll: 4d6 drop lowest          → 4d6, drop the lowest die, sum the rest (D&D stat-rolling)
//   roll: 2d6+3                    → 2d6 with a +3 modifier
//   roll: 1d100, encounters.txt    → roll d100, find the line in encounters.txt whose
//                                     leading "low-high" range covers the rolled value
//   roll: adv 1d20                 → roll twice, take the higher (D&D 5e advantage)
//   roll: dis 1d20                 → roll twice, take the lower
//
// encounters.txt format (one row per range, blank lines + # comments ignored):
//   1-5    Goblins ambush from the trees
//   6-20   Old shrine to a forgotten god
//   21-40  Wagon tracks heading west
//   ...
//
// Output: one cell with the rolled total + individual dice + table-row text if applicable.
//
// Protocol: emit register on startup; then handle activate. No state, no network, no auth.

Console.OutputEncoding = System.Text.Encoding.UTF8;
Console.WriteLine(JsonSerializer.Serialize(new
{
    type = "register",
    prefix = "roll",
    name = "Dice Roller",
    version = "1.0.0",
}));
Console.Out.Flush();

var rng = new Random();
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
    catch { /* swallow bad input lines */ }
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

    if (spec.Length == 0)
    {
        Emit(id, "roll: <NdM[+mod]> [drop lowest|drop highest] [, table.txt]");
        return;
    }

    // Split out optional table file (after a comma)
    string diceExpr = spec;
    string? tablePath = null;
    int comma = spec.IndexOf(',');
    if (comma > 0)
    {
        diceExpr = spec[..comma].Trim();
        tablePath = spec[(comma + 1)..].Trim();
    }

    var result = RollExpression(diceExpr);
    if (result == null)
    {
        Emit(id, $"err: cannot parse '{diceExpr}'");
        return;
    }

    string display = $"= {result.Total}  [{string.Join(", ", result.Rolls)}]"
        + (result.Modifier != 0 ? $" + {result.Modifier}" : "")
        + (result.Note != null ? $"  ({result.Note})" : "");

    if (tablePath != null && File.Exists(tablePath))
    {
        string? row = TableLookup(tablePath, result.Total);
        if (row != null) display = $"{result.Total}  →  {row}";
    }

    Emit(id, display);
}

RollResult? RollExpression(string expr)
{
    expr = expr.Trim();

    bool adv = false, dis = false;
    if (expr.StartsWith("adv ", StringComparison.OrdinalIgnoreCase)) { adv = true; expr = expr[4..]; }
    else if (expr.StartsWith("dis ", StringComparison.OrdinalIgnoreCase)) { dis = true; expr = expr[4..]; }

    bool dropLow = false, dropHigh = false;
    int dropLowIdx = expr.IndexOf("drop lowest", StringComparison.OrdinalIgnoreCase);
    int dropHighIdx = expr.IndexOf("drop highest", StringComparison.OrdinalIgnoreCase);
    if (dropLowIdx > 0) { dropLow = true; expr = expr[..dropLowIdx].Trim(); }
    else if (dropHighIdx > 0) { dropHigh = true; expr = expr[..dropHighIdx].Trim(); }

    // Match e.g. "2d6+3" or "1d20" or "4d6"
    var m = Regex.Match(expr, @"^\s*(\d+)\s*[dD]\s*(\d+)\s*(?:([+-])\s*(\d+))?\s*$");
    if (!m.Success) return null;
    int n = int.Parse(m.Groups[1].Value);
    int faces = int.Parse(m.Groups[2].Value);
    int modifier = 0;
    if (m.Groups[3].Success)
    {
        int mod = int.Parse(m.Groups[4].Value);
        modifier = m.Groups[3].Value == "-" ? -mod : mod;
    }
    if (n < 1 || n > 100 || faces < 2 || faces > 1000) return null;

    int times = (adv || dis) ? 2 : 1;
    int bestTotal = adv ? int.MinValue : (dis ? int.MaxValue : 0);
    int[] bestRolls = Array.Empty<int>();
    string? note = adv ? "advantage" : (dis ? "disadvantage" : null);

    for (int t = 0; t < times; t++)
    {
        var rolls = new int[n];
        for (int i = 0; i < n; i++) rolls[i] = rng.Next(1, faces + 1);
        var working = (int[])rolls.Clone();
        if (dropLow && working.Length > 1)
        {
            Array.Sort(working);
            working = working[1..];
            note = (note == null ? "" : note + ", ") + "dropped lowest";
        }
        else if (dropHigh && working.Length > 1)
        {
            Array.Sort(working);
            working = working[..^1];
            note = (note == null ? "" : note + ", ") + "dropped highest";
        }
        int sum = working.Sum() + modifier;
        if (adv ? sum > bestTotal : dis ? sum < bestTotal : true)
        {
            bestTotal = sum;
            bestRolls = rolls;
        }
        if (!adv && !dis) break;
    }

    return new RollResult(bestTotal, bestRolls, modifier, note);
}

static string? TableLookup(string path, int roll)
{
    foreach (var raw in File.ReadAllLines(path))
    {
        var ln = raw.Trim();
        if (ln.Length == 0 || ln.StartsWith("#")) continue;
        // Leading "N-M  rest" or "N  rest"
        var m = Regex.Match(ln, @"^(\d+)(?:\s*-\s*(\d+))?\s+(.+)$");
        if (!m.Success) continue;
        int low = int.Parse(m.Groups[1].Value);
        int high = m.Groups[2].Success ? int.Parse(m.Groups[2].Value) : low;
        if (roll >= low && roll <= high) return m.Groups[3].Value.Trim();
    }
    return null;
}

void Emit(string id, string text)
{
    var cells = new[] { new { r = 0, c = 0, v = text } };
    Console.WriteLine(JsonSerializer.Serialize(new { type = "write", id, cells }));
    Console.Out.Flush();
}

record RollResult(int Total, int[] Rolls, int Modifier, string? Note);
