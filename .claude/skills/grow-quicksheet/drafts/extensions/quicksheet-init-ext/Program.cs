using System.Text.Json;

// QuickSheet init extension — combat initiative tracker for tabletop RPG GMs.
//
// Cell usage:
//   init: add Goblin 15        → add "Goblin" at initiative 15
//   init: add Ranger 18        → add "Ranger" at initiative 18
//   init: add Wizard 12        → add "Wizard" at initiative 12
//   init: show                  → display sorted initiative order, highlight current turn
//   init: next                  → advance to the next combatant
//   init: prev                  → go back to the previous combatant
//   init: remove Goblin         → remove a combatant
//   init: clear                 → reset the tracker
//   init: round                 → show the current round number
//
// State is kept in-memory for the session. The tracker sorts descending by
// initiative (highest goes first, standard D&D/Pathfinder convention).
//
// Protocol: emit register on startup; then handle activate. No network, no auth.

Console.OutputEncoding = System.Text.Encoding.UTF8;
Console.WriteLine(JsonSerializer.Serialize(new
{
    type = "register",
    prefix = "init",
    name = "Initiative Tracker",
    version = "1.0.0",
}));
Console.Out.Flush();

var combatants = new List<(string Name, int Initiative)>();
int currentIndex = 0;
int round = 1;

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
        Emit(id, "init: add <name> <roll> | show | next | prev | remove <name> | clear | round");
        return;
    }

    var parts = spec.Split(' ', StringSplitOptions.RemoveEmptyEntries);
    string cmd = parts[0].ToLowerInvariant();

    switch (cmd)
    {
        case "add":
            HandleAdd(id, parts);
            break;
        case "show":
            HandleShow(id);
            break;
        case "next":
            HandleNext(id);
            break;
        case "prev":
            HandlePrev(id);
            break;
        case "remove":
        case "rm":
            HandleRemove(id, parts);
            break;
        case "clear":
        case "reset":
            HandleClear(id);
            break;
        case "round":
            Emit(id, $"Round {round}");
            break;
        default:
            // Try parsing as "Name Initiative" shorthand (e.g. "init: Goblin 15")
            if (parts.Length >= 2 && int.TryParse(parts[^1], out int shortInit))
            {
                string shortName = string.Join(" ", parts[..^1]);
                AddCombatant(shortName, shortInit);
                HandleShow(id);
            }
            else
            {
                Emit(id, $"Unknown command: {cmd}. Use: add/show/next/prev/remove/clear/round");
            }
            break;
    }
}

void HandleAdd(string id, string[] parts)
{
    // init: add <name> <initiative>
    // Name can be multi-word: "init: add Goblin Archer 15"
    if (parts.Length < 3 || !int.TryParse(parts[^1], out int init))
    {
        Emit(id, "Usage: init: add <name> <initiative-number>");
        return;
    }
    string name = string.Join(" ", parts[1..^1]);
    AddCombatant(name, init);
    HandleShow(id);
}

void AddCombatant(string name, int init)
{
    // Remove existing entry with same name (case-insensitive) to allow re-rolling
    combatants.RemoveAll(c => c.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
    combatants.Add((name, init));
    // Sort descending by initiative
    combatants.Sort((a, b) => b.Initiative.CompareTo(a.Initiative));
    // Reset current index to top when roster changes
    currentIndex = 0;
    round = 1;
}

void HandleShow(string id)
{
    if (combatants.Count == 0)
    {
        Emit(id, "(empty) — use init: add <name> <roll>");
        return;
    }

    var lines = new List<string>();
    lines.Add($"R{round}:");
    for (int i = 0; i < combatants.Count; i++)
    {
        string marker = i == currentIndex ? "►" : " ";
        lines.Add($"{marker} {combatants[i].Initiative,2} {combatants[i].Name}");
    }
    Emit(id, string.Join(" | ", lines));
}

void HandleNext(string id)
{
    if (combatants.Count == 0)
    {
        Emit(id, "(empty)");
        return;
    }
    currentIndex++;
    if (currentIndex >= combatants.Count)
    {
        currentIndex = 0;
        round++;
    }
    HandleShow(id);
}

void HandlePrev(string id)
{
    if (combatants.Count == 0)
    {
        Emit(id, "(empty)");
        return;
    }
    currentIndex--;
    if (currentIndex < 0)
    {
        currentIndex = combatants.Count - 1;
        if (round > 1) round--;
    }
    HandleShow(id);
}

void HandleRemove(string id, string[] parts)
{
    if (parts.Length < 2)
    {
        Emit(id, "Usage: init: remove <name>");
        return;
    }
    string name = string.Join(" ", parts[1..]);
    int removed = combatants.RemoveAll(c => c.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
    if (removed == 0)
    {
        Emit(id, $"'{name}' not found in tracker");
        return;
    }
    if (currentIndex >= combatants.Count && combatants.Count > 0)
        currentIndex = 0;
    HandleShow(id);
}

void HandleClear(string id)
{
    combatants.Clear();
    currentIndex = 0;
    round = 1;
    Emit(id, "Initiative tracker cleared.");
}

void Emit(string id, string text)
{
    var cells = new[] { new { r = 0, c = 0, v = text } };
    Console.WriteLine(JsonSerializer.Serialize(new { type = "write", id, cells }));
    Console.Out.Flush();
}
