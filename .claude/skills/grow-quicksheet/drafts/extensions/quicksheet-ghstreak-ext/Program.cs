using System.Text;
using System.Text.Json;

// QuickSheet ghstreak extension — shows a GitHub user's commit-day streak +
// commits-today on the wallpaper. Reads only the public unauthenticated REST
// events endpoint (api.github.com/users/<u>/events/public).
//
// Cell usage:
//   ghstreak: <github-username>
//
// Output: one row, four columns:
//   @username  |  N today  |  🔥 Md streak  |  N in 90d
//
// Counts only PushEvent commits (payload.size summed). Streak is the number
// of consecutive UTC days, counting back from today, that have at least one
// public PushEvent. GitHub's events feed only goes back ~90 days / 300
// events; anything beyond that is invisible to this view.
//
// Caches 15 minutes per username. No auth — relies on public events; 60
// req/hr unauthenticated limit.

const string Endpoint = "https://api.github.com/users/{0}/events/public?per_page=100";
const int CacheTtlSeconds = 900;

var http = new HttpClient { Timeout = TimeSpan.FromSeconds(8) };
http.DefaultRequestHeaders.UserAgent.ParseAdd("quicksheet-ghstreak-ext/1.0 (+https://github.com/cemheren)");
http.DefaultRequestHeaders.Accept.ParseAdd("application/vnd.github+json");

Console.OutputEncoding = Encoding.UTF8;
Console.WriteLine(JsonSerializer.Serialize(new
{
    type = "register",
    prefix = "ghstreak",
    name = "GitHub Streak",
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
    catch { /* ignore malformed input */ }
}

void HandleActivate(JsonElement root)
{
    string id = root.TryGetProperty("id", out var idEl) ? idEl.GetString() ?? "" : "";

    string username = "";
    if (root.TryGetProperty("params", out var p) && p.ValueKind == JsonValueKind.Array)
    {
        var en = p.EnumerateArray();
        if (en.MoveNext()) username = en.Current.GetString() ?? "";
    }
    username = username.Trim();

    if (username.Length == 0)
    {
        Emit(id, new[] { new Cell(0, 0, "ghstreak: <github-username>") });
        return;
    }

    var stats = LoadCached(username) ?? Fetch(username);
    if (stats == null)
    {
        Emit(id, new[] { new Cell(0, 0, $"err: user '{username}' not found / rate-limited") });
        return;
    }
    SaveCache(username, stats);

    string streakCell = stats.Streak > 0 ? $"🔥 {stats.Streak}d streak" : "no streak";
    Emit(id, new[]
    {
        new Cell(0, 0, $"@{stats.Username}"),
        new Cell(0, 1, $"{stats.CommitsToday} today"),
        new Cell(0, 2, streakCell),
        new Cell(0, 3, $"{stats.CommitsLast90Days} in 90d"),
    });
}

GhStats? Fetch(string username)
{
    string url = string.Format(Endpoint, Uri.EscapeDataString(username));
    HttpResponseMessage resp;
    try { resp = http.GetAsync(url).GetAwaiter().GetResult(); }
    catch { return null; }
    using (resp)
    {
        if (!resp.IsSuccessStatusCode) return null;
        string json = resp.Content.ReadAsStringAsync().GetAwaiter().GetResult();
        using var doc = JsonDocument.Parse(json);
        if (doc.RootElement.ValueKind != JsonValueKind.Array) return null;

        // Days (UTC) → commit count from PushEvent payloads
        var dayCommits = new Dictionary<DateOnly, int>();
        int total = 0;
        foreach (var ev in doc.RootElement.EnumerateArray())
        {
            string? type = ev.TryGetProperty("type", out var tp) ? tp.GetString() : null;
            if (type != "PushEvent") continue;
            if (!ev.TryGetProperty("created_at", out var ca)) continue;
            if (!DateTimeOffset.TryParse(ca.GetString(), out var when)) continue;
            var day = DateOnly.FromDateTime(when.UtcDateTime);
            int size = 1;
            if (ev.TryGetProperty("payload", out var pl) && pl.TryGetProperty("size", out var sz))
                size = sz.GetInt32();
            dayCommits.TryGetValue(day, out int existing);
            dayCommits[day] = existing + size;
            total += size;
        }

        var today = DateOnly.FromDateTime(DateTime.UtcNow);
        int commitsToday = dayCommits.TryGetValue(today, out int t1) ? t1 : 0;

        // Streak: count back from today (or yesterday if today is empty — common case
        // because the user may not have committed yet today). Stop on first gap.
        int streak = 0;
        var probe = commitsToday > 0 ? today : today.AddDays(-1);
        while (dayCommits.ContainsKey(probe))
        {
            streak++;
            probe = probe.AddDays(-1);
        }

        return new GhStats(username, commitsToday, streak, total, DateTimeOffset.UtcNow.ToUnixTimeSeconds());
    }
}

GhStats? LoadCached(string username)
{
    string path = CachePath(username);
    if (!File.Exists(path)) return null;
    try
    {
        var s = JsonSerializer.Deserialize<GhStats>(File.ReadAllText(path));
        if (s == null) return null;
        long age = DateTimeOffset.UtcNow.ToUnixTimeSeconds() - s.FetchedAt;
        return age < CacheTtlSeconds ? s : null;
    }
    catch { return null; }
}

void SaveCache(string username, GhStats stats)
{
    string path = CachePath(username);
    Directory.CreateDirectory(Path.GetDirectoryName(path)!);
    try { File.WriteAllText(path, JsonSerializer.Serialize(stats)); }
    catch { /* best-effort */ }
}

static string CachePath(string username)
{
    string root = OperatingSystem.IsWindows()
        ? Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
        : (Environment.GetEnvironmentVariable("XDG_CACHE_HOME")
            ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".cache"));
    return Path.Combine(root, "quicksheet-ghstreak", username.ToLowerInvariant() + ".json");
}

void Emit(string id, IEnumerable<Cell> cells)
{
    Console.WriteLine(JsonSerializer.Serialize(new { type = "write", id, cells }));
    Console.Out.Flush();
}

record Cell(int r, int c, string v);
record GhStats(string Username, int CommitsToday, int Streak, int CommitsLast90Days, long FetchedAt);
