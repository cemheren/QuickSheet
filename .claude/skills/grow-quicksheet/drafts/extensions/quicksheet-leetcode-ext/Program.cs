using System.Text;
using System.Text.Json;

// QuickSheet leetcode extension — shows a CS student's LeetCode solved count + streak
// on the wallpaper. Pure read-only against LeetCode's public GraphQL endpoint;
// only public profile data; no auth.
//
// Cell usage:
//   leetcode: <username>
//
// Output: one row, four columns: total solved, easy, medium, hard, current streak.
//
// Caches the response for 1 hour to be polite to LeetCode (and useful for an
// always-on wallpaper). Cache lives under XDG cache dir on Linux, %LOCALAPPDATA%
// on Windows. Cache is per-username.

const string Endpoint = "https://leetcode.com/graphql";
const int CacheTtlSeconds = 3600;

var http = new HttpClient { Timeout = TimeSpan.FromSeconds(8) };
http.DefaultRequestHeaders.UserAgent.ParseAdd("quicksheet-leetcode-ext/1.0 (+https://github.com/cemheren)");

Console.OutputEncoding = Encoding.UTF8;
Console.WriteLine(JsonSerializer.Serialize(new
{
    type = "register",
    prefix = "leetcode",
    name = "LeetCode Stats",
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
        Emit(id, new[] { new Cell(0, 0, "leetcode: <username>") });
        return;
    }

    var stats = LoadCached(username) ?? Fetch(username);
    if (stats == null)
    {
        Emit(id, new[] { new Cell(0, 0, $"err: user '{username}' not found") });
        return;
    }
    SaveCache(username, stats);

    Emit(id, new[]
    {
        new Cell(0, 0, $"@{stats.Username}"),
        new Cell(0, 1, $"{stats.Total} solved"),
        new Cell(0, 2, $"E {stats.Easy} · M {stats.Medium} · H {stats.Hard}"),
        new Cell(0, 3, stats.Streak > 0 ? $"🔥 {stats.Streak}d" : "no streak"),
    });
}

LeetStats? Fetch(string username)
{
    const string query = @"query userStats($u: String!) {
      matchedUser(username: $u) {
        username
        submitStats { acSubmissionNum { difficulty count } }
        userCalendar { streak }
      }
    }";
    var body = new
    {
        operationName = "userStats",
        query,
        variables = new { u = username },
    };
    var req = new HttpRequestMessage(HttpMethod.Post, Endpoint)
    {
        Content = new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json"),
    };
    HttpResponseMessage resp;
    try { resp = http.SendAsync(req).GetAwaiter().GetResult(); }
    catch { return null; }
    using (resp)
    {
        if (!resp.IsSuccessStatusCode) return null;
        string json = resp.Content.ReadAsStringAsync().GetAwaiter().GetResult();
        using var doc = JsonDocument.Parse(json);
        if (!doc.RootElement.TryGetProperty("data", out var data)) return null;
        if (!data.TryGetProperty("matchedUser", out var user) || user.ValueKind == JsonValueKind.Null) return null;

        string uname = user.GetProperty("username").GetString() ?? username;
        int total = 0, easy = 0, medium = 0, hard = 0;
        foreach (var diff in user.GetProperty("submitStats").GetProperty("acSubmissionNum").EnumerateArray())
        {
            int count = diff.GetProperty("count").GetInt32();
            switch (diff.GetProperty("difficulty").GetString())
            {
                case "All":    total = count; break;
                case "Easy":   easy = count; break;
                case "Medium": medium = count; break;
                case "Hard":   hard = count; break;
            }
        }
        int streak = 0;
        if (user.TryGetProperty("userCalendar", out var cal)
            && cal.ValueKind == JsonValueKind.Object
            && cal.TryGetProperty("streak", out var s))
            streak = s.GetInt32();

        return new LeetStats(uname, total, easy, medium, hard, streak, DateTimeOffset.UtcNow.ToUnixTimeSeconds());
    }
}

LeetStats? LoadCached(string username)
{
    string path = CachePath(username);
    if (!File.Exists(path)) return null;
    try
    {
        var s = JsonSerializer.Deserialize<LeetStats>(File.ReadAllText(path));
        if (s == null) return null;
        long age = DateTimeOffset.UtcNow.ToUnixTimeSeconds() - s.FetchedAt;
        return age < CacheTtlSeconds ? s : null;
    }
    catch { return null; }
}

void SaveCache(string username, LeetStats stats)
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
    return Path.Combine(root, "quicksheet-leetcode", username.ToLowerInvariant() + ".json");
}

void Emit(string id, IEnumerable<Cell> cells)
{
    Console.WriteLine(JsonSerializer.Serialize(new { type = "write", id, cells }));
    Console.Out.Flush();
}

record Cell(int r, int c, string v);
record LeetStats(string Username, int Total, int Easy, int Medium, int Hard, int Streak, long FetchedAt);
