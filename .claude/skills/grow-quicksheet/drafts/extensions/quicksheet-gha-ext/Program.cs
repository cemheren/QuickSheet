using System.Net.Http.Headers;
using System.Text.Json;

// QuickSheet gha extension — shows GitHub Actions workflow run statuses.
//
// Cell usage:
//   gha: owner/repo
//   gha: owner/repo, 5          (limit to 5 most recent runs)
//
// Output: one row per workflow run. Columns: workflow, branch, status, conclusion, elapsed.
// Indicators: ✓ (success), ✗ (failure), ⚠ (cancelled), ◉ (in_progress), ○ (queued).
//
// Auth: uses GITHUB_TOKEN env var if set (higher rate limits). Works unauthenticated
// for public repos (60 req/hr).
//
// Protocol: emit register on startup, then handle activate messages.

var http = new HttpClient(new SocketsHttpHandler { AllowAutoRedirect = true })
{
    Timeout = TimeSpan.FromSeconds(10),
};
http.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("quicksheet-gha-ext", "1.0.0"));
http.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/vnd.github+json"));

var token = Environment.GetEnvironmentVariable("GITHUB_TOKEN");
if (!string.IsNullOrEmpty(token))
    http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);

Console.OutputEncoding = System.Text.Encoding.UTF8;
Console.WriteLine(JsonSerializer.Serialize(new
{
    type = "register",
    prefix = "gha",
    name = "GitHub Actions Status",
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
    int gridRows = root.TryGetProperty("gridRows", out var gr) ? gr.GetInt32() : 10;

    string spec = "";
    int limit = 8;
    if (root.TryGetProperty("params", out var p) && p.ValueKind == JsonValueKind.Array)
    {
        var en = p.EnumerateArray();
        if (en.MoveNext()) spec = en.Current.GetString()?.Trim() ?? "";
        if (en.MoveNext() && int.TryParse(en.Current.GetString()?.Trim(), out var lim))
            limit = Math.Clamp(lim, 1, 25);
    }

    if (string.IsNullOrWhiteSpace(spec) || !spec.Contains('/'))
    {
        Emit(id, new[] { new Cell(0, 0, "gha: <owner/repo>[, count]") });
        return;
    }

    var runs = FetchRuns(spec, Math.Min(limit, gridRows));
    if (runs == null)
    {
        Emit(id, new[] { new Cell(0, 0, "⚠ API error"), new Cell(0, 1, spec) });
        return;
    }
    if (runs.Count == 0)
    {
        Emit(id, new[] { new Cell(0, 0, "No runs"), new Cell(0, 1, spec) });
        return;
    }

    var cells = new List<Cell>();
    int r = 0;
    foreach (var run in runs)
    {
        if (r >= gridRows) break;
        cells.Add(new Cell(r, 0, run.indicator));
        cells.Add(new Cell(r, 1, Truncate(run.workflow, 18)));
        cells.Add(new Cell(r, 2, Truncate(run.branch, 14)));
        cells.Add(new Cell(r, 3, run.elapsed));
        r++;
    }
    Emit(id, cells);
}

List<RunInfo>? FetchRuns(string repo, int count)
{
    try
    {
        var url = $"https://api.github.com/repos/{repo}/actions/runs?per_page={count}&status=completed";
        var resp = http.GetAsync(url).GetAwaiter().GetResult();
        if (!resp.IsSuccessStatusCode)
        {
            // Try including in_progress runs too
            url = $"https://api.github.com/repos/{repo}/actions/runs?per_page={count}";
            resp = http.GetAsync(url).GetAwaiter().GetResult();
            if (!resp.IsSuccessStatusCode) return null;
        }

        var json = resp.Content.ReadAsStringAsync().GetAwaiter().GetResult();
        using var doc = JsonDocument.Parse(json);
        var runs = new List<RunInfo>();

        if (!doc.RootElement.TryGetProperty("workflow_runs", out var arr)) return runs;

        foreach (var el in arr.EnumerateArray())
        {
            string workflow = el.TryGetProperty("name", out var n) ? n.GetString() ?? "?" : "?";
            string branch = el.TryGetProperty("head_branch", out var b) ? b.GetString() ?? "?" : "?";
            string status = el.TryGetProperty("status", out var s) ? s.GetString() ?? "" : "";
            string conclusion = el.TryGetProperty("conclusion", out var c) ?
                c.GetString() ?? "" : "";

            string elapsed = "";
            if (el.TryGetProperty("created_at", out var ca) &&
                el.TryGetProperty("updated_at", out var ua))
            {
                if (DateTime.TryParse(ca.GetString(), out var start) &&
                    DateTime.TryParse(ua.GetString(), out var end))
                {
                    var dur = end - start;
                    elapsed = dur.TotalMinutes < 1 ? $"{(int)dur.TotalSeconds}s"
                            : dur.TotalHours < 1 ? $"{(int)dur.TotalMinutes}m"
                            : $"{dur.TotalHours:F1}h";
                }
            }

            string indicator = (status, conclusion) switch
            {
                ("completed", "success") => "✓",
                ("completed", "failure") => "✗",
                ("completed", "cancelled") => "⚠",
                ("completed", "skipped") => "○",
                ("in_progress", _) => "◉",
                ("queued", _) => "○",
                _ => "?",
            };

            runs.Add(new RunInfo(indicator, workflow, branch, elapsed));
        }
        return runs;
    }
    catch { return null; }
}

static string Truncate(string s, int max) =>
    s.Length <= max ? s : s[..(max - 1)] + "…";

void Emit(string id, IEnumerable<Cell> cells)
{
    Console.WriteLine(JsonSerializer.Serialize(new { type = "write", id, cells }));
    Console.Out.Flush();
}

record Cell(int r, int c, string v);
record RunInfo(string indicator, string workflow, string branch, string elapsed);
