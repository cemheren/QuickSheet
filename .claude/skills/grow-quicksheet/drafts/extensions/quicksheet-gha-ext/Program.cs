using System.Net.Http.Headers;
using System.Text.Json;

// QuickSheet gha extension — shows GitHub Actions workflow run status.
//
// Cell usage:
//   gha: owner/repo
//   gha: owner/repo, 5              → show last 5 workflow runs (default: 8)
//   gha: owner/repo, 5, main        → filter to runs on branch 'main'
//
// Output: one row per recent workflow run. Columns: indicator, workflow name, branch, conclusion, elapsed.
// Indicators: ✓ (success), ✗ (failure), ⚠ (cancelled/skipped), ⟳ (in_progress/queued).
//
// Auth: reads GITHUB_TOKEN env var. Without it, only public repos work (60 req/hr rate limit).
// No NuGet deps. Uses System.Net.Http + System.Text.Json.
//
// Protocol: emit register on startup; handle activate messages.

var token = Environment.GetEnvironmentVariable("GITHUB_TOKEN") ?? "";
var http = new HttpClient(new SocketsHttpHandler { AllowAutoRedirect = true })
{
    Timeout = TimeSpan.FromSeconds(10),
};
http.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("quicksheet-gha-ext", "1.0.0"));
http.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/vnd.github+json"));
if (token.Length > 0)
    http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);

Console.OutputEncoding = System.Text.Encoding.UTF8;
Console.WriteLine(JsonSerializer.Serialize(new
{
    type = "register",
    prefix = "gha",
    name = "GitHub Actions",
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

    string spec = "";
    if (root.TryGetProperty("params", out var p) && p.ValueKind == JsonValueKind.Array)
    {
        var en = p.EnumerateArray();
        if (en.MoveNext()) spec = en.Current.GetString() ?? "";
    }
    spec = spec.Trim();

    if (spec.Length == 0)
    {
        Emit(id, new[] { new Cell(0, 0, "gha: <owner/repo>[, count][, branch]") });
        return;
    }

    // Parse: "owner/repo, count, branch"
    var parts = spec.Split(',', StringSplitOptions.TrimEntries);
    string repo = parts[0];
    int count = 8;
    string? branch = null;

    if (parts.Length > 1 && int.TryParse(parts[1], out int n) && n > 0 && n <= 25)
        count = n;
    if (parts.Length > 2 && parts[2].Length > 0)
        branch = parts[2];

    if (!repo.Contains('/'))
    {
        Emit(id, new[] { new Cell(0, 0, "err: expected owner/repo format") });
        return;
    }

    var runs = FetchRuns(repo, count, branch);
    if (runs == null)
    {
        Emit(id, new[] { new Cell(0, 0, "err: API request failed (check GITHUB_TOKEN?)") });
        return;
    }
    if (runs.Count == 0)
    {
        Emit(id, new[] { new Cell(0, 0, "no workflow runs found") });
        return;
    }

    var cells = new List<Cell>();
    int r = 0;
    foreach (var run in runs)
    {
        cells.Add(new Cell(r, 0, run.Indicator));
        cells.Add(new Cell(r, 1, run.WorkflowName));
        cells.Add(new Cell(r, 2, run.Branch));
        cells.Add(new Cell(r, 3, run.Conclusion));
        cells.Add(new Cell(r, 4, run.Elapsed));
        r++;
    }
    Emit(id, cells);
}

List<RunInfo>? FetchRuns(string repo, int count, string? branch)
{
    try
    {
        string url = $"https://api.github.com/repos/{repo}/actions/runs?per_page={count}";
        if (branch != null) url += $"&branch={Uri.EscapeDataString(branch)}";

        var resp = http.GetAsync(url).GetAwaiter().GetResult();
        if (!resp.IsSuccessStatusCode) return null;
        var body = resp.Content.ReadAsStringAsync().GetAwaiter().GetResult();
        using var doc = JsonDocument.Parse(body);

        var runs = new List<RunInfo>();
        if (!doc.RootElement.TryGetProperty("workflow_runs", out var arr)) return runs;

        foreach (var el in arr.EnumerateArray())
        {
            if (runs.Count >= count) break;

            string status = el.TryGetProperty("status", out var s) ? s.GetString() ?? "" : "";
            string conclusion = el.TryGetProperty("conclusion", out var c) && c.ValueKind == JsonValueKind.String
                ? c.GetString() ?? "" : status;
            string workflowName = el.TryGetProperty("name", out var wn) ? wn.GetString() ?? "" : "?";
            string branchName = el.TryGetProperty("head_branch", out var hb) ? hb.GetString() ?? "" : "";

            string elapsed = "";
            if (el.TryGetProperty("created_at", out var ca) && el.TryGetProperty("updated_at", out var ua))
            {
                if (DateTime.TryParse(ca.GetString(), out var start) &&
                    DateTime.TryParse(ua.GetString(), out var end))
                {
                    var dur = end - start;
                    elapsed = dur.TotalMinutes < 1 ? $"{dur.Seconds}s"
                        : dur.TotalHours < 1 ? $"{(int)dur.TotalMinutes}m{dur.Seconds}s"
                        : $"{(int)dur.TotalHours}h{dur.Minutes}m";
                }
            }

            string indicator = conclusion switch
            {
                "success" => "✓",
                "failure" => "✗",
                "cancelled" => "⚠",
                "skipped" => "⚠",
                "in_progress" or "queued" or "waiting" or "pending" or "requested" => "⟳",
                _ => "?",
            };

            // Truncate long workflow names
            if (workflowName.Length > 24) workflowName = workflowName[..22] + "…";
            if (branchName.Length > 20) branchName = branchName[..18] + "…";

            runs.Add(new RunInfo(indicator, workflowName, branchName, conclusion, elapsed));
        }
        return runs;
    }
    catch { return null; }
}

void Emit(string id, IEnumerable<Cell> cells)
{
    Console.WriteLine(JsonSerializer.Serialize(new { type = "write", id, cells }));
    Console.Out.Flush();
}

record Cell(int r, int c, string v);
record RunInfo(string Indicator, string WorkflowName, string Branch, string Conclusion, string Elapsed);
