using System.Diagnostics;
using System.Text.Json;

// QuickSheet health extension — fills a grid of up/down dots for a list of HTTP services.
//
// Cell usage:
//   health: plex=https://plex.lan,pi-hole=http://pi.hole/admin, 4, 5
//   health: /path/to/services.csv, 4, 10
//
// services.csv layout (header optional, columns: name,url):
//   plex,https://plex.lan
//   pi-hole,http://pi.hole/admin
//
// Output: one row per service. Columns: name, indicator, status, latency_ms.
// Indicators: ✓ (2xx/3xx), ⚠ (4xx), ✗ (5xx or unreachable). Latency capped at 8s.
//
// Protocol: emit register on startup, then handle activate messages.

var http = new HttpClient(new SocketsHttpHandler { AllowAutoRedirect = false })
{
    Timeout = TimeSpan.FromSeconds(8),
};

Console.OutputEncoding = System.Text.Encoding.UTF8;
Console.WriteLine(JsonSerializer.Serialize(new
{
    type = "register",
    prefix = "health",
    name = "Service Health",
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
    if (root.TryGetProperty("params", out var p) && p.ValueKind == JsonValueKind.Array)
    {
        var en = p.EnumerateArray();
        if (en.MoveNext()) spec = en.Current.GetString() ?? "";
    }
    spec = spec.Trim();

    var services = ParseServices(spec);
    if (services.Count == 0)
    {
        Emit(id, new[] { new Cell(0, 0, "health: <name=url,...> or path/to/services.csv") });
        return;
    }

    var cells = new List<Cell>();
    int r = 0;
    foreach (var svc in services)
    {
        if (r >= gridRows) break;
        var (code, ms, indicator) = Probe(svc.url);
        cells.Add(new Cell(r, 0, svc.name));
        cells.Add(new Cell(r, 1, indicator));
        cells.Add(new Cell(r, 2, code > 0 ? code.ToString() : "—"));
        cells.Add(new Cell(r, 3, $"{ms} ms"));
        r++;
    }
    Emit(id, cells);
}

static List<(string name, string url)> ParseServices(string spec)
{
    var list = new List<(string, string)>();
    if (spec.Length == 0) return list;

    if (File.Exists(spec))
    {
        foreach (var raw in File.ReadAllLines(spec))
        {
            var ln = raw.Trim();
            if (ln.Length == 0 || ln.StartsWith("#")) continue;
            int comma = ln.IndexOf(',');
            if (comma <= 0) continue;
            string name = ln[..comma].Trim();
            string url = ln[(comma + 1)..].Trim();
            if (name.Equals("name", StringComparison.OrdinalIgnoreCase)) continue; // header row
            if (name.Length > 0 && url.Length > 0) list.Add((name, NormalizeUrl(url)));
        }
        return list;
    }

    foreach (var part in spec.Split(',', StringSplitOptions.RemoveEmptyEntries))
    {
        var pair = part.Trim();
        int eq = pair.IndexOf('=');
        string name, url;
        if (eq > 0) { name = pair[..eq].Trim(); url = pair[(eq + 1)..].Trim(); }
        else { name = pair; url = pair; }
        if (name.Length > 0 && url.Length > 0) list.Add((name, NormalizeUrl(url)));
    }
    return list;
}

static string NormalizeUrl(string u) =>
    (u.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
     u.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
        ? u : "https://" + u;

(int code, long ms, string indicator) Probe(string url)
{
    var sw = Stopwatch.StartNew();
    try
    {
        using var req = new HttpRequestMessage(HttpMethod.Head, url);
        HttpResponseMessage resp;
        try { resp = http.SendAsync(req).GetAwaiter().GetResult(); }
        catch
        {
            using var req2 = new HttpRequestMessage(HttpMethod.Get, url);
            resp = http.SendAsync(req2, HttpCompletionOption.ResponseHeadersRead)
                       .GetAwaiter().GetResult();
        }
        sw.Stop();
        int code = (int)resp.StatusCode;
        resp.Dispose();
        string indicator = code switch
        {
            >= 200 and < 400 => "✓",
            >= 400 and < 500 => "⚠",
            >= 500           => "✗",
            _                => "?",
        };
        return (code, sw.ElapsedMilliseconds, indicator);
    }
    catch
    {
        sw.Stop();
        return (-1, sw.ElapsedMilliseconds, "✗");
    }
}

void Emit(string id, IEnumerable<Cell> cells)
{
    Console.WriteLine(JsonSerializer.Serialize(new { type = "write", id, cells }));
    Console.Out.Flush();
}

record Cell(int r, int c, string v);
