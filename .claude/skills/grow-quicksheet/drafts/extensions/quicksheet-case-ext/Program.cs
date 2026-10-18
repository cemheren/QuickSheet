using System.Net.Http.Headers;
using System.Text.Json;

// QuickSheet case extension — US case law lookup via CourtListener's free API.
//
// Cell usage:
//   case: Roe v Wade
//   case: 410 U.S. 113
//   case: "unreasonable search" fourth amendment
//
// Output: up to gridRows results. Columns: case name, citation, court, year, snippet.
// API: https://www.courtlistener.com/api/rest/v4/search/?type=o&q=<query>
// No auth required for basic search (rate-limited to ~100 req/hr unauthenticated).

var http = new HttpClient();
http.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("QuickSheetCaseExt", "1.0"));
http.Timeout = TimeSpan.FromSeconds(10);

Console.OutputEncoding = System.Text.Encoding.UTF8;
Console.WriteLine(JsonSerializer.Serialize(new
{
    type = "register",
    prefix = "case",
    name = "Case Law",
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
    int gridRows = root.TryGetProperty("gridRows", out var gr) ? gr.GetInt32() : 5;

    string query = "";
    if (root.TryGetProperty("params", out var p) && p.ValueKind == JsonValueKind.Array)
    {
        var en = p.EnumerateArray();
        if (en.MoveNext()) query = en.Current.GetString() ?? "";
    }
    query = query.Trim();

    if (string.IsNullOrEmpty(query))
    {
        Emit(id, new[] { new Cell(0, 0, "case: <citation or keywords>") });
        return;
    }

    var results = Search(query, gridRows);
    if (results.Count == 0)
    {
        Emit(id, new[] { new Cell(0, 0, "No results"), new Cell(0, 1, query) });
        return;
    }

    var cells = new List<Cell>();
    for (int r = 0; r < results.Count && r < gridRows; r++)
    {
        var c = results[r];
        cells.Add(new Cell(r, 0, c.caseName));
        cells.Add(new Cell(r, 1, c.citation));
        cells.Add(new Cell(r, 2, c.court));
        cells.Add(new Cell(r, 3, c.year));
        cells.Add(new Cell(r, 4, c.snippet));
    }
    Emit(id, cells);
}

List<CaseResult> Search(string query, int limit)
{
    var results = new List<CaseResult>();
    try
    {
        string encoded = Uri.EscapeDataString(query);
        string url = $"https://www.courtlistener.com/api/rest/v4/search/?type=o&q={encoded}&page_size={Math.Min(limit, 20)}";
        var response = http.GetAsync(url).GetAwaiter().GetResult();
        if (!response.IsSuccessStatusCode) return results;

        string json = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
        using var doc = JsonDocument.Parse(json);

        if (!doc.RootElement.TryGetProperty("results", out var arr)) return results;

        foreach (var item in arr.EnumerateArray())
        {
            string caseName = item.TryGetProperty("caseName", out var cn)
                ? cn.GetString() ?? "" : "";
            string citation = item.TryGetProperty("citation", out var cit)
                ? FirstCitation(cit) : "";
            string court = item.TryGetProperty("court_citation_string", out var ct)
                ? ct.GetString() ?? "" : "";
            string year = item.TryGetProperty("dateFiled", out var df)
                ? ExtractYear(df.GetString()) : "";
            string snippet = item.TryGetProperty("snippet", out var sn)
                ? StripHtml(sn.GetString() ?? "") : "";

            if (caseName.Length == 0 && citation.Length == 0) continue;
            results.Add(new CaseResult(
                Truncate(caseName, 40),
                Truncate(citation, 25),
                Truncate(court, 15),
                year,
                Truncate(snippet, 60)
            ));
        }
    }
    catch { /* network error — return empty */ }
    return results;
}

static string FirstCitation(JsonElement el)
{
    if (el.ValueKind == JsonValueKind.Array)
    {
        foreach (var c in el.EnumerateArray())
        {
            string? s = c.GetString();
            if (!string.IsNullOrEmpty(s)) return s;
        }
    }
    if (el.ValueKind == JsonValueKind.String) return el.GetString() ?? "";
    return "";
}

static string ExtractYear(string? date)
{
    if (date == null || date.Length < 4) return "";
    return date[..4];
}

static string StripHtml(string s)
{
    // Remove <mark>, </mark>, <em>, </em>, and other common HTML tags from snippets
    var clean = System.Text.RegularExpressions.Regex.Replace(s, "<[^>]+>", "");
    clean = clean.Replace("&amp;", "&").Replace("&lt;", "<").Replace("&gt;", ">");
    clean = clean.Replace("&quot;", "\"").Replace("&#39;", "'");
    clean = clean.Replace("\n", " ").Replace("\r", "");
    // Collapse whitespace
    clean = System.Text.RegularExpressions.Regex.Replace(clean, @"\s+", " ").Trim();
    return clean;
}

static string Truncate(string s, int max) =>
    s.Length <= max ? s : s[..(max - 1)] + "…";

void Emit(string id, IEnumerable<Cell> cells)
{
    Console.WriteLine(JsonSerializer.Serialize(new { type = "write", id, cells }));
    Console.Out.Flush();
}

record Cell(int r, int c, string v);
record CaseResult(string caseName, string citation, string court, string year, string snippet);
