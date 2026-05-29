using System.Collections.Concurrent;
using System.Text.Json;
using System.Text.Json.Serialization;

/// <summary>
/// QuickSheet PubMed Extension — looks up articles via NCBI E-utilities (free, no key required).
/// Registers the "pubmed" prefix.
/// Usage:
///   pubmed: 12345678              → lookup by PMID
///   pubmed: search kidney cancer  → search PubMed, show top results
///   pubmed: help                  → usage reference
/// </summary>
class Program
{
    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        WriteIndented = false
    };

    private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromSeconds(15) };

    static Program()
    {
        Http.DefaultRequestHeaders.UserAgent.ParseAdd(
            "quicksheet-pubmed-ext/1.0 (https://github.com/cemheren/quicksheet-pubmed-ext)");
    }

    private static readonly ConcurrentDictionary<string, List<string>> Cache = new(StringComparer.OrdinalIgnoreCase);

    static void Main()
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;
        string? line;
        while ((line = Console.ReadLine()) != null)
        {
            if (string.IsNullOrWhiteSpace(line)) continue;
            try
            {
                using var doc = JsonDocument.Parse(line);
                string? type = doc.RootElement.TryGetProperty("type", out var tp) ? tp.GetString() : null;
                switch (type)
                {
                    case "init": HandleInit(); break;
                    case "activate": HandleActivate(doc.RootElement); break;
                }
            }
            catch (Exception ex)
            {
                SendLog($"parse error: {ex.Message}");
            }
        }
    }

    static void HandleInit()
    {
        SendJson(new
        {
            type = "register",
            prefix = "pubmed",
            name = "PubMed Lookup",
            version = "1.0.0"
        });
        SendLog("PubMed extension registered with prefix 'pubmed'");
    }

    static void HandleActivate(JsonElement root)
    {
        string id = root.TryGetProperty("id", out var idProp) ? idProp.GetString() ?? "" : "";
        int gridRows = root.TryGetProperty("gridRows", out var gr) ? gr.GetInt32() : 6;

        string[] extParams = [];
        if (root.TryGetProperty("params", out var p) && p.ValueKind == JsonValueKind.Array)
            extParams = p.EnumerateArray().Select(x => x.GetString() ?? "").ToArray();

        if (extParams.Length == 0 || string.IsNullOrWhiteSpace(extParams[0]))
        {
            WriteCells(id, [["pubmed: <PMID | search query>"]]);
            return;
        }

        string query = extParams[0].Trim();

        if (query.Equals("help", StringComparison.OrdinalIgnoreCase))
        {
            WriteCells(id, [
                ["pubmed: <PMID>        — article by ID"],
                ["pubmed: search <term> — search articles"],
                ["Examples:"],
                ["  pubmed: 33782455"],
                ["  pubmed: search CRISPR therapy"],
                ["Source: NCBI E-utilities (free)"]
            ]);
            return;
        }

        try
        {
            List<string> lines;
            if (query.StartsWith("search ", StringComparison.OrdinalIgnoreCase))
            {
                string term = query[7..].Trim();
                lines = SearchArticles(term, gridRows);
            }
            else
            {
                lines = FetchByPmid(query);
            }

            if (lines.Count == 0)
            {
                WriteCells(id, [[$"No results for: {query}"]]);
                return;
            }
            while (lines.Count < gridRows) lines.Add("");
            if (lines.Count > gridRows) lines = lines.Take(gridRows).ToList();
            WriteCells(id, lines.Select(l => new[] { l }));
        }
        catch (Exception ex)
        {
            WriteCells(id, [[$"err: {ex.Message}"]]);
        }
    }

    static List<string> FetchByPmid(string pmid)
    {
        pmid = pmid.Trim().TrimStart('0');
        if (!int.TryParse(pmid, out _))
            return [$"Invalid PMID: {pmid}"];

        if (Cache.TryGetValue($"pmid:{pmid}", out var cached)) return cached;

        string url = $"https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esummary.fcgi?db=pubmed&id={pmid}&retmode=json";
        var resp = Http.GetAsync(url).GetAwaiter().GetResult();
        if (!resp.IsSuccessStatusCode)
        {
            Cache[$"pmid:{pmid}"] = [];
            return [];
        }

        string json = resp.Content.ReadAsStringAsync().GetAwaiter().GetResult();
        var lines = ParseSummary(json, pmid);
        Cache[$"pmid:{pmid}"] = lines;
        return lines;
    }

    static List<string> ParseSummary(string json, string pmid)
    {
        var lines = new List<string>();
        using var doc = JsonDocument.Parse(json);

        if (!doc.RootElement.TryGetProperty("result", out var result)) return lines;
        if (!result.TryGetProperty(pmid, out var article)) return lines;

        // Title
        if (article.TryGetProperty("title", out var titleProp))
        {
            string title = titleProp.GetString() ?? "";
            if (!string.IsNullOrEmpty(title)) lines.Add(title);
        }

        // Authors (first 3 + et al.)
        if (article.TryGetProperty("authors", out var authors) && authors.ValueKind == JsonValueKind.Array)
        {
            var names = new List<string>();
            foreach (var a in authors.EnumerateArray())
            {
                if (a.TryGetProperty("name", out var nameProp))
                    names.Add(nameProp.GetString() ?? "");
                if (names.Count >= 3) break;
            }
            if (names.Count > 0)
            {
                string authStr = string.Join(", ", names);
                if (authors.GetArrayLength() > 3) authStr += ", et al.";
                lines.Add(authStr);
            }
        }

        // Journal + year
        string journal = "";
        if (article.TryGetProperty("fulljournalname", out var jProp))
            journal = jProp.GetString() ?? "";
        else if (article.TryGetProperty("source", out var sProp))
            journal = sProp.GetString() ?? "";

        string pubdate = "";
        if (article.TryGetProperty("pubdate", out var dateProp))
            pubdate = dateProp.GetString() ?? "";

        if (!string.IsNullOrEmpty(journal) || !string.IsNullOrEmpty(pubdate))
            lines.Add($"{journal} ({pubdate})".Trim());

        // DOI
        if (article.TryGetProperty("elocationid", out var elocProp))
        {
            string eloc = elocProp.GetString() ?? "";
            if (eloc.StartsWith("doi:") || eloc.Contains("10."))
                lines.Add(eloc);
        }
        else if (article.TryGetProperty("articleids", out var aidsProp) && aidsProp.ValueKind == JsonValueKind.Array)
        {
            foreach (var aid in aidsProp.EnumerateArray())
            {
                if (aid.TryGetProperty("idtype", out var idtype) && idtype.GetString() == "doi"
                    && aid.TryGetProperty("value", out var val))
                {
                    lines.Add($"doi:{val.GetString()}");
                    break;
                }
            }
        }

        // PubMed link
        lines.Add($"https://pubmed.ncbi.nlm.nih.gov/{pmid}/");

        return lines;
    }

    static List<string> SearchArticles(string term, int gridRows)
    {
        if (Cache.TryGetValue($"search:{term}", out var cached)) return cached;

        int maxResults = Math.Min(5, gridRows / 2);
        string searchUrl = $"https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi?db=pubmed&term={Uri.EscapeDataString(term)}&retmax={maxResults}&retmode=json&sort=relevance";
        var searchResp = Http.GetAsync(searchUrl).GetAwaiter().GetResult();
        if (!searchResp.IsSuccessStatusCode) return [];

        string searchJson = searchResp.Content.ReadAsStringAsync().GetAwaiter().GetResult();
        using var searchDoc = JsonDocument.Parse(searchJson);

        if (!searchDoc.RootElement.TryGetProperty("esearchresult", out var esearch)) return [];

        int count = 0;
        if (esearch.TryGetProperty("count", out var countProp))
            int.TryParse(countProp.GetString(), out count);

        if (!esearch.TryGetProperty("idlist", out var idList) || idList.ValueKind != JsonValueKind.Array)
            return [];

        var pmids = idList.EnumerateArray().Select(x => x.GetString() ?? "").Where(x => x.Length > 0).ToList();
        if (pmids.Count == 0) return [];

        // Fetch summaries for all IDs in one call
        string ids = string.Join(",", pmids);
        string summaryUrl = $"https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esummary.fcgi?db=pubmed&id={ids}&retmode=json";
        var summaryResp = Http.GetAsync(summaryUrl).GetAwaiter().GetResult();
        if (!summaryResp.IsSuccessStatusCode) return [];

        string summaryJson = summaryResp.Content.ReadAsStringAsync().GetAwaiter().GetResult();
        using var summaryDoc = JsonDocument.Parse(summaryJson);
        if (!summaryDoc.RootElement.TryGetProperty("result", out var result)) return [];

        var lines = new List<string> { $"PubMed: \"{term}\" ({count:N0} results)" };

        foreach (var pmid in pmids)
        {
            if (!result.TryGetProperty(pmid, out var article)) continue;

            string title = "";
            if (article.TryGetProperty("title", out var tp))
                title = tp.GetString() ?? "";

            string year = "";
            if (article.TryGetProperty("pubdate", out var dp))
            {
                string pd = dp.GetString() ?? "";
                if (pd.Length >= 4) year = pd[..4];
            }

            string shortAuthors = "";
            if (article.TryGetProperty("authors", out var authProp) && authProp.ValueKind == JsonValueKind.Array)
            {
                var first = authProp.EnumerateArray().FirstOrDefault();
                if (first.ValueKind == JsonValueKind.Object && first.TryGetProperty("name", out var np))
                    shortAuthors = np.GetString() ?? "";
                if (authProp.GetArrayLength() > 1) shortAuthors += " et al.";
            }

            // Compact: one line per result
            string entry = $"PMID:{pmid}";
            if (!string.IsNullOrEmpty(shortAuthors)) entry += $" | {shortAuthors}";
            if (!string.IsNullOrEmpty(year)) entry += $" ({year})";
            lines.Add(entry);

            if (!string.IsNullOrEmpty(title))
            {
                if (title.Length > 80) title = title[..77] + "...";
                lines.Add($"  {title}");
            }
        }

        Cache[$"search:{term}"] = lines;
        return lines;
    }

    static void WriteCells(string id, IEnumerable<string[]> rows)
    {
        SendJson(new { type = "write", id, cells = rows });
    }

    static void SendJson(object obj)
    {
        Console.WriteLine(JsonSerializer.Serialize(obj, JsonOpts));
        Console.Out.Flush();
    }

    static void SendLog(string message)
    {
        SendJson(new { type = "log", message });
    }
}
