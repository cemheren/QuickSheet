using System.Text.Json;
using System.Xml;

// QuickSheet RSS extension — displays live headlines from RSS/Atom feeds on your desktop.
//
// Cell usage:
//   rss: https://hnrss.org/newest?points=50
//   rss: https://feeds.arstechnica.com/arstechnica/index
//   rss: /path/to/feeds.txt
//
// feeds.txt layout (one URL per line, # for comments):
//   https://hnrss.org/newest?points=50
//   https://feeds.arstechnica.com/arstechnica/index
//
// Output: one row per item. Columns: title, link, published date.
// Shows up to gridRows items (default 10). Multiple feeds are merged by date.
//
// Protocol: emit register on startup, then handle activate messages.

var http = new HttpClient(new SocketsHttpHandler { AllowAutoRedirect = true })
{
    Timeout = TimeSpan.FromSeconds(10),
};
http.DefaultRequestHeaders.UserAgent.ParseAdd("QuickSheet-RSS/1.0");

Console.OutputEncoding = System.Text.Encoding.UTF8;
Console.WriteLine(JsonSerializer.Serialize(new
{
    type = "register",
    prefix = "rss",
    name = "RSS Feed",
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

    if (spec.Length == 0)
    {
        Emit(id, new[] { new Cell(0, 0, "rss: <feed-url> or path/to/feeds.txt") });
        return;
    }

    var urls = ParseFeedUrls(spec);
    if (urls.Count == 0)
    {
        Emit(id, new[] { new Cell(0, 0, "No valid feed URLs found") });
        return;
    }

    var allItems = new List<FeedItem>();
    foreach (var url in urls)
    {
        try
        {
            var items = FetchFeed(url);
            allItems.AddRange(items);
        }
        catch
        {
            allItems.Add(new FeedItem($"⚠ Error fetching: {Truncate(url, 40)}", "", null));
        }
    }

    // Sort by date descending, nulls last
    allItems.Sort((a, b) =>
    {
        if (a.Published == null && b.Published == null) return 0;
        if (a.Published == null) return 1;
        if (b.Published == null) return -1;
        return b.Published.Value.CompareTo(a.Published.Value);
    });

    var cells = new List<Cell>();
    int r = 0;
    foreach (var item in allItems.Take(gridRows))
    {
        cells.Add(new Cell(r, 0, Truncate(item.Title, 60)));
        cells.Add(new Cell(r, 1, item.Published?.ToString("MM-dd HH:mm") ?? ""));
        cells.Add(new Cell(r, 2, Truncate(item.Link, 50)));
        r++;
    }

    if (cells.Count == 0)
        cells.Add(new Cell(0, 0, "No items found in feed"));

    Emit(id, cells);
}

List<string> ParseFeedUrls(string spec)
{
    var urls = new List<string>();

    // If it looks like a file path, read URLs from it
    if (!spec.StartsWith("http://", StringComparison.OrdinalIgnoreCase) &&
        !spec.StartsWith("https://", StringComparison.OrdinalIgnoreCase) &&
        File.Exists(spec))
    {
        foreach (var raw in File.ReadAllLines(spec))
        {
            var ln = raw.Trim();
            if (ln.Length == 0 || ln.StartsWith("#")) continue;
            if (ln.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                ln.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
                urls.Add(ln);
        }
        return urls;
    }

    // Otherwise treat as a single URL (or comma-separated URLs)
    foreach (var part in spec.Split(' ', StringSplitOptions.RemoveEmptyEntries))
    {
        var u = part.Trim().TrimEnd(',');
        if (u.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
            u.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
            urls.Add(u);
    }

    // If splitting by space didn't work, try the whole thing as one URL
    if (urls.Count == 0 && spec.StartsWith("http", StringComparison.OrdinalIgnoreCase))
        urls.Add(spec);

    return urls;
}

List<FeedItem> FetchFeed(string url)
{
    string xml = http.GetStringAsync(url).GetAwaiter().GetResult();
    var items = new List<FeedItem>();

    var xdoc = new XmlDocument();
    xdoc.LoadXml(xml);

    // Try RSS 2.0 first
    var rssItems = xdoc.GetElementsByTagName("item");
    if (rssItems.Count > 0)
    {
        foreach (XmlNode node in rssItems)
        {
            string title = node["title"]?.InnerText?.Trim() ?? "(no title)";
            string link = node["link"]?.InnerText?.Trim() ?? "";
            DateTimeOffset? pub = TryParseDate(node["pubDate"]?.InnerText);
            items.Add(new FeedItem(title, link, pub));
        }
        return items;
    }

    // Try Atom
    var nsMgr = new XmlNamespaceManager(xdoc.NameTable);
    nsMgr.AddNamespace("atom", "http://www.w3.org/2005/Atom");
    var entries = xdoc.SelectNodes("//atom:entry", nsMgr);
    if (entries != null && entries.Count > 0)
    {
        foreach (XmlNode node in entries)
        {
            string title = node.SelectSingleNode("atom:title", nsMgr)?.InnerText?.Trim() ?? "(no title)";
            string link = node.SelectSingleNode("atom:link/@href", nsMgr)?.Value?.Trim() ?? "";
            string? dateStr = node.SelectSingleNode("atom:published", nsMgr)?.InnerText
                ?? node.SelectSingleNode("atom:updated", nsMgr)?.InnerText;
            DateTimeOffset? pub = TryParseDate(dateStr);
            items.Add(new FeedItem(title, link, pub));
        }
        return items;
    }

    // Fallback: try plain <entry> without namespace
    var plainEntries = xdoc.GetElementsByTagName("entry");
    if (plainEntries.Count > 0)
    {
        foreach (XmlNode node in plainEntries)
        {
            string title = node["title"]?.InnerText?.Trim() ?? "(no title)";
            var linkNode = node["link"];
            string link = linkNode?.GetAttribute("href") ?? linkNode?.InnerText?.Trim() ?? "";
            string? dateStr = node["published"]?.InnerText ?? node["updated"]?.InnerText;
            DateTimeOffset? pub = TryParseDate(dateStr);
            items.Add(new FeedItem(title, link, pub));
        }
    }

    return items;
}

static DateTimeOffset? TryParseDate(string? s)
{
    if (string.IsNullOrWhiteSpace(s)) return null;
    if (DateTimeOffset.TryParse(s, out var dt)) return dt;
    return null;
}

static string Truncate(string s, int max) =>
    s.Length <= max ? s : s[..(max - 1)] + "…";

void Emit(string id, IEnumerable<Cell> cells)
{
    Console.WriteLine(JsonSerializer.Serialize(new { type = "write", id, cells }));
    Console.Out.Flush();
}

record Cell(int r, int c, string v);
record FeedItem(string Title, string Link, DateTimeOffset? Published);
