using System.Collections.Concurrent;
using System.Text.Json;
using System.Text.Json.Serialization;

/// <summary>
/// QuickSheet Books Extension — looks up books via the Open Library API.
/// Prefix: "books". Free API, no key required.
/// Usage: `books: <title or ISBN>, <cols>, <rows>`
/// Examples:
///   books: Dune, 1, 5
///   books: 978-0-13-468599-1, 1, 5
///   books: search functional programming, 2, 5
/// </summary>
class Program
{
    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        WriteIndented = false
    };

    private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromSeconds(10) };

    static Program()
    {
        Http.DefaultRequestHeaders.UserAgent.ParseAdd(
            "quicksheet-books-ext/1.0 (https://github.com/cemheren/quicksheet-books-ext)");
    }

    private static readonly ConcurrentDictionary<string, List<string[]>> Cache = new(StringComparer.OrdinalIgnoreCase);

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
            prefix = "books",
            name = "Book Lookup (Open Library)",
            version = "1.0.0"
        });
        SendLog("Books extension registered with prefix 'books'");
    }

    static void HandleActivate(JsonElement root)
    {
        string id = root.TryGetProperty("id", out var idProp) ? idProp.GetString() ?? "" : "";
        int gridCols = root.TryGetProperty("gridCols", out var gc) ? gc.GetInt32() : 1;
        int gridRows = root.TryGetProperty("gridRows", out var gr) ? gr.GetInt32() : 5;

        string[] extParams = [];
        if (root.TryGetProperty("params", out var p) && p.ValueKind == JsonValueKind.Array)
            extParams = p.EnumerateArray().Select(x => x.GetString() ?? "").ToArray();

        if (extParams.Length == 0 || string.IsNullOrWhiteSpace(extParams[0]))
        {
            WriteCells(id, [["books: <title or ISBN>, cols, rows"]]);
            return;
        }

        string query = extParams[0].Trim();
        bool isSearch = query.StartsWith("search ", StringComparison.OrdinalIgnoreCase);
        if (isSearch) query = query[7..].Trim();

        try
        {
            List<string[]> rows;
            if (isSearch || !IsIsbn(query))
            {
                rows = SearchBooks(query, gridCols, gridRows, isSearch);
            }
            else
            {
                rows = LookupIsbn(query, gridRows);
            }

            if (rows.Count == 0)
            {
                WriteCells(id, [[$"No results for: {query}"]]);
                return;
            }

            while (rows.Count < gridRows) rows.Add(Enumerable.Repeat("", gridCols).ToArray());
            if (rows.Count > gridRows) rows = rows.Take(gridRows).ToList();
            WriteCells(id, rows);
        }
        catch (Exception ex)
        {
            WriteCells(id, [[$"err: {ex.Message}"]]);
        }
    }

    static bool IsIsbn(string s)
    {
        string digits = new(s.Where(c => char.IsDigit(c) || c == 'X' || c == 'x').ToArray());
        return digits.Length == 10 || digits.Length == 13;
    }

    static List<string[]> SearchBooks(string query, int cols, int rows, bool multiResult)
    {
        string cacheKey = $"search:{query}:{cols}:{rows}";
        if (Cache.TryGetValue(cacheKey, out var cached)) return cached;

        string url = $"https://openlibrary.org/search.json?q={Uri.EscapeDataString(query)}&limit={Math.Min(rows, 10)}&fields=title,author_name,first_publish_year,publisher,number_of_pages_median,isbn";
        var resp = Http.GetAsync(url).GetAwaiter().GetResult();
        if (!resp.IsSuccessStatusCode) return [];
        string json = resp.Content.ReadAsStringAsync().GetAwaiter().GetResult();

        using var doc = JsonDocument.Parse(json);
        if (!doc.RootElement.TryGetProperty("docs", out var docs)) return [];

        var result = new List<string[]>();

        if (multiResult && cols >= 2)
        {
            // Multi-result table: title | author | year
            result.Add(PadRow(["Title", "Author", "Year"], cols));
            foreach (var book in docs.EnumerateArray())
            {
                if (result.Count >= rows) break;
                string title = book.TryGetProperty("title", out var t) ? t.GetString() ?? "" : "";
                string author = GetFirstAuthor(book);
                string year = book.TryGetProperty("first_publish_year", out var y) ? y.GetInt32().ToString() : "";
                result.Add(PadRow([title, author, year], cols));
            }
        }
        else
        {
            // Single result detail (first match)
            var first = docs.EnumerateArray().FirstOrDefault();
            if (first.ValueKind == JsonValueKind.Undefined) return [];
            result = FormatBookDetail(first, cols);
        }

        Cache[cacheKey] = result;
        return result;
    }

    static List<string[]> LookupIsbn(string isbn, int rows)
    {
        string cleanIsbn = new(isbn.Where(c => char.IsDigit(c) || c == 'X' || c == 'x').ToArray());
        string cacheKey = $"isbn:{cleanIsbn}";
        if (Cache.TryGetValue(cacheKey, out var cached)) return cached;

        // Use search API with ISBN for consistent response format
        string url = $"https://openlibrary.org/search.json?isbn={Uri.EscapeDataString(cleanIsbn)}&limit=1&fields=title,author_name,first_publish_year,publisher,number_of_pages_median,isbn";
        var resp = Http.GetAsync(url).GetAwaiter().GetResult();
        if (!resp.IsSuccessStatusCode) return [];
        string json = resp.Content.ReadAsStringAsync().GetAwaiter().GetResult();

        using var doc = JsonDocument.Parse(json);
        if (!doc.RootElement.TryGetProperty("docs", out var docs)) return [];
        var first = docs.EnumerateArray().FirstOrDefault();
        if (first.ValueKind == JsonValueKind.Undefined) return [];

        var result = FormatBookDetail(first, 1);
        Cache[cacheKey] = result;
        return result;
    }

    static List<string[]> FormatBookDetail(JsonElement book, int cols)
    {
        var lines = new List<string[]>();
        string title = book.TryGetProperty("title", out var t) ? t.GetString() ?? "" : "Unknown";
        string author = GetFirstAuthor(book);
        string year = book.TryGetProperty("first_publish_year", out var y) ? y.GetInt32().ToString() : "?";
        string publisher = "";
        if (book.TryGetProperty("publisher", out var pub) && pub.ValueKind == JsonValueKind.Array && pub.GetArrayLength() > 0)
            publisher = pub.EnumerateArray().First().GetString() ?? "";
        string pages = book.TryGetProperty("number_of_pages_median", out var pg) ? $"{pg.GetInt32()} pp" : "";

        lines.Add(PadRow([$"📖 {title}"], cols));
        if (!string.IsNullOrEmpty(author)) lines.Add(PadRow([$"by {author}"], cols));
        if (year != "?") lines.Add(PadRow([$"Published: {year}"], cols));
        if (!string.IsNullOrEmpty(publisher)) lines.Add(PadRow([$"Publisher: {publisher}"], cols));
        if (!string.IsNullOrEmpty(pages)) lines.Add(PadRow([$"Pages: {pages}"], cols));

        return lines;
    }

    static string GetFirstAuthor(JsonElement book)
    {
        if (book.TryGetProperty("author_name", out var authors) && authors.ValueKind == JsonValueKind.Array && authors.GetArrayLength() > 0)
        {
            var first = authors.EnumerateArray().First().GetString() ?? "";
            if (authors.GetArrayLength() > 1) return $"{first} et al.";
            return first;
        }
        return "";
    }

    static string[] PadRow(string[] cells, int cols)
    {
        if (cells.Length >= cols) return cells.Take(cols).ToArray();
        return [.. cells, .. Enumerable.Repeat("", cols - cells.Length)];
    }

    static void WriteCells(string id, IEnumerable<string[]> rows)
    {
        SendJson(new { type = "write", id, cells = rows });
    }

    static void SendJson(object obj)
    {
        Console.WriteLine(JsonSerializer.Serialize(obj, JsonOpts));
    }

    static void SendLog(string message)
    {
        SendJson(new { type = "log", message });
    }
}
