using System.Collections.Concurrent;

namespace ExcelConsole;

/// <summary>
/// Manages background HTTP fetches for "w: url" cells.
/// Results are cached for 5 minutes. Thread-safe.
/// </summary>
public sealed class WebFetchManager : IDisposable
{
    private const int CacheMinutes = 5;
    private const int MaxLineLength = 120;

    private sealed record FetchEntry(string? Content, DateTime FetchedAt, bool InProgress);

    private readonly ConcurrentDictionary<string, FetchEntry> _cache = new();
    private readonly HttpClient _http;
    private int _pendingCount;

    public WebFetchManager()
    {
        _http = new HttpClient { Timeout = TimeSpan.FromSeconds(10) };
        _http.DefaultRequestHeaders.UserAgent.ParseAdd("QuickSheet/1.0");
    }

    /// <summary>True while any URL is being fetched in the background.</summary>
    public bool HasPending => _pendingCount > 0;

    /// <summary>
    /// Returns the display text for a URL cell.
    /// Starts a background fetch on first call or after cache expiry.
    /// </summary>
    public string GetDisplay(string url)
    {
        if (_cache.TryGetValue(url, out var entry))
        {
            if (entry.InProgress) return "⏳";
            if (DateTime.Now - entry.FetchedAt < TimeSpan.FromMinutes(CacheMinutes))
                return entry.Content ?? "⚠ fetch error";
            // Cache expired — fall through to re-fetch
        }

        // Not cached / expired — kick off fetch if not already in-progress
        if (!(_cache.TryGetValue(url, out var check) && check.InProgress))
        {
            _cache[url] = new FetchEntry(null, DateTime.MinValue, InProgress: true);
            Interlocked.Increment(ref _pendingCount);
            _ = FetchAsync(url);
        }
        return "⏳";
    }

    /// <summary>Clears the cache, triggering re-fetches on next render.</summary>
    public void InvalidateAll() => _cache.Clear();

    private async Task FetchAsync(string url)
    {
        string result = "⚠ unknown error";
        try
        {
            string body = await _http.GetStringAsync(url).ConfigureAwait(false);
            // Use first non-empty line; strip leading/trailing whitespace
            string? line = body
                .Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Select(l => l.Trim('\r').Trim())
                .FirstOrDefault(l => l.Length > 0);
            result = line == null ? "(empty)"
                   : line.Length > MaxLineLength ? line[..(MaxLineLength - 1)] + "…"
                   : line;
        }
        catch (Exception ex)
        {
            string msg = ex.Message;
            result = $"⚠ {(msg.Length > 40 ? msg[..37] + "…" : msg)}";
        }
        _cache[url] = new FetchEntry(result, DateTime.Now, InProgress: false);
        Interlocked.Decrement(ref _pendingCount);
    }

    public void Dispose() => _http.Dispose();
}
