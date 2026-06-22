using System.Text;
using System.Text.Json;

// quicksheet-gomod — QuickSheet extension that shows Go module info
// (latest version, publish date, number of published versions) from the
// public Go module proxy. Reads JSON-lines from stdin, writes JSON-lines to
// stdout. Zero NuGet dependencies; uses only the BCL HttpClient.

var http = new HttpClient();
http.DefaultRequestHeaders.Add("Accept", "application/json");
http.DefaultRequestHeaders.Add("User-Agent", "quicksheet-gomod/1.0");
http.Timeout = TimeSpan.FromSeconds(10);

var cache = new Dictionary<string, (List<(int r, int c, string v)> cells, DateTime exp)>(StringComparer.OrdinalIgnoreCase);

string? line;
string activationId = "";
while ((line = Console.ReadLine()) != null)
{
    if (string.IsNullOrWhiteSpace(line)) continue;
    JsonElement msg;
    try { msg = JsonDocument.Parse(line).RootElement; } catch { continue; }

    var type = msg.TryGetProperty("type", out var t) ? t.GetString() : null;

    if (type == "init")
    {
        Console.WriteLine(JsonSerializer.Serialize(new
        {
            type = "register",
            prefix = "gomod:",
            name = "go module info",
            description = "Show the latest version and publish date of a Go module from the public Go module proxy.",
            version = "1.0.0",
            width = 2,
            height = 5
        }));
        continue;
    }

    if (type == "activate")
    {
        activationId = msg.TryGetProperty("id", out var idEl) ? idEl.GetString() ?? "" : "";
        var rawValue = msg.TryGetProperty("value", out var vEl) ? vEl.GetString() ?? "" : "";
        var param = rawValue.StartsWith("gomod:", StringComparison.OrdinalIgnoreCase)
            ? rawValue["gomod:".Length..].Trim()
            : rawValue.Trim();

        if (string.IsNullOrWhiteSpace(param))
        {
            WriteStatus("Usage: gomod: github.com/gin-gonic/gin  or  gomod: github.com/gin-gonic/gin,gorm.io/gorm");
            continue;
        }

        var modules = param.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        var capturedId = activationId;

        var task = Task.Run(async () =>
        {
            List<(int r, int c, string v)> cells;
            if (modules.Length == 1)
            {
                cells = await FetchSingleAsync(modules[0]);
            }
            else
            {
                cells = new List<(int r, int c, string v)>
                {
                    (0, 0, "Module"), (0, 1, "Version"), (0, 2, "Released")
                };
                for (int i = 0; i < modules.Length && i < 20; i++)
                {
                    var info = await FetchInfoAsync(modules[i]);
                    cells.Add((i + 1, 0, ShortName(modules[i])));
                    cells.Add((i + 1, 1, info.version));
                    cells.Add((i + 1, 2, info.released));
                }
            }
            var cellObjs = cells.Select(x => new { r = x.r, c = x.c, v = x.v }).ToArray();
            var output = JsonSerializer.Serialize(new { type = "write", id = capturedId, cells = cellObjs });
            Console.WriteLine(output);
        });
        task.Wait();
        continue;
    }
}

void WriteStatus(string message) =>
    Console.WriteLine(JsonSerializer.Serialize(new { type = "status", id = activationId, message }));

async Task<List<(int r, int c, string v)>> FetchSingleAsync(string mod)
{
    if (cache.TryGetValue(mod, out var hit) && hit.exp > DateTime.UtcNow)
        return hit.cells;

    var info = await FetchInfoAsync(mod);
    var versions = await FetchVersionCountAsync(mod);

    var cells = new List<(int r, int c, string v)>
    {
        (0, 0, "Module"),   (0, 1, ShortName(mod)),
        (1, 0, "Version"),  (1, 1, info.version),
        (2, 0, "Released"), (2, 1, info.released),
    };
    if (versions > 0)
    {
        cells.Add((3, 0, "Versions"));
        cells.Add((3, 1, versions.ToString()));
    }
    cells.Add((4, 0, mod));

    cache[mod] = (cells, DateTime.UtcNow.AddMinutes(30));
    return cells;
}

async Task<(string version, string released)> FetchInfoAsync(string mod)
{
    try
    {
        var url = $"https://proxy.golang.org/{CaseEncode(mod.Trim('/'))}/@latest";
        var json = await http.GetStringAsync(url);
        var doc = JsonDocument.Parse(json).RootElement;

        var version = doc.TryGetProperty("Version", out var ver) ? ver.GetString() ?? "?" : "?";

        string released = "";
        if (doc.TryGetProperty("Time", out var tm) && DateTime.TryParse(tm.GetString(), out var dt))
            released = dt.ToString("yyyy-MM-dd");

        return (version, released);
    }
    catch
    {
        return ("not found", "");
    }
}

async Task<int> FetchVersionCountAsync(string mod)
{
    try
    {
        var url = $"https://proxy.golang.org/{CaseEncode(mod.Trim('/'))}/@v/list";
        var text = await http.GetStringAsync(url);
        return text.Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).Length;
    }
    catch { return 0; }
}

// Go module proxy case-encoding: to stay safe on case-insensitive file systems,
// every uppercase letter in the module path is replaced with "!" + lowercase.
// See https://go.dev/ref/mod#goproxy-protocol.
static string CaseEncode(string s)
{
    var sb = new StringBuilder(s.Length + 4);
    foreach (var ch in s)
    {
        if (ch >= 'A' && ch <= 'Z')
        {
            sb.Append('!');
            sb.Append(char.ToLowerInvariant(ch));
        }
        else
        {
            sb.Append(ch);
        }
    }
    return sb.ToString();
}

// Display the last one or two path segments to keep the cell compact,
// e.g. github.com/gin-gonic/gin → gin-gonic/gin.
static string ShortName(string mod)
{
    var parts = mod.Trim('/').Split('/', StringSplitOptions.RemoveEmptyEntries);
    if (parts.Length <= 1) return mod;
    if (parts.Length == 2) return string.Join('/', parts);
    return string.Join('/', parts[^2], parts[^1]);
}
