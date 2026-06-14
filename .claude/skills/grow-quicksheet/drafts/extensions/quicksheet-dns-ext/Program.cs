using System.Net;
using System.Net.Sockets;
using System.Text.Json;
using System.Text.Json.Serialization;

/// <summary>
/// QuickSheet DNS Extension — resolves hostnames via System.Net.Dns.
/// Prefix: "dns". Usage: dns: example.com[, reverse]
/// Shows resolved IPv4/IPv6 addresses. With "reverse" param, does PTR lookup.
/// </summary>
class Program
{
    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        WriteIndented = false
    };

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
            prefix = "dns",
            name = "DNS Resolver",
            version = "1.0.0"
        });
        SendLog("DNS extension registered with prefix 'dns'");
    }

    static void HandleActivate(JsonElement root)
    {
        string id = root.TryGetProperty("id", out var idProp) ? idProp.GetString() ?? "" : "";
        int gridRows = root.TryGetProperty("gridRows", out var gr) ? gr.GetInt32() : 5;

        string[] extParams = [];
        if (root.TryGetProperty("params", out var p) && p.ValueKind == JsonValueKind.Array)
            extParams = p.EnumerateArray().Select(x => x.GetString() ?? "").ToArray();

        if (extParams.Length == 0 || string.IsNullOrWhiteSpace(extParams[0]))
        {
            WriteCells(id, new[] { new[] { "dns: <hostname>" } });
            return;
        }

        string host = extParams[0].Trim();
        bool reverse = extParams.Length > 1 &&
            extParams[1].Trim().Equals("reverse", StringComparison.OrdinalIgnoreCase);

        try
        {
            if (reverse)
                ResolveReverse(id, host, gridRows);
            else
                ResolveForward(id, host, gridRows);
        }
        catch (Exception ex)
        {
            string msg = ex.InnerException?.Message ?? ex.Message;
            WriteCells(id, new[] { new[] { $"err: {msg}" } });
        }
    }

    static void ResolveForward(string id, string host, int gridRows)
    {
        var entry = Dns.GetHostEntry(host);
        var rows = new List<string[]> { new[] { $"⟨dns⟩ {host}" } };

        var ipv4 = entry.AddressList
            .Where(a => a.AddressFamily == AddressFamily.InterNetwork)
            .ToList();
        var ipv6 = entry.AddressList
            .Where(a => a.AddressFamily == AddressFamily.InterNetworkV6)
            .ToList();

        if (ipv4.Count > 0)
            rows.Add(new[] { $"A    {string.Join(", ", ipv4.Select(a => a.ToString()))}" });
        if (ipv6.Count > 0)
            rows.Add(new[] { $"AAAA {string.Join(", ", ipv6.Select(a => a.ToString()))}" });

        if (entry.Aliases.Length > 0)
            rows.Add(new[] { $"CNAME {string.Join(", ", entry.Aliases)}" });

        if (ipv4.Count == 0 && ipv6.Count == 0)
            rows.Add(new[] { "(no addresses)" });

        // Pad or trim to grid size
        while (rows.Count < gridRows) rows.Add(new[] { "" });
        if (rows.Count > gridRows) rows = rows.Take(gridRows).ToList();
        WriteCells(id, rows);
    }

    static void ResolveReverse(string id, string ipStr, int gridRows)
    {
        if (!IPAddress.TryParse(ipStr, out var ip))
        {
            WriteCells(id, new[] { new[] { $"err: not a valid IP: {ipStr}" } });
            return;
        }

        var entry = Dns.GetHostEntry(ip);
        var rows = new List<string[]>
        {
            new[] { $"⟨ptr⟩ {ipStr}" },
            new[] { $"→ {entry.HostName}" }
        };

        if (entry.Aliases.Length > 0)
            rows.Add(new[] { $"aliases: {string.Join(", ", entry.Aliases)}" });

        while (rows.Count < gridRows) rows.Add(new[] { "" });
        if (rows.Count > gridRows) rows = rows.Take(gridRows).ToList();
        WriteCells(id, rows);
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
