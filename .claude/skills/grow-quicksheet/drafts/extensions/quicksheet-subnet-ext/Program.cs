using System.Net;
using System.Text.Json;
using System.Text.Json.Serialization;

/// <summary>
/// QuickSheet Subnet Extension — CIDR/subnet calculator.
/// Given a CIDR notation (e.g. 192.168.1.0/24), displays network address, broadcast,
/// usable host range, host count, and netmask. Useful for SRE desktop dashboards.
/// Usage: `subnet: 10.0.0.0/16`
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
            prefix = "subnet",
            name = "Subnet Calculator",
            version = "1.0.0"
        });
        SendLog("Subnet extension registered with prefix 'subnet'");
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
            WriteCells(id, new[] { new[] { "subnet: <cidr>" }, new[] { "e.g. 10.0.0.0/24" } }, gridRows);
            return;
        }

        string input = extParams[0].Trim();

        try
        {
            var info = SubnetCalc.Parse(input);
            var rows = new List<string[]>
            {
                new[] { $"🌐 {info.Cidr}" },
                new[] { $"Network:   {info.Network}" },
                new[] { $"Broadcast: {info.Broadcast}" },
                new[] { $"Range:     {info.FirstHost} – {info.LastHost}" },
                new[] { $"Hosts:     {info.UsableHosts:N0}" },
                new[] { $"Netmask:   {info.Netmask}" }
            };
            WriteCells(id, rows, gridRows);
        }
        catch (Exception ex)
        {
            WriteCells(id, new[] { new[] { $"err: {ex.Message}" } }, gridRows);
        }
    }

    static void WriteCells(string id, IEnumerable<string[]> rows, int gridRows)
    {
        var list = rows.ToList();
        while (list.Count < gridRows) list.Add(new[] { "" });
        if (list.Count > gridRows) list = list.Take(gridRows).ToList();
        SendJson(new { type = "write", id, cells = list });
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

static class SubnetCalc
{
    public record SubnetInfo(
        string Cidr,
        string Network,
        string Broadcast,
        string Netmask,
        string FirstHost,
        string LastHost,
        long UsableHosts
    );

    public static SubnetInfo Parse(string cidr)
    {
        // Support bare IP (assume /32) or IP/prefix
        int prefixLen;
        string ipPart;

        int slashIdx = cidr.IndexOf('/');
        if (slashIdx < 0)
        {
            ipPart = cidr;
            prefixLen = 32;
        }
        else
        {
            ipPart = cidr[..slashIdx];
            if (!int.TryParse(cidr[(slashIdx + 1)..], out prefixLen) || prefixLen < 0 || prefixLen > 32)
                throw new ArgumentException($"Invalid prefix length: {cidr[(slashIdx + 1)..]}");
        }

        if (!IPAddress.TryParse(ipPart, out var ip) || ip.AddressFamily != System.Net.Sockets.AddressFamily.InterNetwork)
            throw new ArgumentException($"Invalid IPv4 address: {ipPart}");

        uint ipNum = IpToUint(ip);
        uint mask = prefixLen == 0 ? 0 : 0xFFFFFFFF << (32 - prefixLen);
        uint network = ipNum & mask;
        uint broadcast = network | ~mask;

        uint firstHost, lastHost;
        long usableHosts;

        if (prefixLen == 32)
        {
            firstHost = lastHost = network;
            usableHosts = 1;
        }
        else if (prefixLen == 31)
        {
            // Point-to-point link (RFC 3021)
            firstHost = network;
            lastHost = broadcast;
            usableHosts = 2;
        }
        else
        {
            firstHost = network + 1;
            lastHost = broadcast - 1;
            usableHosts = (long)(broadcast - network - 1);
        }

        return new SubnetInfo(
            Cidr: $"{UintToIp(network)}/{prefixLen}",
            Network: UintToIp(network),
            Broadcast: UintToIp(broadcast),
            Netmask: UintToIp(mask),
            FirstHost: UintToIp(firstHost),
            LastHost: UintToIp(lastHost),
            UsableHosts: usableHosts
        );
    }

    private static uint IpToUint(IPAddress ip)
    {
        byte[] bytes = ip.GetAddressBytes();
        return (uint)(bytes[0] << 24 | bytes[1] << 16 | bytes[2] << 8 | bytes[3]);
    }

    private static string UintToIp(uint ip)
    {
        return $"{(ip >> 24) & 0xFF}.{(ip >> 16) & 0xFF}.{(ip >> 8) & 0xFF}.{ip & 0xFF}";
    }
}
