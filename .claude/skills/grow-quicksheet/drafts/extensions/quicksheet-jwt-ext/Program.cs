using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

/// <summary>
/// QuickSheet JWT Extension — decodes JWT tokens inline.
/// Usage: jwt: eyJhbGciOiJIUzI1NiIs..., 1, 5
/// Displays decoded header and payload claims in the grid.
/// </summary>
class Program
{
    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        WriteIndented = false
    };

    private static readonly JsonSerializerOptions PrettyOpts = new()
    {
        WriteIndented = false
    };

    static void Main()
    {
        Console.OutputEncoding = Encoding.UTF8;
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
            prefix = "jwt",
            name = "JWT Decoder",
            version = "1.0.0"
        });
        SendLog("JWT extension registered with prefix 'jwt'");
    }

    static void HandleActivate(JsonElement root)
    {
        string id = root.TryGetProperty("id", out var idProp) ? idProp.GetString() ?? "" : "";
        int gridRows = root.TryGetProperty("gridRows", out var gr) ? gr.GetInt32() : 8;

        string[] extParams = [];
        if (root.TryGetProperty("params", out var p) && p.ValueKind == JsonValueKind.Array)
            extParams = p.EnumerateArray().Select(x => x.GetString() ?? "").ToArray();

        if (extParams.Length == 0 || string.IsNullOrWhiteSpace(extParams[0]))
        {
            WriteCells(id, new[] { new[] { "jwt: <token>" } });
            return;
        }

        string token = extParams[0].Trim();
        var parts = token.Split('.');

        if (parts.Length < 2)
        {
            WriteCells(id, new[] { new[] { "⚠ Invalid JWT (need 2+ parts)" } });
            return;
        }

        try
        {
            var header = DecodeBase64Url(parts[0]);
            var payload = DecodeBase64Url(parts[1]);

            using var headerDoc = JsonDocument.Parse(header);
            using var payloadDoc = JsonDocument.Parse(payload);

            var rows = new List<string[]>();
            rows.Add(new[] { "── Header ──" });

            foreach (var prop in headerDoc.RootElement.EnumerateObject())
            {
                rows.Add(new[] { $"  {prop.Name}: {FormatValue(prop.Value)}" });
            }

            rows.Add(new[] { "── Payload ──" });

            foreach (var prop in payloadDoc.RootElement.EnumerateObject())
            {
                string val = FormatValue(prop.Value);
                // Decode exp/iat/nbf as human-readable dates
                if ((prop.Name == "exp" || prop.Name == "iat" || prop.Name == "nbf")
                    && prop.Value.ValueKind == JsonValueKind.Number
                    && prop.Value.TryGetInt64(out long epoch))
                {
                    var dt = DateTimeOffset.FromUnixTimeSeconds(epoch).UtcDateTime;
                    val = $"{epoch} ({dt:yyyy-MM-dd HH:mm} UTC)";
                }
                rows.Add(new[] { $"  {prop.Name}: {val}" });
            }

            // Show signature status
            rows.Add(new[] { parts.Length >= 3 ? "── Sig: present ──" : "── Sig: none ──" });

            // Trim or pad to gridRows
            while (rows.Count < gridRows) rows.Add(new[] { "" });
            if (rows.Count > gridRows) rows = rows.Take(gridRows).ToList();

            WriteCells(id, rows);
        }
        catch (Exception ex)
        {
            WriteCells(id, new[] { new[] { $"⚠ Decode error: {ex.Message}" } });
        }
    }

    static string DecodeBase64Url(string input)
    {
        string padded = input.Replace('-', '+').Replace('_', '/');
        switch (padded.Length % 4)
        {
            case 2: padded += "=="; break;
            case 3: padded += "="; break;
        }
        byte[] bytes = Convert.FromBase64String(padded);
        return Encoding.UTF8.GetString(bytes);
    }

    static string FormatValue(JsonElement el)
    {
        return el.ValueKind switch
        {
            JsonValueKind.String => el.GetString() ?? "",
            JsonValueKind.Number => el.GetRawText(),
            JsonValueKind.True => "true",
            JsonValueKind.False => "false",
            JsonValueKind.Null => "null",
            JsonValueKind.Array => $"[{el.GetArrayLength()} items]",
            JsonValueKind.Object => "{...}",
            _ => el.GetRawText()
        };
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
