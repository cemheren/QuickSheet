using System.Text.Json;
using System.Text.Json.Serialization;

namespace ExcelConsole.Extensions;

/// <summary>
/// JSON-lines protocol messages between QuickSheet host and extension processes.
/// All messages are single-line JSON objects with a "type" discriminator.
/// </summary>
public static class ExtensionProtocol
{
    public const int ProtocolVersion = 1;

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        WriteIndented = false
    };

    public static string Serialize<T>(T message) => JsonSerializer.Serialize(message, JsonOptions);

    public static T? Deserialize<T>(string json) => JsonSerializer.Deserialize<T>(json, JsonOptions);

    /// <summary>
    /// Parses the "type" field from a JSON message without full deserialization.
    /// </summary>
    public static string? GetMessageType(string json)
    {
        try
        {
            using var doc = JsonDocument.Parse(json);
            if (doc.RootElement.TryGetProperty("type", out var typeProp))
                return typeProp.GetString();
        }
        catch { }
        return null;
    }

    // ── Host → Extension messages ────────────────────────────────────

    public class InitMessage
    {
        public string Type => "init";
        public int Version { get; set; } = ProtocolVersion;
    }

    public class ActivateMessage
    {
        public string Type => "activate";
        public string Id { get; set; } = "";
        public CellPosition Anchor { get; set; } = new();
        public string[] Params { get; set; } = [];
        public int GridCols { get; set; }
        public int GridRows { get; set; }
    }

    public class DeactivateMessage
    {
        public string Type => "deactivate";
        public string Id { get; set; } = "";
        public CellPosition Anchor { get; set; } = new();
    }

    // ── Extension → Host messages ────────────────────────────────────

    public class RegisterMessage
    {
        public string Type { get; set; } = "register";
        public string Prefix { get; set; } = "";
        public string Name { get; set; } = "";
        public string Version { get; set; } = "";
    }

    public class WriteCellsMessage
    {
        public string Type { get; set; } = "write";
        public string Id { get; set; } = "";
        public CellWrite[] Cells { get; set; } = [];
    }

    public class ErrorMessage
    {
        public string Type { get; set; } = "error";
        public string Id { get; set; } = "";
        public string Message { get; set; } = "";
    }

    public class LogMessage
    {
        public string Type { get; set; } = "log";
        public string Level { get; set; } = "info";
        public string Message { get; set; } = "";
    }

    // ── Shared types ─────────────────────────────────────────────────

    public class CellPosition
    {
        public int Row { get; set; }
        public int Col { get; set; }
    }

    public class CellWrite
    {
        [JsonPropertyName("r")]
        public int Row { get; set; }

        [JsonPropertyName("c")]
        public int Col { get; set; }

        [JsonPropertyName("v")]
        public string Value { get; set; } = "";
    }
}
