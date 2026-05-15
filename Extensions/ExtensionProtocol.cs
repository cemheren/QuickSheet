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

        [JsonConverter(typeof(CellWriteArrayConverter))]
        public CellWrite[] Cells { get; set; } = [];
    }

    /// <summary>
    /// Handles two cell formats from extensions:
    /// 1. Object format: [{r:0, c:0, v:"text"}, ...] — explicit positioning
    /// 2. Grid format:   [["a","b"], ["c","d"]]      — row-major relative to anchor
    /// </summary>
    public class CellWriteArrayConverter : JsonConverter<CellWrite[]>
    {
        public override CellWrite[] Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            if (reader.TokenType != JsonTokenType.StartArray)
                throw new JsonException("Expected array for cells");

            using var doc = JsonDocument.ParseValue(ref reader);
            var root = doc.RootElement;
            var result = new System.Collections.Generic.List<CellWrite>();

            for (int i = 0; i < root.GetArrayLength(); i++)
            {
                var element = root[i];
                if (element.ValueKind == JsonValueKind.Object)
                {
                    // Object format: {r, c, v}
                    int r = element.TryGetProperty("r", out var rp) ? rp.GetInt32() : 0;
                    int c = element.TryGetProperty("c", out var cp) ? cp.GetInt32() : 0;
                    string v = element.TryGetProperty("v", out var vp) ? vp.GetString() ?? "" : "";
                    result.Add(new CellWrite { Row = r, Col = c, Value = v });
                }
                else if (element.ValueKind == JsonValueKind.Array)
                {
                    // Grid format: each inner array is a row of string values
                    for (int j = 0; j < element.GetArrayLength(); j++)
                    {
                        string val = element[j].GetString() ?? "";
                        result.Add(new CellWrite { Row = i, Col = j, Value = val });
                    }
                }
                else if (element.ValueKind == JsonValueKind.String)
                {
                    // Flat array of strings: each element is a row with one column
                    result.Add(new CellWrite { Row = i, Col = 0, Value = element.GetString() ?? "" });
                }
            }

            return result.ToArray();
        }

        public override void Write(Utf8JsonWriter writer, CellWrite[] value, JsonSerializerOptions options)
        {
            writer.WriteStartArray();
            foreach (var cell in value)
            {
                writer.WriteStartObject();
                writer.WriteNumber("r", cell.Row);
                writer.WriteNumber("c", cell.Col);
                writer.WriteString("v", cell.Value);
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
        }
    }

    public class ErrorMessage
    {
        public string Type { get; set; } = "error";
        public string Id { get; set; } = "";
        public string Message { get; set; } = "";
    }

    public class StatusMessage
    {
        public string Type { get; set; } = "status";
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
