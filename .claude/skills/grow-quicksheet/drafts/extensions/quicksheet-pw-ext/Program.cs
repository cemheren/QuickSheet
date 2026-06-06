using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Serialization;

/// <summary>
/// QuickSheet Password Extension — generates secure random passwords.
/// Registers the "pw" prefix. Zero network, cryptographically secure randomness.
/// Usage:
///   pw: 16             → 16-char password (letters+digits+symbols)
///   pw: 24 alpha       → 24-char letters+digits only
///   pw: 32 hex         → 32-char hex string
///   pw: 4 words        → 4-word passphrase (diceware-style)
///   pw: pin 6          → 6-digit PIN
///   pw:                → default 20-char password
/// </summary>
class Program
{
    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        WriteIndented = false
    };

    private const string AlphaNum = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
    private const string Symbols = "!@#$%^&*()-_=+[]{}|;:,.<>?";
    private const string Full = AlphaNum + Symbols;
    private const string HexChars = "0123456789abcdef";

    // Compact word list for passphrases (EFF short list subset, 256 words)
    private static readonly string[] Words =
    [
        "acid", "acorn", "acre", "aged", "agent", "agile", "aging", "agree",
        "ahead", "aide", "ajar", "alarm", "album", "alert", "alias", "alibi",
        "alien", "align", "alley", "allot", "allow", "aloft", "alone", "amaze",
        "amino", "ample", "amuse", "angel", "anger", "angle", "ankle", "annex",
        "anvil", "apart", "apple", "april", "apron", "aqua", "arena", "argue",
        "arise", "armor", "aroma", "array", "arrow", "arson", "asset", "atlas",
        "atom", "attic", "audio", "avert", "avid", "avoid", "awake", "award",
        "bacon", "badge", "bagel", "baker", "balmy", "baron", "basin", "batch",
        "beach", "beard", "beast", "begin", "being", "below", "bench", "berry",
        "bison", "blade", "blame", "blank", "blast", "blaze", "bleak", "bleed",
        "blend", "bless", "blimp", "blind", "bliss", "block", "bloom", "blown",
        "bluff", "blunt", "blurt", "board", "bogey", "boil", "bold", "bolt",
        "bonus", "booth", "bound", "bowed", "boxer", "brake", "brand", "brave",
        "bread", "break", "brick", "brief", "bring", "brink", "brisk", "broad",
        "broil", "brook", "broth", "brush", "budge", "buggy", "build", "bulge",
        "bunch", "bunny", "burst", "cabin", "cable", "camel", "candy", "cargo",
        "carol", "carry", "carve", "catch", "cause", "cedar", "chain", "chair",
        "chalk", "chant", "charm", "chart", "chase", "cheap", "check", "chess",
        "chief", "child", "chill", "chirp", "chop", "chunk", "cider", "cigar",
        "cinch", "civic", "civil", "claim", "clamp", "clasp", "clash", "class",
        "claw", "clay", "clean", "clear", "clerk", "click", "cliff", "climb",
        "cling", "cloak", "clock", "clone", "cloth", "cloud", "clown", "coach",
        "coast", "cobra", "cocoa", "comet", "comic", "coral", "couch", "count",
        "cover", "craft", "crane", "crash", "crawl", "crazy", "creek", "crest",
        "crime", "crisp", "cross", "crowd", "crush", "cubic", "curve", "cycle",
        "daily", "dance", "darts", "dealt", "debug", "decal", "decay", "decoy",
        "defer", "delta", "demon", "dense", "depot", "depth", "derby", "deter",
        "detox", "diary", "digit", "dimly", "diner", "disco", "ditch", "dodge",
        "doing", "donor", "donut", "doubt", "dough", "dowel", "draft", "drain",
        "drape", "drawn", "dream", "dress", "dried", "drift", "drill", "drink",
        "drive", "drone", "drown", "drums", "drunk", "dryer", "dully", "dummy",
        "dusty", "dwarf", "dwell", "dying", "eager", "early", "earth", "easel"
    ];

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
            prefix = "pw",
            name = "Password Generator",
            version = "1.0.0"
        });
        SendLog("Password extension registered with prefix 'pw'");
    }

    static void HandleActivate(JsonElement root)
    {
        string id = root.TryGetProperty("id", out var idProp) ? idProp.GetString() ?? "" : "";
        int gridRows = root.TryGetProperty("gridRows", out var gr) ? gr.GetInt32() : 3;

        string[] extParams = [];
        if (root.TryGetProperty("params", out var p) && p.ValueKind == JsonValueKind.Array)
            extParams = p.EnumerateArray().Select(x => x.GetString() ?? "").ToArray();

        string input = extParams.Length > 0 ? extParams[0].Trim() : "";
        var parts = input.Split(' ', StringSplitOptions.RemoveEmptyEntries);

        try
        {
            var rows = new List<string[]>();
            string mode = DetectMode(parts, out int length);

            switch (mode)
            {
                case "full":
                    rows.Add(new[] { "🔑 " + GenerateFromCharset(Full, length) });
                    rows.Add(new[] { $"{length} chars (letters+digits+symbols)" });
                    break;
                case "alpha":
                    rows.Add(new[] { "🔑 " + GenerateFromCharset(AlphaNum, length) });
                    rows.Add(new[] { $"{length} chars (alphanumeric)" });
                    break;
                case "hex":
                    rows.Add(new[] { "🔑 " + GenerateFromCharset(HexChars, length) });
                    rows.Add(new[] { $"{length} chars (hex)" });
                    break;
                case "pin":
                    rows.Add(new[] { "🔑 " + GenerateFromCharset("0123456789", length) });
                    rows.Add(new[] { $"{length}-digit PIN" });
                    break;
                case "words":
                    var phrase = GeneratePassphrase(length);
                    rows.Add(new[] { "🔑 " + phrase });
                    rows.Add(new[] { $"{length}-word passphrase" });
                    break;
            }

            // Strength indicator
            int bits = EstimateEntropy(mode, length);
            string strength = bits >= 128 ? "🟢 Very strong"
                : bits >= 80 ? "🟢 Strong"
                : bits >= 60 ? "🟡 Good"
                : bits >= 40 ? "🟡 Fair"
                : "🔴 Weak";
            rows.Add(new[] { $"{strength} (~{bits} bits)" });

            while (rows.Count < gridRows) rows.Add(new[] { "" });
            if (rows.Count > gridRows) rows = rows.Take(gridRows).ToList();
            WriteCells(id, rows);
        }
        catch (Exception ex)
        {
            WriteCells(id, new[] { new[] { $"err: {ex.Message}" } });
        }
    }

    static string DetectMode(string[] parts, out int length)
    {
        length = 20; // default
        string mode = "full";

        foreach (var part in parts)
        {
            if (int.TryParse(part, out int n) && n > 0 && n <= 256)
                length = n;
            else if (part.Equals("alpha", StringComparison.OrdinalIgnoreCase)
                  || part.Equals("alnum", StringComparison.OrdinalIgnoreCase))
                mode = "alpha";
            else if (part.Equals("hex", StringComparison.OrdinalIgnoreCase))
                mode = "hex";
            else if (part.Equals("pin", StringComparison.OrdinalIgnoreCase))
            {
                mode = "pin";
                if (length == 20) length = 6; // default PIN length
            }
            else if (part.Equals("words", StringComparison.OrdinalIgnoreCase)
                  || part.Equals("passphrase", StringComparison.OrdinalIgnoreCase))
            {
                mode = "words";
                if (length == 20) length = 5; // default word count
            }
        }

        return mode;
    }

    static string GenerateFromCharset(string charset, int length)
    {
        var result = new char[length];
        for (int i = 0; i < length; i++)
            result[i] = charset[RandomNumberGenerator.GetInt32(charset.Length)];
        return new string(result);
    }

    static string GeneratePassphrase(int wordCount)
    {
        var selected = new string[wordCount];
        for (int i = 0; i < wordCount; i++)
            selected[i] = Words[RandomNumberGenerator.GetInt32(Words.Length)];
        return string.Join("-", selected);
    }

    static int EstimateEntropy(string mode, int length)
    {
        double bitsPerUnit = mode switch
        {
            "full" => Math.Log2(Full.Length),       // ~6.5
            "alpha" => Math.Log2(AlphaNum.Length),  // ~5.95
            "hex" => 4.0,
            "pin" => Math.Log2(10),                 // ~3.32
            "words" => Math.Log2(Words.Length),     // ~8.0
            _ => 6.0
        };
        return (int)(bitsPerUnit * length);
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
