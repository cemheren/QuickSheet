using System.Text;
using System.Text.RegularExpressions;

namespace ExcelConsole.Terminal;

/// <summary>
/// Fallback IVtScreen used when libghostty-vt is not available: strips ANSI escapes with a regex
/// and keeps a simple line buffer (with carriage-return overwrite for progress bars).
/// </summary>
internal sealed class LegacyVtScreen : IVtScreen
{
    private static readonly Regex AnsiPattern = new(
        @"\x1b(?:\[[0-9;?]*[A-Za-z@~]|\][^\x07]*\x07|\([A-Z0-9]|[A-Z=<>7-9])|[\x00-\x08\x0e-\x1f]",
        RegexOptions.Compiled);

    private static readonly Regex IncompleteEscape = new(
        @"\x1b(?:\[[0-9;?]*|\][^\x07]*|\(?)$", RegexOptions.Compiled);
    private const int MaxPartialChars = 64 * 1024;

    private readonly object _lock = new();
    private readonly Decoder _decoder = Encoding.UTF8.GetDecoder();
    private readonly List<string> _lines = new();
    private readonly StringBuilder _current = new();
    private readonly int _max;
    private bool _pendingCr;
    private bool _dirty = true;
    private string? _cache;

    public LegacyVtScreen(int cols, int rows, int maxScrollbackLines)
    {
        _max = Math.Max(1, maxScrollbackLines);
    }

    public bool IsDirty { get { lock (_lock) return _dirty; } }

    public void Write(ReadOnlySpan<byte> data)
    {
        if (data.IsEmpty) return;
        lock (_lock)
        {
            var chars = new char[_decoder.GetCharCount(data, false)];
            int n = _decoder.GetChars(data, chars, false);
            for (int i = 0; i < n; i++)
            {
                char ch = chars[i];
                if (ch == '\n')
                {
                    string l = Strip(_current.ToString());
                    if (l.Length > 0) _lines.Add(l);
                    _current.Clear();
                    _pendingCr = false;
                }
                else if (ch == '\r')
                {
                    _pendingCr = true;
                }
                else
                {
                    if (_pendingCr) { _current.Clear(); _pendingCr = false; }
                    _current.Append(ch);
                    if (_current.Length > MaxPartialChars)
                        _current.Remove(0, _current.Length - MaxPartialChars);
                }
            }
            if (_lines.Count > _max)
                _lines.RemoveRange(0, _lines.Count - _max);
            _dirty = true;
        }
    }

    public void Resize(int cols, int rows)
    {
        lock (_lock) _dirty = true;
    }

    public string GetText()
    {
        lock (_lock)
        {
            if (!_dirty && _cache != null) return _cache;
            var sb = new StringBuilder();
            foreach (var l in _lines)
            {
                if (sb.Length > 0) sb.Append('\n');
                sb.Append(l.TrimEnd());
            }
            string partial = Strip(IncompleteEscape.Replace(_current.ToString(), ""));
            if (partial.Length > 0)
            {
                if (sb.Length > 0) sb.Append('\n');
                sb.Append(partial.TrimEnd());
            }
            _cache = sb.ToString().TrimEnd('\n', ' ');
            _dirty = false;
            return _cache;
        }
    }

    private static string Strip(string s) => AnsiPattern.Replace(s, "");

    public void Dispose() { }
}
