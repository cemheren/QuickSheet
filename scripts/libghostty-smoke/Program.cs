using System.Text;
using ExcelConsole.Terminal;

// Smoke test: load a libghostty-vt build via QuickSheet's real P/Invoke bindings.
// Usage: QUICKSHEET_LIBGHOSTTY=/abs/path/to/lib dotnet run -c Release

var libPath = Environment.GetEnvironmentVariable("QUICKSHEET_LIBGHOSTTY");
if (string.IsNullOrEmpty(libPath))
{
    Console.Error.WriteLine("QUICKSHEET_LIBGHOSTTY must be set to the absolute path of the library under test.");
    return 2;
}

int failures = 0;
void Check(string name, bool ok, string? detail = null)
{
    Console.WriteLine($"{(ok ? "PASS" : "FAIL")}: {name}" + (!ok && detail != null ? $"  ({detail})" : ""));
    if (!ok) failures++;
}
static string Show(string s) => s.Replace("\n", "\\n").Replace("\r", "\\r").Replace("\x1b", "\\e");

Console.WriteLine($"Library: {libPath}");
if (!GhosttyNative.IsAvailable)
{
    Console.Error.WriteLine($"FAIL: libghostty-vt could not be loaded: {GhosttyNative.LoadError}");
    return 1;
}
Console.WriteLine($"libghostty-vt version: {GhosttyNative.Version}");
Check("version is exactly 1.0.0", GhosttyNative.Version == "1.0.0", GhosttyNative.Version);
{
    var cmp = OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
    string loaded = GhosttyNative.LoadedPath is { } lp ? Path.GetFullPath(lp) : "(null)";
    Check("loaded the library named by QUICKSHEET_LIBGHOSTTY",
        string.Equals(loaded, Path.GetFullPath(libPath), cmp), $"loaded {loaded}");
}

static void W(IVtScreen s, string text) => s.Write(Encoding.UTF8.GetBytes(text));
static void WChunks(IVtScreen s, byte[] bytes, params int[] cuts)
{
    int prev = 0;
    foreach (var c in cuts.Append(bytes.Length))
    {
        s.Write(bytes.AsSpan(prev, c - prev));
        prev = c;
    }
}

try
{
    using (var s = VtScreen.Create(40, 5, 100))
    {
        Check("VtScreen.Create returns GhosttyVtScreen", s is GhosttyVtScreen, s.GetType().Name);

        // Split mid-escape-sequence and mid-UTF-8 ('é' = C3 A9).
        var b1 = Encoding.UTF8.GetBytes("\x1b[31mred\x1b[0m é\r\n");
        int esc = Array.IndexOf(b1, (byte)'[') + 1;     // inside "\x1b[31m"
        int e = Array.IndexOf(b1, (byte)0xC3) + 1;       // between C3 and A9
        WChunks(s, b1, esc, esc + 2, e);
        var b2 = Encoding.UTF8.GetBytes("progress 10%\rprogress 99%\r\n");
        WChunks(s, b2, 5, 13, 20);
        var b3 = Encoding.UTF8.GetBytes("to be erased\x1b[2K\rkept\r\n");
        WChunks(s, b3, 14, 16);

        var text = s.GetText();
        Check("chunked writes, colors, CR overwrite, erase-line", text == "red é\nprogress 99%\nkept", Show(text));
        Check("IsDirty false after GetText", !s.IsDirty);
        W(s, "x");
        Check("IsDirty true after Write", s.IsDirty);

        W(s, "\x1b[2J\x1b[Hfresh");
        text = s.GetText();
        Check("clear screen + home", text.EndsWith("fresh"), Show(text));

        s.Resize(10, 3);
        s.Resize(40, 5);
        Check("Resize down and back", true);
        _ = s.GetText();

        s.Dispose();
        s.Write("after dispose"u8);
        s.Resize(5, 5);
        Check("Write/Resize after Dispose is a no-op", true);
        // 'using' disposes again: double dispose must be safe.
    }

    using (var s = VtScreen.Create(20, 3, 1000))
    {
        for (int i = 0; i < 50; i++) W(s, $"line {i}\r\n");
        var text = s.GetText();
        Check("scrollback keeps early and late lines",
            text.Contains("line 0\n") && text.Contains("line 49"), Show(text.Length > 80 ? text[..80] + "..." : text));
    }
}
catch (Exception ex)
{
    Check("no unexpected exception", false, ex.ToString());
}

Console.WriteLine(failures == 0 ? "ALL PASS" : $"{failures} check(s) FAILED");
return failures == 0 ? 0 : 1;
