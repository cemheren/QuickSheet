using System.Reflection;
using System.Runtime.InteropServices;

namespace ExcelConsole.Terminal;

/// <summary>
/// Raw P/Invoke bindings for libghostty-vt (Ghostty's VT emulator C library, stable C ABI v1.0, tag vt-1.0.0).
/// libghostty-vt is NOT a NuGet package. `dotnet build` downloads it from this repository's CI-built release
/// (pinned SHA-256 in libghostty.props), or it comes from a local native/&lt;rid&gt;/ build made by
/// scripts/build-libghostty.sh / .ps1, or from the path in QUICKSHEET_LIBGHOSTTY; it ends up next to the
/// executable. If it is absent, QuickSheet falls back to LegacyVtScreen.
/// </summary>
internal static class GhosttyNative
{
    public const string Lib = "ghostty-vt";

    public const int OptScrollbackMaxBytes = 28;
    public const int OptScrollbackMaxLines = 29;
    public const int OptResizePullScrollback = 41;
    public const int FormatterOptEmit = 1;
    public const int FormatterOptUnwrap = 2;
    public const int FormatterOptTrim = 3;
    public const int FormatPlain = 0;

    // GhosttyResult: SUCCESS=0, OUT_OF_MEMORY=-1, INVALID_VALUE=-2, OUT_OF_SPACE=-3, UNSUPPORTED=-8
    public const int Success = 0;

    [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
    public static extern int ghostty_terminal_new(IntPtr allocator, out IntPtr terminal, ushort cols, ushort rows);

    [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
    public static extern void ghostty_terminal_free(IntPtr terminal);

    [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
    public static extern void ghostty_terminal_vt_write(IntPtr terminal, ref byte data, nuint len);

    [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
    public static extern int ghostty_terminal_resize(IntPtr terminal, ushort cols, ushort rows, uint cellWidthPx, uint cellHeightPx);

    [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
    public static extern int ghostty_terminal_set(IntPtr terminal, int option, IntPtr value);

    [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
    public static extern int ghostty_formatter_terminal_new(IntPtr allocator, out IntPtr formatter, IntPtr terminal);

    [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
    public static extern int ghostty_formatter_set(IntPtr formatter, int option, IntPtr value);

    [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
    public static extern int ghostty_formatter_format_alloc(IntPtr formatter, IntPtr allocator, out IntPtr outPtr, out nuint outLen);

    [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
    public static extern void ghostty_formatter_free(IntPtr formatter);

    [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
    public static extern void ghostty_free(IntPtr allocator, IntPtr ptr, nuint len);

    [DllImport(Lib, CallingConvention = CallingConvention.Cdecl, ExactSpelling = true)]
    public static extern IntPtr ghostty_version();

    // ---- helpers: pass a value by pointer ----

    public static int TerminalSetNuint(IntPtr terminal, int option, nuint value)
    {
        IntPtr p = Marshal.AllocHGlobal(IntPtr.Size);
        try
        {
            Marshal.WriteIntPtr(p, unchecked((IntPtr)(nint)value));
            return ghostty_terminal_set(terminal, option, p);
        }
        finally { Marshal.FreeHGlobal(p); }
    }

    /// <summary>Sets a C `bool` terminal option (1 byte).</summary>
    public static int TerminalSetBool(IntPtr terminal, int option, bool value)
    {
        IntPtr p = Marshal.AllocHGlobal(1);
        try
        {
            Marshal.WriteByte(p, value ? (byte)1 : (byte)0);
            return ghostty_terminal_set(terminal, option, p);
        }
        finally { Marshal.FreeHGlobal(p); }
    }

    public static int FormatterSetInt(IntPtr formatter, int option, int value)
    {
        IntPtr p = Marshal.AllocHGlobal(sizeof(int));
        try
        {
            Marshal.WriteInt32(p, value);
            return ghostty_formatter_set(formatter, option, p);
        }
        finally { Marshal.FreeHGlobal(p); }
    }

    /// <summary>Sets a C `bool` option (1 byte).</summary>
    public static int FormatterSetBool(IntPtr formatter, int option, bool value)
    {
        IntPtr p = Marshal.AllocHGlobal(1);
        try
        {
            Marshal.WriteByte(p, value ? (byte)1 : (byte)0);
            return ghostty_formatter_set(formatter, option, p);
        }
        finally { Marshal.FreeHGlobal(p); }
    }

    // ---- library loading ----

    private static IntPtr _handle;
    private static string? _version;
    private static string? _loadError;
    private static int _resolverRegistered;

    private static readonly Lazy<bool> _available = new(Init, LazyThreadSafetyMode.ExecutionAndPublication);

    public static bool IsAvailable => _available.Value;
    public static string? Version { get { _ = _available.Value; return _version; } }
    /// <summary>Path or name of the library that was actually loaded (null if none).</summary>
    public static string? LoadedPath { get; private set; }
    public static string? LoadError { get { _ = _available.Value; return _loadError; } }

    private static bool Init()
    {
        try
        {
            if (Interlocked.Exchange(ref _resolverRegistered, 1) == 0)
            {
                try { NativeLibrary.SetDllImportResolver(typeof(GhosttyNative).Assembly, Resolver); }
                catch (InvalidOperationException) { /* another resolver already registered */ }
            }

            // Resolve explicitly too, so we still work if someone else owns the assembly's resolver.
            if (_handle == IntPtr.Zero)
                _handle = ProbeLoad();
            if (_handle == IntPtr.Zero)
            {
                _loadError = "native library not found (set QUICKSHEET_LIBGHOSTTY or place it next to the executable)";
                return false;
            }

            IntPtr vp = ghostty_version();
            string? v = vp == IntPtr.Zero ? null : Marshal.PtrToStringUTF8(vp);
            _version = v;
            if (string.IsNullOrEmpty(v))
            {
                _loadError = "ghostty_version() returned no version string";
                return false;
            }
            int dot = v.IndexOf('.');
            string major = dot >= 0 ? v.Substring(0, dot) : v;
            if (!int.TryParse(major, out int maj) || maj != 1)
            {
                _loadError = $"unsupported libghostty-vt version '{v}' (need 1.x)";
                return false;
            }
            return true;
        }
        catch (Exception ex)
        {
            _loadError = $"{ex.GetType().Name}: {ex.Message}";
            return false;
        }
    }

    private static IntPtr Resolver(string libraryName, Assembly assembly, DllImportSearchPath? searchPath)
    {
        if (libraryName != Lib) return IntPtr.Zero;
        if (_handle == IntPtr.Zero) _handle = ProbeLoad();
        return _handle;
    }

    private static IntPtr ProbeLoad()
    {
        string? env = Environment.GetEnvironmentVariable("QUICKSHEET_LIBGHOSTTY");
        // Only absolute paths: a relative name would be searched via CWD/PATH.
        if (!string.IsNullOrWhiteSpace(env) && Path.IsPathFullyQualified(env) && NativeLibrary.TryLoad(env, out IntPtr h1))
        {
            LoadedPath = env;
            return h1;
        }

        string[] names;
        if (OperatingSystem.IsWindows()) names = new[] { "ghostty-vt.dll" };
        else if (OperatingSystem.IsMacOS()) names = new[] { "libghostty-vt.1.dylib", "libghostty-vt.dylib" };
        else names = new[] { "libghostty-vt.so.1", "libghostty-vt.so" };

        string baseDir = AppContext.BaseDirectory;
        foreach (var n in names)
        {
            string full = Path.Combine(baseDir, n);
            if (File.Exists(full) && NativeLibrary.TryLoad(full, out IntPtr h2))
            {
                LoadedPath = full;
                return h2;
            }
        }
        // Bare-name system probe is skipped on Windows: it would search PATH and the CWD (DLL hijack).
        if (!OperatingSystem.IsWindows())
        {
            foreach (var n in names)
            {
                if (NativeLibrary.TryLoad(n, out IntPtr h3))
                {
                    LoadedPath = n;
                    return h3;
                }
            }
        }
        return IntPtr.Zero;
    }
}
