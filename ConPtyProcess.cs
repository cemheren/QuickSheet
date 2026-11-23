#if PLATFORM_WINDOWS
using System.Runtime.InteropServices;
using ExcelConsole.Terminal;
using Microsoft.Win32.SafeHandles;

namespace ExcelConsole;

/// <summary>
/// Runs a command inside a Windows Pseudo Console (ConPTY), capturing all output
/// including from programs that use Console APIs directly (TUI/interactive apps).
/// Raw output bytes are fed to a VT emulator (<see cref="IVtScreen"/>) which maintains
/// the final screen + scrollback.
/// </summary>
internal sealed class ConPtyProcess : ITerminalPty
{
    #region P/Invoke

    [StructLayout(LayoutKind.Sequential)]
    private struct COORD { public short X, Y; }

    [StructLayout(LayoutKind.Sequential)]
    private struct SECURITY_ATTRIBUTES
    {
        public int nLength;
        public IntPtr lpSecurityDescriptor;
        [MarshalAs(UnmanagedType.Bool)] public bool bInheritHandle;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct STARTUPINFOEX
    {
        public STARTUPINFO StartupInfo;
        public IntPtr lpAttributeList;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct STARTUPINFO
    {
        public int cb;
        public IntPtr lpReserved;
        public IntPtr lpDesktop;
        public IntPtr lpTitle;
        public int dwX, dwY, dwXSize, dwYSize;
        public int dwXCountChars, dwYCountChars;
        public int dwFillAttribute;
        public int dwFlags;
        public short wShowWindow;
        public short cbReserved2;
        public IntPtr lpReserved2;
        public IntPtr hStdInput, hStdOutput, hStdError;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct PROCESS_INFORMATION
    {
        public IntPtr hProcess;
        public IntPtr hThread;
        public int dwProcessId;
        public int dwThreadId;
    }

    private const uint EXTENDED_STARTUPINFO_PRESENT = 0x00080000;
    private static readonly IntPtr PROC_THREAD_ATTRIBUTE_PSEUDOCONSOLE = (IntPtr)0x00020016;
    private const uint STILL_ACTIVE = 259;

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool CreatePipe(out IntPtr hReadPipe, out IntPtr hWritePipe,
        ref SECURITY_ATTRIBUTES sa, int nSize);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern int CreatePseudoConsole(COORD size, IntPtr hInput, IntPtr hOutput,
        uint dwFlags, out IntPtr phPC);

    [DllImport("kernel32.dll")]
    private static extern int ResizePseudoConsole(IntPtr hPC, COORD size);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern void ClosePseudoConsole(IntPtr hPC);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool InitializeProcThreadAttributeList(IntPtr lpAttributeList,
        int dwAttributeCount, int dwFlags, ref IntPtr lpSize);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool UpdateProcThreadAttribute(IntPtr lpAttributeList, uint dwFlags,
        IntPtr Attribute, IntPtr lpValue, IntPtr cbSize, IntPtr lpPreviousValue, IntPtr lpReturnSize);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool DeleteProcThreadAttributeList(IntPtr lpAttributeList);

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool CreateProcessW(string? lpApplicationName, string lpCommandLine,
        IntPtr lpProcessAttributes, IntPtr lpThreadAttributes, bool bInheritHandles,
        uint dwCreationFlags, IntPtr lpEnvironment, string? lpCurrentDirectory,
        ref STARTUPINFOEX lpStartupInfo, out PROCESS_INFORMATION lpProcessInformation);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool CloseHandle(IntPtr hObject);

    [DllImport("kernel32.dll")]
    private static extern bool GetExitCodeProcess(IntPtr hProcess, out uint lpExitCode);

    #endregion

    private const int MaxLines = 5000;

    private IntPtr _hPC;
    private IntPtr _hProcess;
    private IntPtr _hThread;
    private IntPtr _outputReadPipe;
    private IntPtr _inputWritePipe;
    private IntPtr _attrList;
    private Thread? _readerThread;
    private volatile bool _disposed;

    private IVtScreen? _screen;

    // Interactive (t: pane) mode: raw output goes to this sink instead of a VtScreen.
    private Action<byte[], int>? _outputSink;
    private FileStream? _input;
    private readonly object _writeLock = new();

    private const uint CREATE_UNICODE_ENVIRONMENT = 0x00000400;

    public int ExitCode
    {
        get
        {
            if (_hProcess == IntPtr.Zero) return -1;
            GetExitCodeProcess(_hProcess, out uint code);
            return code == STILL_ACTIVE ? -1 : unchecked((int)code);
        }
    }

    public bool HasNewOutput => _screen?.IsDirty ?? false;

    public bool HasExited
    {
        get
        {
            if (_hProcess == IntPtr.Zero) return true;
            GetExitCodeProcess(_hProcess, out uint code);
            return code != STILL_ACTIVE;
        }
    }

    /// <summary>
    /// Starts the command in a pseudo console. Returns true on success.
    /// </summary>
    public bool Start(string command, int cols = 120, int rows = 30) =>
        StartCore($"cmd.exe /c {command}", cols, rows, null, null);

    /// <summary>
    /// Start an interactive shell for a t: pane: pwsh if installed, else Windows PowerShell, else
    /// %COMSPEC%. Output goes to <paramref name="onOutput"/> (reader thread, reused buffer).
    /// The shell starts in the user profile directory and inherits QuickSheet's environment minus
    /// parent agent-session variables.
    /// </summary>
    public bool StartShell(int cols, int rows, Action<byte[], int> onOutput)
    {
        _outputSink = onOutput;
        string shell = FindShell();
        string home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        return StartCore(shell, cols, rows, BuildEnvironmentBlock(), Directory.Exists(home) ? home : null);
    }

    private static string FindShell()
    {
        foreach (var exe in new[] { "pwsh.exe", "powershell.exe" })
        {
            foreach (var dir in (Environment.GetEnvironmentVariable("PATH") ?? "").Split(';', StringSplitOptions.RemoveEmptyEntries))
            {
                try
                {
                    string full = Path.Combine(dir.Trim(), exe);
                    if (Path.IsPathFullyQualified(full) && File.Exists(full)) return $"\"{full}\" -NoLogo";
                }
                catch { }
            }
        }
        return Environment.GetEnvironmentVariable("COMSPEC") is { Length: > 0 } c ? $"\"{c}\"" : "cmd.exe";
    }

    /// <summary>Unicode environment block: sorted "K=V\0" entries, then a final "\0".</summary>
    private static string BuildEnvironmentBlock()
    {
        var vars = new SortedDictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (System.Collections.DictionaryEntry e in Environment.GetEnvironmentVariables())
        {
            string k = (string)e.Key;
            if (k.Length == 0 || k.Contains('=') || ShellEnvironment.IsExcluded(k)) continue;
            vars[k] = (string?)e.Value ?? "";
        }
        vars["TERM_PROGRAM"] = "QuickSheet";
        vars["COLORTERM"] = "truecolor";
        var sb = new System.Text.StringBuilder();
        foreach (var (k, v) in vars) sb.Append(k).Append('=').Append(v).Append('\0');
        sb.Append('\0');
        return sb.ToString();
    }

    private bool StartCore(string cmdLine, int cols, int rows, string? environment, string? workingDir)
    {
        var sa = new SECURITY_ATTRIBUTES
        {
            nLength = Marshal.SizeOf<SECURITY_ATTRIBUTES>(),
            bInheritHandle = true
        };

        // Output pipe: we read, pseudo console writes
        if (!CreatePipe(out var outRead, out var outWrite, ref sa, 0))
            return false;

        // Input pipe: we write, pseudo console reads
        if (!CreatePipe(out var inRead, out var inWrite, ref sa, 0))
        {
            CloseHandle(outRead);
            CloseHandle(outWrite);
            return false;
        }

        _outputReadPipe = outRead;
        _inputWritePipe = inWrite;

        var size = new COORD { X = (short)Math.Clamp(cols, 1, short.MaxValue), Y = (short)Math.Clamp(rows, 1, short.MaxValue) };
        int hr = CreatePseudoConsole(size, inRead, outWrite, 0, out _hPC);

        // These ends now belong to the pseudo console
        CloseHandle(inRead);
        CloseHandle(outWrite);

        if (hr != 0)
        {
            Cleanup();
            return false;
        }

        // Attribute list with pseudo console handle
        IntPtr attrSize = IntPtr.Zero;
        InitializeProcThreadAttributeList(IntPtr.Zero, 1, 0, ref attrSize);
        _attrList = Marshal.AllocHGlobal(attrSize);
        if (!InitializeProcThreadAttributeList(_attrList, 1, 0, ref attrSize))
        {
            Cleanup();
            return false;
        }

        if (!UpdateProcThreadAttribute(_attrList, 0, PROC_THREAD_ATTRIBUTE_PSEUDOCONSOLE,
            _hPC, (IntPtr)IntPtr.Size, IntPtr.Zero, IntPtr.Zero))
        {
            Cleanup();
            return false;
        }

        var si = new STARTUPINFOEX();
        si.StartupInfo.cb = Marshal.SizeOf<STARTUPINFOEX>();
        si.lpAttributeList = _attrList;

        IntPtr envBlock = environment == null ? IntPtr.Zero : Marshal.StringToHGlobalUni(environment);
        bool created;
        PROCESS_INFORMATION pi;
        try
        {
            uint flags = EXTENDED_STARTUPINFO_PRESENT | (envBlock != IntPtr.Zero ? CREATE_UNICODE_ENVIRONMENT : 0);
            created = CreateProcessW(null, cmdLine, IntPtr.Zero, IntPtr.Zero, false,
                flags, envBlock, workingDir, ref si, out pi);
        }
        finally
        {
            if (envBlock != IntPtr.Zero) Marshal.FreeHGlobal(envBlock);
        }
        if (!created)
        {
            Cleanup();
            return false;
        }

        _hProcess = pi.hProcess;
        _hThread = pi.hThread;

        cols = Math.Clamp(cols, 1, short.MaxValue);
        rows = Math.Clamp(rows, 1, short.MaxValue);
        if (_outputSink == null)
            _screen = VtScreen.Create(cols, rows, MaxLines);
        else
            _input = new FileStream(new SafeFileHandle(_inputWritePipe, ownsHandle: false), FileAccess.Write, 1, false);

        _readerThread = new Thread(ReadOutput) { IsBackground = true, Name = "ConPTY-Reader" };
        _readerThread.Start();

        return true;
    }

    private void ReadOutput()
    {
        try
        {
            using var handle = new SafeFileHandle(_outputReadPipe, ownsHandle: false);
            using var stream = new FileStream(handle, FileAccess.Read, 4096, false);
            var buffer = new byte[4096];

            while (!_disposed)
            {
                int n;
                try { n = stream.Read(buffer, 0, buffer.Length); }
                catch { break; }
                if (n == 0) break;

                if (_outputSink is { } sink) sink(buffer, n);
                else _screen?.Write(buffer.AsSpan(0, n));
            }
        }
        catch { }
    }

    public string? GetOutput()
    {
        var t = _screen?.GetText();
        return string.IsNullOrEmpty(t) ? null : t;
    }

    /// <summary>Write input to the shell (keys, pastes, terminal replies). Thread-safe.</summary>
    public void Write(ReadOnlySpan<byte> data)
    {
        lock (_writeLock)
        {
            if (_disposed || _input == null) return;
            try { _input.Write(data); _input.Flush(); } catch { }
        }
    }

    /// <summary>Resize the pseudo console and the emulator screen.</summary>
    public void Resize(int cols, int rows)
    {
        cols = Math.Clamp(cols, 1, short.MaxValue);
        rows = Math.Clamp(rows, 1, short.MaxValue);
        if (_hPC != IntPtr.Zero)
            ResizePseudoConsole(_hPC, new COORD { X = (short)cols, Y = (short)rows });
        _screen?.Resize(cols, rows);
    }

    private void Cleanup()
    {
        if (_outputReadPipe != IntPtr.Zero) { CloseHandle(_outputReadPipe); _outputReadPipe = IntPtr.Zero; }
        if (_inputWritePipe != IntPtr.Zero) { CloseHandle(_inputWritePipe); _inputWritePipe = IntPtr.Zero; }
        if (_hPC != IntPtr.Zero) { ClosePseudoConsole(_hPC); _hPC = IntPtr.Zero; }
        if (_attrList != IntPtr.Zero)
        {
            DeleteProcThreadAttributeList(_attrList);
            Marshal.FreeHGlobal(_attrList);
            _attrList = IntPtr.Zero;
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        lock (_writeLock) { _input?.Dispose(); _input = null; }

        // Close pseudo console first — breaks the reader pipe
        if (_hPC != IntPtr.Zero) { ClosePseudoConsole(_hPC); _hPC = IntPtr.Zero; }

        _readerThread?.Join(2000);

        _screen?.Dispose();
        _screen = null;

        if (_hProcess != IntPtr.Zero) { CloseHandle(_hProcess); _hProcess = IntPtr.Zero; }
        if (_hThread != IntPtr.Zero) { CloseHandle(_hThread); _hThread = IntPtr.Zero; }
        if (_outputReadPipe != IntPtr.Zero) { CloseHandle(_outputReadPipe); _outputReadPipe = IntPtr.Zero; }
        if (_inputWritePipe != IntPtr.Zero) { CloseHandle(_inputWritePipe); _inputWritePipe = IntPtr.Zero; }
        if (_attrList != IntPtr.Zero)
        {
            DeleteProcThreadAttributeList(_attrList);
            Marshal.FreeHGlobal(_attrList);
            _attrList = IntPtr.Zero;
        }
    }
}
#endif
