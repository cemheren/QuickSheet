#if PLATFORM_WINDOWS
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Win32.SafeHandles;

namespace ExcelConsole;

/// <summary>
/// Runs a command inside a Windows Pseudo Console (ConPTY), capturing all output
/// including from programs that use Console APIs directly (TUI/interactive apps).
/// </summary>
internal sealed class ConPtyProcess : IDisposable
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

    private const int MaxLines = 200;

    private IntPtr _hPC;
    private IntPtr _hProcess;
    private IntPtr _hThread;
    private IntPtr _outputReadPipe;
    private IntPtr _inputWritePipe;
    private IntPtr _attrList;
    private Thread? _readerThread;
    private volatile bool _disposed;

    private readonly object _lock = new();
    private readonly List<string> _outputLines = new();

    public bool HasNewOutput { get; set; }

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
    public bool Start(string command, int cols = 120, int rows = 30)
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

        var size = new COORD { X = (short)cols, Y = (short)rows };
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

        string cmdLine = $"cmd.exe /c {command}";

        if (!CreateProcessW(null, cmdLine, IntPtr.Zero, IntPtr.Zero, false,
            EXTENDED_STARTUPINFO_PRESENT, IntPtr.Zero, null, ref si, out var pi))
        {
            Cleanup();
            return false;
        }

        _hProcess = pi.hProcess;
        _hThread = pi.hThread;

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
            var lineBuffer = new StringBuilder();

            while (!_disposed)
            {
                int n;
                try { n = stream.Read(buffer, 0, buffer.Length); }
                catch { break; }
                if (n == 0) break;

                string text = Encoding.UTF8.GetString(buffer, 0, n);

                foreach (char ch in text)
                {
                    if (ch == '\n')
                    {
                        string line = StripAnsi(lineBuffer.ToString().TrimEnd('\r'));
                        if (!string.IsNullOrEmpty(line))
                            AppendLine(line);
                        lineBuffer.Clear();
                    }
                    else
                    {
                        lineBuffer.Append(ch);
                    }
                }
            }

            if (lineBuffer.Length > 0)
            {
                string final = StripAnsi(lineBuffer.ToString().TrimEnd('\r'));
                if (!string.IsNullOrEmpty(final))
                    AppendLine(final);
            }
        }
        catch { }
    }

    private void AppendLine(string line)
    {
        lock (_lock)
        {
            _outputLines.Add(line);
            while (_outputLines.Count > MaxLines)
                _outputLines.RemoveAt(0);
            HasNewOutput = true;
        }
    }

    public string? GetOutput()
    {
        lock (_lock)
        {
            HasNewOutput = false;
            if (_outputLines.Count == 0) return null;
            return string.Join("\n", _outputLines);
        }
    }

    // Strips ANSI/VT100 escape sequences and control chars
    private static readonly Regex AnsiPattern = new(
        @"\x1b(?:\[[0-9;?]*[A-Za-z@~]|\][^\x07]*\x07|\([A-Z0-9]|[A-Z=<>7-9])|[\x00-\x08\x0e-\x1f]",
        RegexOptions.Compiled);

    private static string StripAnsi(string text) =>
        AnsiPattern.Replace(text, "");

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

        // Close pseudo console first — breaks the reader pipe
        if (_hPC != IntPtr.Zero) { ClosePseudoConsole(_hPC); _hPC = IntPtr.Zero; }

        _readerThread?.Join(2000);

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
