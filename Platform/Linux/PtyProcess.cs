#if PLATFORM_LINUX
using System.Runtime.InteropServices;
using ExcelConsole.Terminal;

namespace ExcelConsole.Platform.Linux;

/// <summary>
/// Runs a command inside a Linux pseudo-terminal (PTY) via openpty + fork + exec.
/// Linux equivalent of Windows ConPTY: gives child a real tty, so streaming/TUI
/// programs (e.g. claude, top, progress bars) emit live output instead of
/// detecting a pipe and going to fully-buffered mode.
///
/// Raw output bytes are fed to a VT emulator (<see cref="IVtScreen"/>) which keeps
/// the final screen + scrollback.
///
/// Only libutil.so.1 + libc.so.6 — both present on every Linux. Zero NuGet deps
/// per project policy.
/// </summary>
internal sealed class PtyProcess : IDisposable
{
    #region P/Invoke

    [StructLayout(LayoutKind.Sequential)]
    private struct Winsize
    {
        public ushort ws_row, ws_col, ws_xpixel, ws_ypixel;
    }

    // Linux ioctl request codes (see <asm-generic/ioctls.h>).
    private const ulong TIOCSWINSZ = 0x5414;
    private const ulong TIOCSCTTY  = 0x540E;

    private const int WNOHANG = 1;
    private const int SIGTERM = 15;
    private const int SIGKILL = 9;

    [DllImport("libutil.so.1", SetLastError = true)]
    private static extern int openpty(out int amaster, out int aslave,
        IntPtr name, IntPtr termp, ref Winsize winp);

    [DllImport("libc.so.6", SetLastError = true)] private static extern int close(int fd);
    [DllImport("libc.so.6", SetLastError = true)] private static extern int kill(int pid, int sig);
    [DllImport("libc.so.6", SetLastError = true)] private static extern int waitpid(int pid, out int status, int options);
    [DllImport("libc.so.6", SetLastError = true)] private static extern int ioctl(int fd, ulong request, IntPtr arg);
    [DllImport("libc.so.6", SetLastError = true)] private static extern long read(int fd, byte[] buf, ulong count);

    // posix_spawn family — atomic fork+exec, no managed code in between (the
    // ".NET fork" hazard goes away).
    [DllImport("libc.so.6", SetLastError = true)]
    private static extern int posix_spawn(out int pid, IntPtr path, IntPtr fileActions,
        IntPtr attrp, IntPtr[] argv, IntPtr[] envp);

    [DllImport("libc.so.6", SetLastError = true)]
    private static extern int posix_spawn_file_actions_init(IntPtr fileActions);

    [DllImport("libc.so.6", SetLastError = true)]
    private static extern int posix_spawn_file_actions_destroy(IntPtr fileActions);

    [DllImport("libc.so.6", SetLastError = true)]
    private static extern int posix_spawn_file_actions_adddup2(IntPtr fileActions, int fd, int newfd);

    [DllImport("libc.so.6", SetLastError = true)]
    private static extern int posix_spawn_file_actions_addclose(IntPtr fileActions, int fd);

    [DllImport("libc.so.6", SetLastError = true)]
    private static extern int posix_spawn_file_actions_addopen(IntPtr fileActions, int fd,
        IntPtr path, int oflag, uint mode);

    private static readonly IntPtr s_devNullPath = Marshal.StringToHGlobalAnsi("/dev/null");
    private const int O_RDONLY = 0;

    #endregion

    private const int MaxLines = 5000;

    private int _masterFd = -1;
    private int _pid = -1;
    private Thread? _readerThread;
    private volatile bool _disposed;
    private volatile bool _exited;
    private int _exitCode = -1;

    private IVtScreen? _screen;
    // Exit marker / error text live outside the emulator on purpose: a TUI may have
    // left the cursor mid-screen, so we never feed synthetic text into it.
    private volatile string? _exitMarker;
    private volatile bool _markerPending;
    private string? _errorText;

    public bool HasNewOutput => (_screen?.IsDirty ?? false) || _markerPending;

    public bool HasExited
    {
        get
        {
            if (_exited) return true;
            if (_pid <= 0) return true;
            int r = waitpid(_pid, out int status, WNOHANG);
            if (r == _pid)
            {
                _exited = true;
                _exitCode = (status >> 8) & 0xff;
                return true;
            }
            return false;
        }
    }

    public int ExitCode => _exitCode;

    /// <summary>Pre-populated error placeholder (no child process). Always reports HasExited=true.</summary>
    public static PtyProcess CreateError(string message)
    {
        var p = new PtyProcess();
        p._exited = true;
        p._errorText = message;
        p._markerPending = true;
        return p;
    }

    /// <summary>
    /// Start command in PTY. Returns true on success.
    /// cols/rows match visual span — child sees real terminal of that size.
    /// </summary>
    public bool Start(string command, int cols = 120, int rows = 30)
    {
        var ws = new Winsize { ws_row = (ushort)Math.Max(1, rows), ws_col = (ushort)Math.Max(1, cols) };

        if (openpty(out int master, out int slave, IntPtr.Zero, IntPtr.Zero, ref ws) != 0)
            return false;

        // Pre-marshal everything the child needs BEFORE fork. After fork, the
        // child must only call async-signal-safe syscalls (no allocs, no locks).
        IntPtr pPath = Marshal.StringToHGlobalAnsi("/bin/sh");
        IntPtr pArg0 = Marshal.StringToHGlobalAnsi("sh");
        IntPtr pArg1 = Marshal.StringToHGlobalAnsi("-c");
        IntPtr pArg2 = Marshal.StringToHGlobalAnsi(command);
        IntPtr pTerm = Marshal.StringToHGlobalAnsi("TERM=xterm-256color");

        // Build COLUMNS/LINES env so libreadline / Python / curses respect span size.
        IntPtr pCols = Marshal.StringToHGlobalAnsi($"COLUMNS={cols}");
        IntPtr pLines = Marshal.StringToHGlobalAnsi($"LINES={rows}");

        // Inherit a few useful env vars from parent.
        string? home = Environment.GetEnvironmentVariable("HOME");
        string? path = Environment.GetEnvironmentVariable("PATH");
        string? user = Environment.GetEnvironmentVariable("USER");
        IntPtr pHome = home is null ? IntPtr.Zero : Marshal.StringToHGlobalAnsi($"HOME={home}");
        IntPtr pPathEnv = path is null ? IntPtr.Zero : Marshal.StringToHGlobalAnsi($"PATH={path}");
        IntPtr pUser = user is null ? IntPtr.Zero : Marshal.StringToHGlobalAnsi($"USER={user}");

        var argvList = new List<IntPtr> { pArg0, pArg1, pArg2, IntPtr.Zero };
        var envpList = new List<IntPtr> { pTerm, pCols, pLines };
        if (pHome != IntPtr.Zero) envpList.Add(pHome);
        if (pPathEnv != IntPtr.Zero) envpList.Add(pPathEnv);
        if (pUser != IntPtr.Zero) envpList.Add(pUser);
        envpList.Add(IntPtr.Zero);

        IntPtr[] argv = argvList.ToArray();
        IntPtr[] envp = envpList.ToArray();

        // posix_spawn_file_actions_t is opaque. Glibc x86-64 size is ~76 bytes;
        // allocate generously to stay safe across libc versions.
        IntPtr fa = Marshal.AllocHGlobal(256);
        int spawnRc;
        int pid;
        try
        {
            if (posix_spawn_file_actions_init(fa) != 0) { close(master); close(slave); return false; }
            // stdin = /dev/null (non-tty) so programs that check isatty(0) treat themselves
            // as non-interactive (claude, less, etc. exit instead of dropping into a REPL).
            // stdout + stderr = pty slave so programs still detect a tty for output and stream
            // line-by-line instead of fully buffering.
            posix_spawn_file_actions_addopen(fa, 0, s_devNullPath, O_RDONLY, 0);
            posix_spawn_file_actions_adddup2(fa, slave, 1);
            posix_spawn_file_actions_adddup2(fa, slave, 2);
            posix_spawn_file_actions_addclose(fa, slave);
            posix_spawn_file_actions_addclose(fa, master);

            spawnRc = posix_spawn(out pid, pPath, fa, IntPtr.Zero, argv, envp);
            posix_spawn_file_actions_destroy(fa);
        }
        finally { Marshal.FreeHGlobal(fa); }

        if (spawnRc != 0)
        {
            close(master); close(slave);
            FreeAll(pPath, pArg0, pArg1, pArg2, pTerm, pCols, pLines, pHome, pPathEnv, pUser);
            return false;
        }

        // ── PARENT ──
        close(slave);
        FreeAll(pPath, pArg0, pArg1, pArg2, pTerm, pCols, pLines, pHome, pPathEnv, pUser);

        _masterFd = master;
        _pid = pid;

        _screen = VtScreen.Create(Math.Max(1, cols), Math.Max(1, rows), MaxLines);

        _readerThread = new Thread(ReadLoop) { IsBackground = true, Name = "PTY-Reader" };
        _readerThread.Start();
        return true;
    }

    private static void FreeAll(params IntPtr[] ptrs)
    {
        foreach (var p in ptrs) if (p != IntPtr.Zero) Marshal.FreeHGlobal(p);
    }

    private void ReadLoop()
    {
        var buf = new byte[4096];
        try
        {
            while (!_disposed)
            {
                long n = read(_masterFd, buf, (ulong)buf.Length);
                if (n <= 0) break; // EOF or EIO (slave hung up) → child exited
                _screen?.Write(buf.AsSpan(0, (int)n));
            }
        }
        catch { }

        _exited = true;
        if (_pid > 0)
        {
            // Reap to set exit code.
            int r = waitpid(_pid, out int status, 0);
            if (r == _pid) _exitCode = (status >> 8) & 0xff;
        }
        _exitMarker = $"[exited {_exitCode}]";
        _markerPending = true;
    }

    public string? GetOutput()
    {
        // Clear first so a marker set concurrently re-arms the flag.
        _markerPending = false;
        if (_errorText is { } err) return err;

        string text = _screen?.GetText() ?? "";
        string? marker = _exitMarker;
        if (marker is null) return text.Length == 0 ? null : text;
        return text.Length == 0 ? marker : text + "\n" + marker;
    }

    /// <summary>Resize the PTY's window. Programs receive SIGWINCH.</summary>
    public void Resize(int cols, int rows)
    {
        _screen?.Resize(cols, rows);
        if (_masterFd < 0) return;
        var ws = new Winsize { ws_row = (ushort)Math.Max(1, rows), ws_col = (ushort)Math.Max(1, cols) };
        IntPtr p = Marshal.AllocHGlobal(Marshal.SizeOf<Winsize>());
        try
        {
            Marshal.StructureToPtr(ws, p, false);
            ioctl(_masterFd, TIOCSWINSZ, p);
        }
        finally { Marshal.FreeHGlobal(p); }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        if (_pid > 0 && !_exited)
        {
            kill(_pid, SIGTERM);
            // Brief grace period.
            for (int i = 0; i < 20 && !_exited; i++)
            {
                int r = waitpid(_pid, out _, WNOHANG);
                if (r == _pid) { _exited = true; break; }
                Thread.Sleep(25);
            }
            if (!_exited) kill(_pid, SIGKILL);
        }

        if (_masterFd >= 0) { close(_masterFd); _masterFd = -1; }
        _readerThread?.Join(500);

        _screen?.Dispose();
        _screen = null;
    }
}
#endif
