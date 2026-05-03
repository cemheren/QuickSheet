using System.Collections.Concurrent;
using System.Diagnostics;
using System.Text;
#if PLATFORM_LINUX
using ExcelConsole.Platform.Linux;
#endif

namespace ExcelConsole;

/// <summary>
/// Manages live subprocesses for inline command execution (i: cells pointing to r: cells).
/// On Windows, uses ConPTY (Pseudo Console) to capture output from interactive/TUI programs.
/// On Linux, uses openpty + fork + exec (PtyProcess) for the same effect — gives the
/// child a real tty so streaming/TUI commands behave normally.
/// Thread-safe: output can be read from the UI thread while processes write from background threads.
/// </summary>
public class InlineProcessManager : IDisposable
{
    private const int MaxOutputLines = 200;

    private readonly ConcurrentDictionary<(int row, int col), ManagedProcess> _processes = new();

    private class ManagedProcess : IDisposable
    {
        public string Command { get; set; } = "";

#if PLATFORM_WINDOWS
        private ConPtyProcess? _pty;

        public bool HasNewOutput => _pty?.HasNewOutput ?? false;
        public bool HasExited => _pty?.HasExited ?? true;

        public void StartConPty(string cmdText, int ptyCols = 120, int ptyRows = 30)
        {
            _pty = new ConPtyProcess();
            if (!_pty.Start(cmdText, ptyCols, ptyRows))
            {
                _pty.Dispose();
                _pty = null;
            }
        }

        public string? GetOutput() => _pty?.GetOutput();

        public void Dispose() => _pty?.Dispose();
#else
        private PtyProcess? _pty;

        public bool HasNewOutput => _pty?.HasNewOutput ?? false;
        public bool HasExited => _pty?.HasExited ?? true;

        public void StartPty(string cmdText, int cols = 120, int rows = 30)
        {
            _pty = new PtyProcess();
            if (!_pty.Start(cmdText, cols, rows))
            {
                _pty.Dispose();
                _pty = null;
            }
        }

        public void AppendError(string msg)
        {
            _pty?.Dispose();
            _pty = PtyProcess.CreateError(msg);
        }

        public string? GetOutput() => _pty?.GetOutput();

        public void Dispose() => _pty?.Dispose();
#endif
    }

    /// <summary>
    /// Starts or restarts a process for the given pointer cell.
    /// ptyCols/ptyRows set the pseudo console size to match the visual span.
    /// </summary>
    public void EnsureRunning(int pointerRow, int pointerCol, string command, int ptyCols = 120, int ptyRows = 30)
    {
        var key = (pointerRow, pointerCol);

        if (_processes.TryGetValue(key, out var existing))
        {
            if (existing.Command == command)
                return; // same command — keep it and its output
            StopProcess(pointerRow, pointerCol);
        }

        var managed = new ManagedProcess { Command = command };

        string cmdText = command;
        if (CellPrefix.IsCommand(cmdText))
            cmdText = cmdText[3..].Trim();

        if (string.IsNullOrWhiteSpace(cmdText)) return;

#if PLATFORM_WINDOWS
        managed.StartConPty(cmdText, ptyCols, ptyRows);
        _processes[key] = managed;
#else
        try
        {
            managed.StartPty(cmdText, ptyCols, ptyRows);
            _processes[key] = managed;
        }
        catch (Exception ex)
        {
            managed.AppendError($"[failed to start: {ex.Message}]");
            _processes[key] = managed;
        }
#endif
    }

    /// <summary>
    /// Gets the current output buffer for a pointer cell. Returns null if no process.
    /// </summary>
    public string? GetOutput(int pointerRow, int pointerCol)
    {
        return _processes.TryGetValue((pointerRow, pointerCol), out var mp) ? mp.GetOutput() : null;
    }

    /// <summary>
    /// Returns true if any managed process has new output or is still running.
    /// </summary>
    public bool HasAnyNewOutput()
    {
        foreach (var mp in _processes.Values)
            if (mp.HasNewOutput || !mp.HasExited)
                return true;
        return false;
    }

    public void StopProcess(int pointerRow, int pointerCol)
    {
        if (_processes.TryRemove((pointerRow, pointerCol), out var mp))
            mp.Dispose();
    }

    public void StopAll()
    {
        foreach (var key in _processes.Keys.ToList())
            if (_processes.TryRemove(key, out var mp))
                mp.Dispose();
    }

    public void Dispose()
    {
        StopAll();
    }
}
