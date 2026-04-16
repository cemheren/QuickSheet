using System.Collections.Concurrent;
using System.Diagnostics;
using System.Text;

namespace ExcelConsole;

/// <summary>
/// Manages live subprocesses for inline command execution (i: cells pointing to r: cells).
/// On Windows, uses ConPTY (Pseudo Console) to capture output from interactive/TUI programs.
/// On Linux, falls back to stdout/stderr pipe redirection.
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

        public void StartConPty(string cmdText)
        {
            _pty = new ConPtyProcess();
            if (!_pty.Start(cmdText))
            {
                _pty.Dispose();
                _pty = null;
            }
        }

        public string? GetOutput() => _pty?.GetOutput();

        public void Dispose() => _pty?.Dispose();
#else
        public Process? Process { get; set; }
        public bool HasNewOutput { get; set; }
        public bool HasExited => Process?.HasExited ?? true;
        private readonly object _lock = new();
        private readonly List<string> _outputLines = new();

        public void AppendLine(string line)
        {
            lock (_lock)
            {
                _outputLines.Add(line);
                while (_outputLines.Count > MaxOutputLines)
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

        public void Dispose()
        {
            try
            {
                if (Process != null && !Process.HasExited)
                    Process.Kill(entireProcessTree: true);
            }
            catch { }
            Process?.Dispose();
        }
#endif
    }

    /// <summary>
    /// Starts or restarts a process for the given pointer cell.
    /// If a process is already running for this cell with the same command, does nothing.
    /// </summary>
    public void EnsureRunning(int pointerRow, int pointerCol, string command)
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
        managed.StartConPty(cmdText);
        _processes[key] = managed;
#else
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = "/bin/sh",
                Arguments = $"-c \"{cmdText.Replace("\"", "\\\"")}\"",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8,
            };

            var proc = new Process { StartInfo = psi, EnableRaisingEvents = true };
            managed.Process = proc;

            proc.OutputDataReceived += (_, e) =>
            {
                if (e.Data != null) managed.AppendLine(e.Data);
            };
            proc.ErrorDataReceived += (_, e) =>
            {
                if (e.Data != null) managed.AppendLine($"[err] {e.Data}");
            };

            proc.Start();
            proc.BeginOutputReadLine();
            proc.BeginErrorReadLine();
            _processes[key] = managed;
        }
        catch (Exception ex)
        {
            managed.AppendLine($"[failed to start: {ex.Message}]");
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
