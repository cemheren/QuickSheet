using System.Collections.Concurrent;
using System.Diagnostics;
using System.Text;

namespace ExcelConsole;

/// <summary>
/// Manages live subprocesses for inline command execution (i: cells pointing to r: cells).
/// Captures stdout/stderr asynchronously with a configurable line buffer cap.
/// Thread-safe: output can be read from the UI thread while processes write from background threads.
/// </summary>
public class InlineProcessManager : IDisposable
{
    private const int MaxOutputLines = 200;

    private readonly ConcurrentDictionary<(int row, int col), ManagedProcess> _processes = new();

    private class ManagedProcess : IDisposable
    {
        public Process? Process { get; set; }
        public string Command { get; set; } = "";
        private readonly object _lock = new();
        private readonly List<string> _outputLines = new();
        public bool HasNewOutput { get; set; }

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

        public string GetOutput()
        {
            lock (_lock)
            {
                HasNewOutput = false;
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
            // Same command already running → skip
            if (existing.Command == command && existing.Process != null && !existing.Process.HasExited)
                return;
            // Different command or exited → stop old, start new
            StopProcess(pointerRow, pointerCol);
        }

        var managed = new ManagedProcess { Command = command };

        // Parse command (strip r: prefix if present)
        string cmdText = command;
        if (CellPrefix.IsCommand(cmdText))
            cmdText = cmdText[3..].Trim();

        if (string.IsNullOrWhiteSpace(cmdText)) return;

        try
        {
            // Use cmd.exe on Windows to support shell commands
            var psi = new ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = $"/c {cmdText}",
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
                if (e.Data != null)
                    managed.AppendLine(e.Data);
            };

            proc.ErrorDataReceived += (_, e) =>
            {
                if (e.Data != null)
                    managed.AppendLine($"[err] {e.Data}");
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
    }

    /// <summary>
    /// Gets the current output buffer for a pointer cell. Returns null if no process.
    /// </summary>
    public string? GetOutput(int pointerRow, int pointerCol)
    {
        return _processes.TryGetValue((pointerRow, pointerCol), out var mp) ? mp.GetOutput() : null;
    }

    /// <summary>
    /// Returns true if any managed process has new output since last GetOutput call.
    /// </summary>
    public bool HasAnyNewOutput()
    {
        foreach (var mp in _processes.Values)
            if (mp.HasNewOutput) return true;
        return false;
    }

    /// <summary>
    /// Stops and removes the process for the given pointer cell.
    /// </summary>
    public void StopProcess(int pointerRow, int pointerCol)
    {
        if (_processes.TryRemove((pointerRow, pointerCol), out var mp))
            mp.Dispose();
    }

    /// <summary>
    /// Stops all managed processes.
    /// </summary>
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
