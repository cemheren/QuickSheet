namespace ExcelConsole.Terminal;

/// <summary>
/// A pseudo-terminal running an interactive shell for a t: pane.
/// Linux: <c>PtyProcess.StartShell</c> (openpty + posix_spawn). Windows: <c>ConPtyProcess.StartShell</c> (ConPTY).
/// </summary>
internal interface ITerminalPty : IDisposable
{
    bool HasExited { get; }
    int ExitCode { get; }

    /// <summary>Send input bytes to the program. Thread-safe.</summary>
    void Write(ReadOnlySpan<byte> data);

    /// <summary>Resize the pty window (the program gets SIGWINCH / a ConPTY resize).</summary>
    void Resize(int cols, int rows);
}
