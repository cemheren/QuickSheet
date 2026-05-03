using System.Diagnostics;
using System.Text.Json;
using static ExcelConsole.Extensions.ExtensionProtocol;

namespace ExcelConsole.Extensions;

/// <summary>
/// Wraps a single extension child process. Communicates via JSON-lines over stdin/stdout.
/// Background reader thread consumes stdout and queues parsed messages.
/// </summary>
public class ExtensionProcess : IDisposable
{
    private Process? _process;
    private Thread? _readerThread;
    private Thread? _errorReaderThread;
    private volatile bool _disposed;

    private readonly object _lock = new();
    private readonly Queue<string> _incomingMessages = new();

    public string ExtensionName { get; }
    public string? RegisteredPrefix { get; private set; }
    public bool IsRunning => _process is { HasExited: false };

    public ExtensionProcess(string extensionName)
    {
        ExtensionName = extensionName;
    }

    /// <summary>
    /// Starts the extension process using the given entry command and working directory.
    /// </summary>
    public bool Start(IExtensionEnvironment env, string entryCommand, string workingDirectory)
    {
        var (shell, args) = env.GetShellCommand(entryCommand, workingDirectory);

        var psi = new ProcessStartInfo
        {
            FileName = shell,
            Arguments = args,
            UseShellExecute = false,
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            WorkingDirectory = workingDirectory
        };
        psi.Environment["QUICKSHEET_EXTENSIONS_DIR"] = env.ExtensionsDirectory;
        psi.Environment["QUICKSHEET_PROTOCOL_VERSION"] = ExtensionProtocol.ProtocolVersion.ToString();

        try
        {
            _process = Process.Start(psi);
            if (_process == null) return false;

            _readerThread = new Thread(ReadLoop) { IsBackground = true, Name = $"ExtReader-{ExtensionName}" };
            _readerThread.Start();

            _errorReaderThread = new Thread(ErrorReadLoop) { IsBackground = true, Name = $"ExtError-{ExtensionName}" };
            _errorReaderThread.Start();

            // Send init message
            SendMessage(new InitMessage());
            return true;
        }
        catch
        {
            _process?.Dispose();
            _process = null;
            return false;
        }
    }

    /// <summary>
    /// Sends a message to the extension process via stdin.
    /// </summary>
    public void SendMessage<T>(T message)
    {
        if (_process?.HasExited != false) return;
        try
        {
            string json = ExtensionProtocol.Serialize(message);
            _process.StandardInput.WriteLine(json);
            _process.StandardInput.Flush();
        }
        catch { }
    }

    /// <summary>
    /// Dequeues all pending messages from the extension. Thread-safe.
    /// </summary>
    public List<string> DrainMessages()
    {
        lock (_lock)
        {
            if (_incomingMessages.Count == 0) return [];
            var messages = _incomingMessages.ToList();
            _incomingMessages.Clear();
            return messages;
        }
    }

    /// <summary>
    /// Processes the register message from the extension, setting the prefix.
    /// </summary>
    public void HandleRegister(RegisterMessage msg)
    {
        RegisteredPrefix = msg.Prefix.ToLowerInvariant();
    }

    private void ReadLoop()
    {
        try
        {
            while (!_disposed && _process?.HasExited == false)
            {
                string? line = _process.StandardOutput.ReadLine();
                if (line == null) break;
                if (string.IsNullOrWhiteSpace(line)) continue;

                Console.Error.WriteLine($"[ext:{ExtensionName}:stdout] {line}");

                lock (_lock)
                {
                    _incomingMessages.Enqueue(line);
                }
            }
        }
        catch when (_disposed) { }
        catch { }
    }

    private void ErrorReadLoop()
    {
        try
        {
            while (!_disposed && _process?.HasExited == false)
            {
                string? line = _process.StandardError.ReadLine();
                if (line == null) break;
                if (string.IsNullOrWhiteSpace(line)) continue;

                Console.Error.WriteLine($"[ext:{ExtensionName}:stderr] {line}");
            }
        }
        catch when (_disposed) { }
        catch { }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        try
        {
            if (_process is { HasExited: false })
            {
                _process.StandardInput.Close();
                if (!_process.WaitForExit(3000))
                    _process.Kill();
            }
            _process?.Dispose();
        }
        catch { }
    }
}
