namespace ExcelConsole.Extensions;

/// <summary>
/// Platform-specific environment for extensions.
/// Abstracts directory paths and process-launching details that differ between OSes.
/// </summary>
public interface IExtensionEnvironment
{
    /// <summary>
    /// Root directory where extensions are installed.
    /// e.g., ~/.quicksheet/extensions/ on Linux, %APPDATA%/QuickSheet/extensions/ on Windows.
    /// </summary>
    string ExtensionsDirectory { get; }

    /// <summary>
    /// Resolves the shell and arguments needed to run an extension entry command.
    /// For example, on Linux: shell="/bin/bash", args="-c \"dotnet run ...\"".
    /// On Windows: shell="cmd.exe", args="/c dotnet run ...".
    /// </summary>
    (string shell, string arguments) GetShellCommand(string entryCommand, string workingDirectory);
}
