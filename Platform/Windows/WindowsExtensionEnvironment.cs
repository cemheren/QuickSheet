namespace ExcelConsole.Platform.Windows;

/// <summary>
/// Windows implementation of IExtensionEnvironment.
/// Extensions installed under %APPDATA%/QuickSheet/extensions/.
/// </summary>
public class WindowsExtensionEnvironment : Extensions.IExtensionEnvironment
{
    public string ExtensionsDirectory { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "QuickSheet", "extensions");

    public (string shell, string arguments) GetShellCommand(string entryCommand, string workingDirectory)
    {
        return ("cmd.exe", $"/c cd /d \"{workingDirectory}\" && {entryCommand}");
    }
}
