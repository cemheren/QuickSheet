namespace ExcelConsole.Platform.Linux;

/// <summary>
/// Linux implementation of IExtensionEnvironment.
/// Extensions installed under ~/.quicksheet/extensions/.
/// </summary>
public class LinuxExtensionEnvironment : Extensions.IExtensionEnvironment
{
    public string ExtensionsDirectory { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
        ".quicksheet", "extensions");

    public (string shell, string arguments) GetShellCommand(string entryCommand, string workingDirectory)
    {
        return ("/bin/bash", $"-c \"cd {EscapeBash(workingDirectory)} && {entryCommand}\"");
    }

    private static string EscapeBash(string s) =>
        s.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("$", "\\$").Replace("`", "\\`");
}
