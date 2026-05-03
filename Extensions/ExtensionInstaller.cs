using System.Diagnostics;
using System.Text.Json;

namespace ExcelConsole.Extensions;

/// <summary>
/// Handles installing extensions from GitHub repositories via git clone.
/// </summary>
public static class ExtensionInstaller
{
    /// <summary>
    /// Installs an extension from a GitHub reference like "github:user/repo".
    /// Clones into the extensions directory if not already present.
    /// Returns the local directory path, or null on failure.
    /// </summary>
    public static string? Install(string source, IExtensionEnvironment env)
    {
        if (!source.StartsWith("github:", StringComparison.OrdinalIgnoreCase))
            return null;

        string repoRef = source[7..]; // strip "github:"
        if (string.IsNullOrWhiteSpace(repoRef) || !repoRef.Contains('/'))
            return null;

        // Derive local directory name from repo (user--repo)
        string dirName = repoRef.Replace('/', '-').Replace('\\', '-');
        string targetDir = Path.Combine(env.ExtensionsDirectory, dirName);

        if (Directory.Exists(targetDir))
        {
            // Already installed — pull latest changes
            string manifestPath = Path.Combine(targetDir, "quicksheet-extension.json");
            if (!File.Exists(manifestPath)) return null;
            TryPull(targetDir);
            return targetDir;
        }

        // Ensure parent directory exists
        Directory.CreateDirectory(env.ExtensionsDirectory);

        string gitUrl = $"https://github.com/{repoRef}.git";

        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = "git",
                Arguments = $"clone --depth 1 {gitUrl} {targetDir}",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            using var proc = Process.Start(psi);
            if (proc == null) return null;

            proc.WaitForExit(60_000); // 60s timeout
            if (proc.ExitCode != 0) return null;

            string manifestPath = Path.Combine(targetDir, "quicksheet-extension.json");
            return File.Exists(manifestPath) ? targetDir : null;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Pulls latest changes from the remote. Non-blocking, best-effort.
    /// </summary>
    private static void TryPull(string extensionDir)
    {
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = "git",
                Arguments = "pull --ff-only",
                WorkingDirectory = extensionDir,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            using var proc = Process.Start(psi);
            proc?.WaitForExit(15_000); // 15s timeout
        }
        catch { }
    }

    /// <summary>
    /// Reads the extension manifest from a directory.
    /// </summary>
    public static ExtensionManifest? ReadManifest(string extensionDir)
    {
        string manifestPath = Path.Combine(extensionDir, "quicksheet-extension.json");
        if (!File.Exists(manifestPath)) return null;

        try
        {
            string json = File.ReadAllText(manifestPath);
            return JsonSerializer.Deserialize<ExtensionManifest>(json, new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            });
        }
        catch
        {
            return null;
        }
    }
}

/// <summary>
/// Represents the quicksheet-extension.json manifest file.
/// </summary>
public class ExtensionManifest
{
    public string Name { get; set; } = "";
    public string Version { get; set; } = "";
    public string Prefix { get; set; } = "";
    public string Description { get; set; } = "";
    public string Entry { get; set; } = "";
    public int MinProtocolVersion { get; set; } = 1;
}
