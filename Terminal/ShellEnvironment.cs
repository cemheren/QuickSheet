namespace ExcelConsole.Terminal;

/// <summary>Environment rules for shells started in t: panes.</summary>
internal static class ShellEnvironment
{
    // Variables that describe the *parent's* agent session. If QuickSheet itself was started from
    // inside a Claude Code session, passing these on would make a `claude` run in a pane think it
    // is that session's child (e.g. transcript saving off).
    private static readonly HashSet<string> s_sessionVars = new(StringComparer.Ordinal)
    {
        "CLAUDECODE", "CLAUDE_PID", "CLAUDE_EFFORT", "AI_AGENT",
        "CLAUDE_CODE_CHILD_SESSION", "CLAUDE_CODE_SESSION_ID", "CLAUDE_CODE_SESSION_ATTENDED",
        "CLAUDE_CODE_BRIDGE_SESSION_ID", "CLAUDE_CODE_ENTRYPOINT", "CLAUDE_CODE_EXECPATH",
        "CLAUDE_CODE_SSE_PORT",
    };

    /// <summary>True for variables a pane's shell should not inherit from QuickSheet.</summary>
    public static bool IsExcluded(string name) =>
        s_sessionVars.Contains(name)
        || name.StartsWith("CLAUDE_CODE_MESSAGING_", StringComparison.Ordinal)
        || name is "TERM" or "COLORTERM" or "TERM_PROGRAM" or "COLUMNS" or "LINES";
}
