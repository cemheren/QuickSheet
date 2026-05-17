# quicksheet-leetcode-ext

A QuickSheet extension that shows a LeetCode user's solved count + difficulty breakdown + current streak on the wallpaper. Aimed at CS students and grinders who want their progress glanceable behind every IDE window.

```
ext: github:cemheren/quicksheet-leetcode-ext
leetcode: <username>
```

## Output

One row, four columns:

| @username    | total solved | difficulty breakdown      | streak       |
|--------------|--------------|---------------------------|--------------|
| @neetcode    | 312 solved   | E 124 · M 156 · H 32      | 🔥 42d       |

## Why this lives on a wallpaper

- The wallpaper is behind every window — every alt-tab out of VSCode shows your number.
- One cell per LeetCode account; pair with `gh:` (planned) for a GitHub commit-today cell next to it.
- Built for the post-LLM CS-major identity moment: a visible solved-count is the modern "I'm not just AI-vibing" flex.

## Honesty

- **Public profile data only.** Uses LeetCode's public GraphQL endpoint; no auth, no token, your password is not involved. Private profiles return "user not found."
- **Cached 1 hour.** LeetCode rate-limits hard; the extension caches per-username under `$XDG_CACHE_HOME/quicksheet-leetcode/` (Linux) or `%LOCALAPPDATA%\quicksheet-leetcode\` (Windows). The streak shown can therefore be up to an hour stale.
- **Not an account manager.** No submitting solutions, no problem browser. There are other LeetCode tools for that. This is one cell.

## Build

```bash
dotnet build LeetCodeExtension.csproj
```

.NET 9, MIT, zero NuGet dependencies (stdlib `HttpClient` + `System.Text.Json` only).

## Protocol

Standard QuickSheet JSON-lines over stdin/stdout:

1. On startup, emit `{"type":"register","prefix":"leetcode","name":"LeetCode Stats","version":"1.0.0"}`.
2. On each `{"type":"activate","id":"...","params":["<username>"]}`, query the LeetCode GraphQL endpoint (cache-aware), and reply with `{"type":"write","id":"...","cells":[{"r":0,"c":0,"v":"@<user>"}, {"r":0,"c":1,"v":"<N> solved"}, ...]}`.

See [docs/extension-protocol.md](https://github.com/cemheren/QuickSheet/blob/main/docs/extension-protocol.md).

## License

MIT.
