# quicksheet-ghstreak-ext

A QuickSheet extension that shows a GitHub user's commit-day streak + commits-today on the wallpaper. Pairs with [quicksheet-leetcode-ext](https://github.com/cemheren/quicksheet-leetcode-ext) for the **CS-student wallpaper flex bundle**.

```
ext: github:cemheren/quicksheet-ghstreak-ext
ghstreak: <github-username>
```

## Output

One row, four columns:

| @username  | 5 today | 🔥 12d streak | 247 in 90d |
|------------|---------|---------------|------------|

- **today**: commits pushed today (UTC), summed across all public repos.
- **streak**: consecutive days (UTC) ending today with at least one public push. If today has no commits yet, the streak counts from yesterday.
- **90d**: total commits across the last ~90 days. GitHub's public events feed only retains 300 events / 90 days; anything older is invisible.

## Why on the wallpaper

- CS-student / grinder identity: a visible commit-today number is the modern "I'm not just AI-vibing" flex.
- Pairs naturally with `leetcode:` in the same row. r/unixporn rice candidate (see `docs/for-students.md` in the main repo once published).
- Behind every IDE window — alt-tab and you see whether you've pushed today.

## Honesty

- **Public events only.** Uses the unauthenticated `https://api.github.com/users/<u>/events/public` endpoint. Private commits and contributions to private repos don't show up. Heavy contributors-to-private-orgs will see suspiciously-low numbers.
- **No auth.** The unauthenticated rate limit is 60 requests/hour per IP. The extension caches per-username for 15 minutes — enough headroom for several wallpaper installs on one network.
- **Today is UTC.** Your local calendar day may not match GitHub's UTC day; expect ±1 day skew for night-owl committers.
- **Not the "official" streak count.** GitHub removed streak counters from profiles in 2016. This re-derives a streak from public events — it'll be close but won't match third-party "longest streak" trackers that scrape the contributions calendar SVG.

## Build

```bash
dotnet build GhStreakExtension.csproj
```

.NET 9, MIT, zero NuGet dependencies (`HttpClient` + `System.Text.Json` only).

## Protocol

Standard QuickSheet JSON-lines over stdin/stdout:

1. On startup, emit `{"type":"register","prefix":"ghstreak","name":"GitHub Streak","version":"1.0.0"}`.
2. On each `{"type":"activate","id":"...","params":["<username>"]}`, fetch (or load from 15-min cache), parse PushEvent payloads, and reply with `{"type":"write","id":"...","cells":[{"r":0,"c":0,"v":"@<user>"}, ...]}`.

See [docs/extension-protocol.md](https://github.com/cemheren/QuickSheet/blob/main/docs/extension-protocol.md).

## License

MIT.
