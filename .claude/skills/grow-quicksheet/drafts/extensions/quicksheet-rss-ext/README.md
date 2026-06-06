# quicksheet-rss-ext

**RSS/Atom feed reader for [QuickSheet](https://github.com/cemheren/QuickSheet) — live headlines on your desktop wallpaper.**

Never miss a post. Pin your favourite feeds to the desktop and see the latest headlines update in real time — no browser tab, no notification noise.

## Install

```
ext: github:cemheren/quicksheet-rss-ext
```

Requires QuickSheet with extension support.

## Usage

Type into any cell:

```
rss: https://hnrss.org/newest?points=50
```

Or point to a file listing multiple feeds:

```
rss: ~/feeds.txt
```

Where `feeds.txt` contains one URL per line:

```
# My feeds
https://hnrss.org/newest?points=50
https://feeds.arstechnica.com/arstechnica/index
https://lobste.rs/rss
https://blog.golang.org/feed.atom
```

## Output

| Column | Content |
|--------|---------|
| 0 | Article title (truncated to 60 chars) |
| 1 | Published date (MM-dd HH:mm) |
| 2 | Link URL |

Items from multiple feeds are merged and sorted by date (newest first). Up to `gridRows` items are shown (default 10).

## Examples

```
rss: https://lobste.rs/rss
rss: https://feeds.arstechnica.com/arstechnica/index
rss: https://hnrss.org/newest?points=100
rss: https://www.reddit.com/r/commandline/.rss
```

## Screenshot

<!-- TODO: Add screenshot showing headlines on desktop -->

## How it works

- Fetches RSS 2.0 and Atom feeds using .NET's built-in `HttpClient` + `System.Xml`
- Zero NuGet dependencies — just the .NET 9 SDK
- Parses `<item>` (RSS) and `<entry>` (Atom) elements
- Supports feed list files for multi-feed aggregation
- 10-second timeout per feed; errors shown inline

## Build

```bash
dotnet build RssExtension.csproj
```

## License

MIT
