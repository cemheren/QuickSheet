# Mastodon + Bluesky drafts — QuickSheet

Drafted 2026-05-14. Both networks favor single-post hooks with one strong visual. Threads OK but shorter than Twitter.

## Mastodon (500 char limit)

Recommended instance: `hachyderm.io`, `fosstodon.org`, or `mastodon.social` depending on where you have a presence. Hachyderm and Fosstodon skew toward devs / FOSS — better signal for QuickSheet.

### Single post

```
I made the desktop wallpaper an interactive spreadsheet.

QuickSheet runs as TUI or — with --desktop — embeds itself as a transparent grid behind your windows. Cell prefixes do the work: `r:` for runnable commands, `i:` for live subprocess output, `s:` for sparklines, `ext: github:u/r` to install an extension by git URL.

Cross-platform .NET 9. Zero NuGet dependencies (P/Invoke all the way).

#dotnet #linux #terminal #cli

https://github.com/cemheren/QuickSheet
```

Attach: launcher screenshot or — better — a short MP4 / animated WebP of wallpaper-mode interaction. Mastodon does well with MP4.

### Hashtag tips

- Pick 3–4 hashtags max, placed at the end.
- `#dotnet`, `#csharp`, `#terminal`, `#cli`, `#linux`, `#windows`, `#opensource` are all reasonable.
- Avoid generic hashtags like `#technology` or `#programming` — too noisy.

## Bluesky (300 char limit)

Bluesky leans more discussion-y; threads of 2–3 posts perform better than single mega-posts.

### Post 1 — the hook

```
my desktop wallpaper is now a spreadsheet 🟦

QuickSheet replaces the wallpaper with an interactive grid. notes, runnable commands, live process output, sparklines, hyperlinks — all in cells.

windows + linux. .NET 9. zero NuGet deps.

github.com/cemheren/QuickSheet
```

Attach a screenshot or short video.

### Post 2 — the wedge

```
cell prefixes are the whole feature set:

r: code .            ← launches VS Code
i: ping example.com  ← output streams into the cell live
s: A1::A10          ← renders the range as ▁▂▃▅▇ sparkline
ext: github:u/r      ← installs an extension by git URL
```

### Post 3 — extension ecosystem CTA

```
extensions live in their own repos. one git URL installs a new cell prefix. seven public so far: weather, crypto prices, TLS cert checker, dictionary, mortgage calc, MX records, pomodoro timer.

all zero-dep, all MIT.

docs.tour: github.com/cemheren/QuickSheet/blob/main/docs/tour.md
```

## Timing

- **Mastodon**: weekdays, 9–11am wherever your followers are concentrated. Avoid weekends — federated reach is bursty.
- **Bluesky**: late afternoon US Eastern works best (recent measurements). Bluesky's algorithm is chronological-ish, so post when your network is awake.

## Etiquette

- Mastodon: image alt-text is **mandatory** social etiquette. Write a short description of any attached screenshot. Posts without alt-text get downranked manually by instance admins.
- Bluesky: starter packs work. If anyone has a "developer tools" starter pack you fit into, ask the maintainer to add the repo.
- Do not crosspost identical text. Both platforms can detect copy-paste from Twitter and downrank.

## Cross-promotion notes

- Mastodon and Bluesky are mostly disjoint audiences from HN/Reddit. Posting on these alongside or after the HN window is fine — they don't penalize crossposting like Lobsters does.
- Twitter to Mastodon to Bluesky: rephrase the opener each time. Same screenshot/GIF is fine.
