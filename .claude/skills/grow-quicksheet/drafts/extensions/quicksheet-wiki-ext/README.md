# quicksheet-wiki-ext

Wikipedia article summaries for [QuickSheet](https://github.com/cemheren/QuickSheet). Instant reference right on your wallpaper — no browser tab needed.

## Install

Type into any cell:

```
ext: github:cemheren/quicksheet-wiki-ext
```

## Use

```
wiki: Pythagorean theorem, 1, 5
```

Fills 5 rows with the article title and summary extract:

```
📖 Pythagorean theorem
In mathematics, the Pythagorean theorem
or Pythagoras' theorem is a fundamental
relation in Euclidean geometry between
the three sides of a right triangle.
```

### More examples

```
wiki: Rust (programming language), 1, 4
wiki: Marie Curie, 1, 6
wiki: Black hole, 1, 5
```

## Why

Students: pin definitions and quick-reference summaries to the wallpaper while working. Researchers: keep context visible without tab-switching. Pairs with `define:` for single-word lookups and `cite:` for formal citations.

## Build

Requires .NET 9. Zero NuGet dependencies — only BCL networking and JSON.

```
dotnet build WikiExtension.csproj
```

## Notes

- Uses the Wikimedia REST API `/page/summary/` endpoint.
- No API key required. Be polite — one request per activation, no polling.
- English Wikipedia only. Change the URL hostname for other languages.
- Text is word-wrapped to ~50 chars per row for typical cell widths.
- Use `L: <cell>, 0` (one-shot) — wiki content doesn't change frequently.

## Protocol

Reads JSON lines on stdin, writes JSON lines on stdout. See QuickSheet's main README for the full extension protocol spec.

## License

MIT
