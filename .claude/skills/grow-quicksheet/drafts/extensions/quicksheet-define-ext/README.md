# quicksheet-define-ext

Inline dictionary lookups for [QuickSheet](https://github.com/cemheren/QuickSheet). Uses the free, no-key dictionaryapi.dev service.

## Install

Type into any cell:

```
ext: github:cemheren/quicksheet-define-ext
```

## Use

```
def: laconic, 1, 4
```

Fills 4 rows:

```
laconic
(adjective) Using as few words as possible; pithy and concise.
```

Pass any English word. Definitions are cached for 24 hours.

## Why

Writers, students, and curious people end up with a dictionary tab open all day. A `def:` cell in your QuickSheet wallpaper means a definition is a multi-select-Enter away — no browser, no app switch.

Pairs well with a notes-cells column: drop the word you're looking up, the definition fills in next to it.

## Build

Requires .NET 9. Zero NuGet dependencies — just BCL `System.Net.Http` and `System.Text.Json`.

```
dotnet build DefineExtension.csproj
```

## Notes

- Uses https://dictionaryapi.dev (free, MIT-licensed). No API key.
- Returns one definition per part-of-speech to keep cells compact.
- Cache lives in-memory only — restart the extension to refresh.

## Protocol

Reads JSON lines on stdin, writes JSON lines on stdout. See QuickSheet's main README for the full extension protocol spec.

## License

MIT
