# quicksheet-thes-ext

Inline thesaurus for [QuickSheet](https://github.com/cemheren/QuickSheet). Uses the free Datamuse API (no key).

## Install

```
ext: github:cemheren/quicksheet-thes-ext
```

## Use

```
thes: laconic, 1, 6
```

Fills 6 rows:

```
laconic
concise
terse
succinct
pithy
brief
```

Up to 8 synonyms (truncated to fit `gridRows`).

## Why

Pairs nicely with [quicksheet-define-ext](https://github.com/cemheren/quicksheet-define-ext) — definition in one column, synonyms in the next. Writers, editors, anyone with a word-on-the-tip-of-their-tongue gets a wallpaper-level reference panel without opening a browser.

## Build

Requires .NET 9. Zero NuGet dependencies — only BCL.

```
dotnet build ThesExtension.csproj
```

## Notes

- Uses https://api.datamuse.com/words?rel_syn=... (free, no key).
- Results cached in-memory; synonyms don't change.
- Returns the API's "synonyms" relation (`rel_syn`) specifically, not the looser "means like" (`ml`) — fewer but more accurate.

## License

MIT
