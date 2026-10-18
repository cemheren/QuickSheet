# quicksheet-gomod

[![QuickSheet Extension](https://img.shields.io/badge/QuickSheet-Extension-blue)](https://github.com/cemheren/QuickSheet)

A [QuickSheet](https://github.com/cemheren/QuickSheet) extension that shows **Go module info** directly in your spreadsheet cells — the latest version, publish date, and number of published versions — straight from the public Go module proxy.

## Usage

Type in any cell using the `gomod:` prefix:

```
gomod: github.com/gin-gonic/gin
gomod: gorm.io/gorm
gomod: github.com/gin-gonic/gin,gorm.io/gorm,github.com/spf13/cobra
```

**Single module** — shows a 2-column info card:

| Field    | Value                          |
|----------|--------------------------------|
| Module   | gin-gonic/gin                  |
| Version  | v1.10.0                        |
| Released | 2024-05-07                     |
| Versions | 39                             |
| github.com/gin-gonic/gin       |              |

**Multiple modules** — shows a comparison table (up to 20):

| Module          | Version | Released   |
|-----------------|---------|------------|
| gin-gonic/gin   | v1.10.0 | 2024-05-07 |
| gorm.io/gorm    | v1.25.x | 2024-xx-xx |
| spf13/cobra     | v1.8.x  | 2024-xx-xx |

## Install

```
ext: github:Deskworks/quicksheet-gomod
```

Requires [QuickSheet](https://github.com/cemheren/QuickSheet) and the [.NET 9 SDK](https://dotnet.microsoft.com/download).

## Features

- 🐹 Latest version + publish date for any Go module path
- 📊 Multi-module comparison table (comma-separated)
- 🔢 Count of published versions
- ⚡ 30-minute cache — no repeated proxy calls
- 🔑 No API key required — uses the public [Go module proxy](https://proxy.golang.org/)
- 🚫 Zero NuGet dependencies

## Data source

- `https://proxy.golang.org/{module}/@latest` — latest version + timestamp
- `https://proxy.golang.org/{module}/@v/list` — published version list

Both are free, unauthenticated, public endpoints. Module paths are
[case-encoded](https://go.dev/ref/mod#goproxy-protocol) per the Go module
proxy protocol (uppercase letters become `!` + lowercase).

## Part of the QuickSheet extension ecosystem

Sits alongside the other package-registry extensions (`npm:`, `pypi:`,
`crates:`, `nuget:`, `maven:`, `rubygems:`, `hackage:`). See the
[full extension directory](https://github.com/cemheren/QuickSheet#extensions).

## License

MIT
