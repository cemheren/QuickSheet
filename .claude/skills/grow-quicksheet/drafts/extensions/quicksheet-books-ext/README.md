# quicksheet-books-ext

Book lookup extension for [QuickSheet](https://github.com/cemheren/QuickSheet) — search titles, look up ISBNs, browse authors. Powered by [Open Library](https://openlibrary.org/) (free, no API key).

## Install

Type into any QuickSheet cell:

```
ext: github:cemheren/quicksheet-books-ext
```

## Usage

```
books: Dune, 1, 5
```

Displays title, author, year, publisher, and page count for the best match.

### Search multiple results

```
books: search functional programming, 2, 5
```

Returns a table of matching books (title, author, year).

### ISBN lookup

```
books: 978-0-13-468599-1, 1, 5
```

Direct lookup by ISBN-10 or ISBN-13.

## Examples

| Cell | What you get |
|------|-------------|
| `books: The Pragmatic Programmer, 1, 5` | Detail card — title, author, year, publisher, pages |
| `books: search Tolkien, 2, 8` | Table of Tolkien's works |
| `books: 0-596-51774-1, 1, 5` | ISBN lookup (JavaScript: The Good Parts) |

## Screenshot

<!-- TODO: Add screenshot of extension running in QuickSheet desktop mode -->

## How it works

Uses the [Open Library Search API](https://openlibrary.org/dev/docs/api/search) — completely free, no registration, no rate-limit key. Results are cached in-memory for the session.

## Requirements

- .NET 9 SDK
- QuickSheet with extension support

## License

MIT
