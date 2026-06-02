# quicksheet-grav-ext

Gravatar lookup for [QuickSheet](https://github.com/cemheren/QuickSheet). Given an email, returns the display name, location, and avatar URL.

## Install

```
ext: github:cemheren/quicksheet-grav-ext
```

## Use

```
grav: jane@example.com, 1, 4
```

Fills 4 rows:

```
jane@example.com
Jane Doe
Berlin, Germany
https://www.gravatar.com/avatar/<md5>?d=identicon
```

If the email has no public Gravatar profile, the name row reads `(no Gravatar profile)` and the avatar URL still works (Gravatar falls back to an identicon).

## Why

Useful when you maintain a contacts sheet, a customer list, or a list of email-based aliases and want a quick "is this person real" check. The avatar URL is also handy in Markdown emails or status reports.

## Build

Requires .NET 9. Zero NuGet dependencies — only BCL crypto, HTTP, and JSON.

```
dotnet build GravExtension.csproj
```

## Notes

- Hashes the email with MD5 per the Gravatar spec — that's an externally-mandated algorithm, not a security choice.
- Profile fetch is best-effort. Many emails don't have a public Gravatar profile; that returns 404 and the extension just skips the name/location.
- Results cached for 24 hours in-memory.

## Protocol

Reads JSON lines on stdin, writes JSON lines on stdout. See QuickSheet's main README for the full extension protocol spec.

## License

MIT
