# quicksheet-gha-ext

GitHub Actions workflow status dashboard for [QuickSheet](https://github.com/cemheren/QuickSheet).

Monitor your CI/CD pipelines at a glance — right on your desktop wallpaper.

## Install

In any QuickSheet cell:

```
ext: github:cemheren/quicksheet-gha-ext
```

## Usage

```
gha: owner/repo
gha: owner/repo, 5
```

| Column | Content |
|--------|---------|
| 1 | Status indicator (✓ ✗ ⚠ ◉ ○) |
| 2 | Workflow name |
| 3 | Branch |
| 4 | Duration |

### Indicators

| Symbol | Meaning |
|--------|---------|
| ✓ | Success |
| ✗ | Failure |
| ⚠ | Cancelled |
| ◉ | In progress |
| ○ | Queued / Skipped |

## Authentication

Works unauthenticated for public repos (60 requests/hour).

Set `GITHUB_TOKEN` environment variable for private repos and higher rate limits (5000 req/hr).

## Examples

```
gha: torvalds/linux, 3
gha: dotnet/runtime
gha: myorg/my-private-repo, 10
```

## Screenshot

<!-- TODO: add screenshot showing CI status grid on desktop -->

## Requirements

- .NET 9 SDK
- QuickSheet v0.35+

## License

MIT
