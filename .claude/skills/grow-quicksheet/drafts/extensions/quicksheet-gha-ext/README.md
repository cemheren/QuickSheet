# quicksheet-gha-ext

GitHub Actions workflow status monitor for [QuickSheet](https://github.com/cemheren/QuickSheet). See your CI/CD pipeline health right on your desktop wallpaper — green checks and red crosses at a glance, no browser tab needed.

## Install

In any QuickSheet cell:

```
ext: github:cemheren/quicksheet-gha-ext
```

## Usage

```
gha: owner/repo
gha: owner/repo, 5
gha: owner/repo, 5, main
```

| Parameter   | Description                                     | Default |
|-------------|-------------------------------------------------|---------|
| `owner/repo`| GitHub repository (e.g. `torvalds/linux`)       | —       |
| count       | Number of recent runs to display (1–25)         | 8       |
| branch      | Filter runs to a specific branch                | all     |

## Output

One row per workflow run:

| Column | Content                                        |
|--------|------------------------------------------------|
| 0      | Indicator: ✓ success, ✗ failure, ⚠ cancelled, ⟳ in-progress |
| 1      | Workflow name                                  |
| 2      | Branch                                         |
| 3      | Conclusion (success/failure/cancelled/…)       |
| 4      | Elapsed time                                   |

## Authentication

Set the `GITHUB_TOKEN` environment variable for private repos and higher rate limits (5000 req/hr vs 60). A fine-grained PAT with `actions:read` scope is sufficient.

Without a token, only public repositories are accessible.

## Example

```
gha: cemheren/QuickSheet, 5, main
```

Shows the last 5 workflow runs on the `main` branch:

```
✓  CI          main  success    2m14s
✓  CI          main  success    1m58s
✗  Release     main  failure    4m02s
✓  CI          main  success    2m01s
⟳  CI          main  in_progress  —
```

## Why on a wallpaper?

- **SRE dashboard**: pair with `tls:`, `health:`, `docker:` extensions for a full ops-at-a-glance desktop.
- **Open-source maintainer**: see PR check status without opening GitHub.
- **Team lead**: monitor multiple repos — one `gha:` cell per project.

## Build

```bash
dotnet build GhaExtension.csproj
```

.NET 9, MIT, zero NuGet dependencies.

## Protocol

Standard QuickSheet JSON-lines:

1. On startup, emit `{"type":"register","prefix":"gha","name":"GitHub Actions","version":"1.0.0"}`.
2. On each `{"type":"activate","id":"...","params":["owner/repo, 5, main"]}`, fetch workflow runs from the GitHub API and reply with `{"type":"write","id":"...","cells":[...]}`.

See [docs/extension-protocol.md](https://github.com/cemheren/QuickSheet/blob/main/docs/extension-protocol.md).

## License

MIT.
