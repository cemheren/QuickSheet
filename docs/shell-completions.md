# Shell Completions

QuickSheet ships shell completion helpers for `bash` and `zsh`.

## Bash

```bash
source /path/to/QuickSheet/completions/quicksheet.bash
```

This registers completions for `ExcelConsole` and `quicksheet`.

## Zsh

```bash
source /path/to/QuickSheet/completions/quicksheet.zsh
```

This registers completions for `ExcelConsole` and `quicksheet`.

## What completes

- `--help`, `-h`
- `--version`, `-v`
- `--list-extensions`
- `--export-md`
- `--export-html`
- `--export-json`
- CSV file paths by default

For export commands, the next argument is completed as a file path.
