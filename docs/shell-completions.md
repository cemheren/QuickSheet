# Shell completions

QuickSheet ships tab-completion scripts for **bash** and **zsh** under
[`completions/`](../completions). They complete CLI flags
(`--help`, `--version`, `--list-extensions`, `--export-md`, `--export-html`,
`--export-json`) and suggest `.csv` files for the input argument and output
files (including `-` for stdout) after each `--export-*` flag.

The scripts register for both the `quicksheet` command and the raw
`ExcelConsole` binary name, so completion works regardless of how you alias it.

## bash

Quick try (current shell only):

```bash
source completions/quicksheet.bash
```

Persistent — copy into a completion directory:

```bash
# system-wide
sudo cp completions/quicksheet.bash /etc/bash_completion.d/quicksheet

# or per-user (then source it from ~/.bashrc)
mkdir -p ~/.local/share/bash-completion/completions
cp completions/quicksheet.bash ~/.local/share/bash-completion/completions/quicksheet
```

Open a new shell (or re-`source ~/.bashrc`) and press `<Tab>`:

```bash
quicksheet --ex<Tab>        # → --export-md / --export-html / --export-json
quicksheet data.<Tab>       # → completes *.csv files
quicksheet data.csv --export-md <Tab>   # → file completion for the output
```

## zsh

Place the script on your `$fpath` named `_quicksheet`, then reload completions:

```zsh
mkdir -p ~/.zsh/completions
cp completions/quicksheet.zsh ~/.zsh/completions/_quicksheet
# add to ~/.zshrc (before compinit):
#   fpath=(~/.zsh/completions $fpath)
autoload -Uz compinit && compinit
```

Then:

```zsh
quicksheet --<Tab>          # lists all flags with descriptions
quicksheet --export-json <Tab>   # file completion for the output path
```

> Tip: if you invoke the build output directly as `ExcelConsole`, completion
> works for that name too — both scripts register both command names.
