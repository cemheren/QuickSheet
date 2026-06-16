# Shell Completions

QuickSheet ships tab-completion scripts for **bash**, **zsh**, and **fish**.
They complete all CLI flags and default to `.csv` files for positional arguments.

## Bash

```bash
# Option 1: Source in current session
source completions/quicksheet.bash

# Option 2: Install system-wide
sudo cp completions/quicksheet.bash /etc/bash_completion.d/quicksheet
```

## Zsh

```zsh
# Copy to a directory in your fpath
mkdir -p ~/.zsh/completions
cp completions/quicksheet.zsh ~/.zsh/completions/_quicksheet

# Ensure ~/.zsh/completions is in fpath (add to .zshrc):
fpath=(~/.zsh/completions $fpath)
autoload -Uz compinit && compinit
```

## Fish

```fish
cp completions/quicksheet.fish ~/.config/fish/completions/quicksheet.fish
```

## What's completed

| Token | Completes to |
|-------|-------------|
| `--` | All long flags (`--help`, `--export-md`, `--delimiter`, etc.) |
| `-` | Short flags (`-h`, `-v`, `-d`) |
| After `--export-*` | File paths |
| Positional args | `.csv` files |
