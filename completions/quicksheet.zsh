#compdef quicksheet ExcelConsole
# zsh completion for QuickSheet
#
# Install:
#   1. Place this file in a directory on your $fpath, named `_quicksheet`, e.g.:
#        mkdir -p ~/.zsh/completions
#        cp completions/quicksheet.zsh ~/.zsh/completions/_quicksheet
#        fpath=(~/.zsh/completions $fpath)
#   2. Reload completions:
#        autoload -Uz compinit && compinit

_quicksheet() {
    local -a flags
    flags=(
        '(--help -h)'{--help,-h}'[Show help]'
        '(--version -v)'{--version,-v}'[Show version]'
        '--list-extensions[List installed extensions]'
        '--export-md[Headless: export CSV to a Markdown table]:output file (- for stdout):_files'
        '--export-html[Headless: export CSV to a styled HTML table]:output file (- for stdout):_files'
        '--export-json[Headless: export CSV to a JSON array]:output file (- for stdout):_files'
    )

    _arguments -s \
        $flags \
        '*:csv file:_files -g "*.csv"'
}

_quicksheet "$@"
