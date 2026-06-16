#compdef quicksheet ExcelConsole

# Zsh completion for QuickSheet (ExcelConsole)
# Install: copy to a directory in your $fpath (e.g. ~/.zsh/completions/)
# and rename to _quicksheet, then run: autoload -Uz compinit && compinit

_quicksheet() {
    local -a flags

    flags=(
        '(-h --help)'{-h,--help}'[Show help and usage examples]'
        '(-v --version)'{-v,--version}'[Show version information]'
        '--list-extensions[List installed extensions]'
        '--export-md[Export CSV to Markdown table]:output file:_files'
        '--export-html[Export CSV to styled HTML table]:output file:_files'
        '--export-json[Export CSV to JSON array]:output file:_files'
        '--export-tsv[Export CSV to TSV]:output file:_files'
        '(-d --delimiter)'{-d,--delimiter}'[Set field delimiter for headless export]:delimiter:'
        '--info[Show file size and last-modified date]'
    )

    _arguments -s $flags \
        '*:CSV file:_files -g "*.csv"'
}

_quicksheet "$@"
