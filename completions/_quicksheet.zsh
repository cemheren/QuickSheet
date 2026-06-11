#compdef ExcelConsole quicksheet

# Zsh completion for QuickSheet (ExcelConsole)
# Copy to a directory in your $fpath (e.g. ~/.zsh/completions/)

_quicksheet() {
    local -a opts
    opts=(
        '(-h --help)'{-h,--help}'[Show help]'
        '(-v --version)'{-v,--version}'[Show version]'
        '--list-extensions[List installed extensions]'
        '--export-md[Export CSV to Markdown table]:output file:_files -g "*.md"'
        '--export-html[Export CSV to styled HTML table]:output file:_files -g "*.html"'
        '--export-json[Export CSV to JSON array]:output file:_files -g "*.json"'
        '--info[Show file info (size, last-modified, dimensions)]'
        '--sort[Sort rows by column (ascending)]:column:'
        '--rsort[Sort rows by column (descending)]:column:'
        '(-d --delimiter)'{-d,--delimiter}'[Set field delimiter (default: comma)]:delimiter:'
    )

    _arguments -s $opts \
        '1:CSV file:_files -g "*.csv"'
}

_quicksheet "$@"
