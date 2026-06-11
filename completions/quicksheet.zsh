#compdef ExcelConsole quicksheet

_quicksheet() {
  local -a opts
  opts=(
    '--help[Show help]'
    '-h[Show help]'
    '--version[Show version]'
    '-v[Show version]'
    '--list-extensions[List installed extensions]'
    '--export-md[Export CSV to Markdown]:output file:_files'
    '--export-html[Export CSV to HTML]:output file:_files'
    '--export-json[Export CSV to JSON]:output file:_files'
  )

  _arguments -s \
    '(-h --help)'{-h,--help}'[Show help]' \
    '(-v --version)'{-v,--version}'[Show version]' \
    '--list-extensions[List installed extensions]' \
    '--export-md[Export CSV to Markdown]:output file:_files' \
    '--export-html[Export CSV to HTML]:output file:_files' \
    '--export-json[Export CSV to JSON]:output file:_files' \
    '*:CSV file:_files -g "*.csv"'
}
