# Fish completion for QuickSheet (ExcelConsole)
# Install: copy to ~/.config/fish/completions/quicksheet.fish

complete -c quicksheet -s h -l help -d 'Show help and usage examples'
complete -c quicksheet -s v -l version -d 'Show version information'
complete -c quicksheet -l list-extensions -d 'List installed extensions'
complete -c quicksheet -l export-md -r -F -d 'Export CSV to Markdown table'
complete -c quicksheet -l export-html -r -F -d 'Export CSV to styled HTML table'
complete -c quicksheet -l export-json -r -F -d 'Export CSV to JSON array'
complete -c quicksheet -l export-tsv -r -F -d 'Export CSV to TSV'
complete -c quicksheet -s d -l delimiter -x -d 'Set field delimiter for headless export'
complete -c quicksheet -l info -d 'Show file size and last-modified date'

# Also register for ExcelConsole binary name
complete -c ExcelConsole -w quicksheet
