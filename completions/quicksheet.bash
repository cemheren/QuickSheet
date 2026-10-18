# Bash completion for QuickSheet (ExcelConsole)
# Install: source this file, or copy to /etc/bash_completion.d/quicksheet

_quicksheet() {
    local cur prev opts
    COMPREPLY=()
    cur="${COMP_WORDS[COMP_CWORD]}"
    prev="${COMP_WORDS[COMP_CWORD-1]}"

    # Flags that take a file argument
    case "$prev" in
        --export-md|--export-html|--export-json|--export-tsv)
            COMPREPLY=( $(compgen -f -- "$cur") )
            return 0
            ;;
        --delimiter|-d)
            # User provides a delimiter character; no completion
            return 0
            ;;
    esac

    # Complete flags
    if [[ "$cur" == --* ]]; then
        opts="--help --version --list-extensions --export-md --export-html --export-json --export-tsv --delimiter --info"
        COMPREPLY=( $(compgen -W "$opts" -- "$cur") )
        return 0
    fi

    if [[ "$cur" == -* ]]; then
        opts="-h -v -d"
        COMPREPLY=( $(compgen -W "$opts" -- "$cur") )
        return 0
    fi

    # Default: complete CSV files and directories
    COMPREPLY=( $(compgen -f -X '!*.csv' -- "$cur") $(compgen -d -- "$cur") )
    return 0
}

complete -F _quicksheet quicksheet
complete -F _quicksheet ExcelConsole
