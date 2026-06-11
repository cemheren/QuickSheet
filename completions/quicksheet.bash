# Bash completion for QuickSheet (ExcelConsole)
# Source this file or copy to /etc/bash_completion.d/quicksheet

_quicksheet() {
    local cur prev opts
    COMPREPLY=()
    cur="${COMP_WORDS[COMP_CWORD]}"
    prev="${COMP_WORDS[COMP_CWORD-1]}"

    opts="--help --version --list-extensions --export-md --export-html --export-json --info --sort --rsort --delimiter -h -v -d"

    case "$prev" in
        --export-md)
            COMPREPLY=( $(compgen -f -X '!*.md' -- "$cur") )
            return 0
            ;;
        --export-html)
            COMPREPLY=( $(compgen -f -X '!*.html' -- "$cur") )
            return 0
            ;;
        --export-json)
            COMPREPLY=( $(compgen -f -X '!*.json' -- "$cur") )
            return 0
            ;;
        --sort|--rsort)
            # Column name or number — no completion available
            return 0
            ;;
        --delimiter|-d)
            # Single character — no completion available
            return 0
            ;;
    esac

    if [[ "$cur" == -* ]]; then
        COMPREPLY=( $(compgen -W "$opts" -- "$cur") )
        return 0
    fi

    # Default: complete CSV files
    COMPREPLY=( $(compgen -f -X '!*.csv' -- "$cur") )
    return 0
}

complete -o filenames -F _quicksheet ExcelConsole
complete -o filenames -F _quicksheet quicksheet
