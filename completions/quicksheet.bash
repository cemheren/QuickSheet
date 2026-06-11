_quicksheet()
{
    local cur prev opts
    COMPREPLY=()
    cur="${COMP_WORDS[COMP_CWORD]}"
    prev="${COMP_WORDS[COMP_CWORD-1]}"

    case "$prev" in
        --export-md|--export-html|--export-json)
            COMPREPLY=( $(compgen -f -- "$cur") )
            return 0
            ;;
    esac

    opts="--help -h --version -v --list-extensions --export-md --export-html --export-json"
    if [[ "$cur" == -* ]]; then
        COMPREPLY=( $(compgen -W "$opts" -- "$cur") )
    else
        COMPREPLY=( $(compgen -f -- "$cur") )
    fi
}

complete -F _quicksheet ExcelConsole quicksheet
