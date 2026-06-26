# bash completion for QuickSheet
#
# Install:
#   source completions/quicksheet.bash
# or copy to a completions dir, e.g.:
#   sudo cp completions/quicksheet.bash /etc/bash_completion.d/quicksheet
#
# Registers for both `quicksheet` and the raw `ExcelConsole` binary name.

_quicksheet()
{
    local cur prev words cword
    if declare -F _init_completion >/dev/null 2>&1; then
        _init_completion || return
    else
        COMPREPLY=()
        cur="${COMP_WORDS[COMP_CWORD]}"
        prev="${COMP_WORDS[COMP_CWORD-1]}"
    fi

    local flags="--help -h --version -v --list-extensions \
--export-md --export-html --export-json"

    # Flags that take an output-file argument next.
    case "$prev" in
        --export-md|--export-html|--export-json)
            # Suggest files (including - for stdout) and directories.
            COMPREPLY=( $(compgen -f -- "$cur") )
            return 0
            ;;
    esac

    if [[ "$cur" == -* ]]; then
        COMPREPLY=( $(compgen -W "$flags" -- "$cur") )
        return 0
    fi

    # Default: complete .csv files and directories.
    local csv_matches
    csv_matches=$(compgen -f -X '!*.csv' -- "$cur")
    local dir_matches
    dir_matches=$(compgen -d -- "$cur")
    COMPREPLY=( $csv_matches $dir_matches )
    return 0
}

complete -F _quicksheet quicksheet
complete -F _quicksheet ExcelConsole
