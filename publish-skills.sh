#!/bin/sh
set -eu

script_directory=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)

if command -v python3 >/dev/null 2>&1; then
    exec python3 "$script_directory/publish-skills.py" "$@"
fi
if command -v python >/dev/null 2>&1; then
    exec python "$script_directory/publish-skills.py" "$@"
fi

printf '%s\n' 'Python 3 is required. Install Python 3 and run this command again.' >&2
exit 1
