#!/bin/sh
cd "$(dirname "$0")" || exit 1
python3 ibr_tfvars_generator.py "$@"
rc=$?
if [ "$rc" -ne 0 ]; then
  printf '\nThe generator exited with code %s.\n' "$rc"
  printf 'Press Enter to close...'
  read _
fi
exit "$rc"
