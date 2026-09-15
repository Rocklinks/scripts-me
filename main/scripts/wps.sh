#!/bin/bash
# fix-xlsx — Fixes MIME type + WPS crash for a single xlsx file
# Usage: fix-xlsx file.xlsx

set -euo pipefail

if [ $# -eq 0 ]; then
    echo "Usage: fix-xlsx <file.xlsx>"
    exit 1
fi

FILE="$1"

if [ ! -f "$FILE" ]; then
    echo "Error: $FILE not found"
    exit 1
fi

if [[ ! "$FILE" =~ \.xlsx$ ]]; then
    echo "Error: $FILE is not an .xlsx file"
    exit 1
fi

python3 -c "
import zipfile, sys, shutil, os

filepath = sys.argv[1]

try:
    with zipfile.ZipFile(filepath, 'r') as zin:
        names = zin.namelist()
        if names and names[0] == '[Content_Types].xml':
            print('OK  $FILE — already correct')
            sys.exit(0)
        entries = [(n, zin.read(n)) for n in names]
        content_types = [(n, d) for n, d in entries if n == '[Content_Types].xml']
        others = [(n, d) for n, d in entries if n != '[Content_Types].xml']
        reordered = content_types + others

    tmp = filepath + '.tmp'
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in reordered:
            zout.writestr(name, data)
    shutil.move(tmp, filepath)
    print('FIXED  $FILE')
except Exception as e:
    print(f'ERROR  $FILE: {e}', file=sys.stderr)
    sys.exit(1)
" "$FILE"
