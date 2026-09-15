#!/bin/bash

# ============================================================
# Part 1: Absent Manager Detection (pure bash, NO Python)
# ============================================================

find_absent_managers() {
    > 1.txt
    echo "Today App not loggedin managers" > 1.txt
    count=0
    for f in "$@"; do
        branch=$(basename "$f" .csv)
        result=$(awk -F',' '
            tolower($2) ~ /branch manager/ {
                last_ox = ""
                for (i = 6; i <= NF; i++) {
                    val = $i
                    if (val !~ /^[-]+$/) {
                        for (j = length(val); j >= 1; j--) {
                            c = substr(val, j, 1)
                            if (c == "O" || c == "X") { last_ox = c; break }
                        }
                    }
                }
                print last_ox; exit
            }
        ' "$f")
        if [ "$result" = "O" ]; then
            count=$((count + 1))
            echo "$count. $branch" >> 1.txt
        fi
    done
    echo "Absent managers written to 1.txt ($count found)"
}

# ============================================================
# Part 2: Excel Generation (Python)
# ============================================================

CSV_FILES=()
if [ -z "$1" ]; then
    for f in *.csv; do
        [ -f "$f" ] && CSV_FILES+=("$f")
    done
else
    for arg in "$@"; do
        if [ ! -f "$arg" ]; then
            echo "Skipped: $arg (not found)"
            continue
        fi
        CSV_FILES+=("$arg")
    done
fi

if [ ${#CSV_FILES[@]} -eq 0 ]; then
    echo "No CSV files found."
    exit 1
fi

# Sort files by name
IFS=$'\n' SORTED=($(for f in "${CSV_FILES[@]}"; do echo "$f"; done | sort -f)); unset IFS

# Run absent manager detection on all CSV files
find_absent_managers "${SORTED[@]}"

# Generate Excel
if [ ${#SORTED[@]} -eq 1 ]; then
    OUT=$(basename "${SORTED[0]}" .csv)
    OUT="${OUT^}.xlsx"
else
    OUT="All_Branch_Companion.xlsx"
fi

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
python3 "$SCRIPT_DIR/_comp_excel.py" "${SORTED[@]}" "$OUT"

echo "Done. Excel: $OUT | Absent managers: 1.txt"
