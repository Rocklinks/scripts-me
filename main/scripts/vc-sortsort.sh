#!/bin/bash
# sort_keep_formulas.sh - sort every table on every sheet descending by
# its % (or ACH%, Ach %, achieved%... any spelling/case - it just looks
# for the '%' symbol) column. Formulas are never deleted or baked into
# numbers: a row's own formulas move WITH it, and only a formula's
# reference to ITS OWN row gets bumped to the new row number (exactly
# what Excel itself does when you sort/cut-paste a row). Formulas
# anchored with $ (fixed ranges, external-workbook lookups, SUMIF
# tables) are left untouched since they don't depend on row position.
#
# Not tied to this or that file's specific formulas - works on any
# sheet where each row's formulas only look at cells in their own row
# or a $-anchored fixed range, which is normal spreadsheet practice.
#
# usage: ./sort_keep_formulas.sh file.xlsx [output.xlsx]
#   - if output.xlsx is omitted, file.xlsx is overwritten in place.

set -e

IN="$1"
OUT="${2:-$IN}"

[ -z "$IN" ] && { echo "usage: $0 file.xlsx [output.xlsx]"; exit 1; }
[ -f "$IN" ] || { echo "not found: $IN"; exit 1; }

python3 - "$IN" "$OUT" << 'PYEOF'
import sys, re
import openpyxl

src, dst = sys.argv[1], sys.argv[2]

wb_val = openpyxl.load_workbook(src, data_only=True)   # cached values, used only to find headers / sort key
wb = openpyxl.load_workbook(src, data_only=False)      # formulas, edited in place & saved

REF_RE = re.compile(r'(\$?)([A-Z]{1,3})(\$?)(\d+)')

def shift_formula(value, old_row, new_row):
    """Bump only UNANCHORED (no $) references to old_row -> new_row.
    $-anchored refs (fixed ranges, external lookup tables) are left as-is."""
    if not isinstance(value, str) or not value.startswith("="):
        return value

    def repl(m):
        dollar_col, col, dollar_row, row = m.groups()
        if dollar_row == "" and int(row) == old_row:
            return f"{dollar_col}{col}{dollar_row}{new_row}"
        return m.group(0)

    return REF_RE.sub(repl, value)

def to_number(x, default=-1):
    try:
        return float(x)
    except (TypeError, ValueError):
        return default

def find_blocks(ws_val):
    """A header row = text in col A + some cell containing '%' (any
    case/spelling: %, ACH%, Ach %, Achieved%...). Data rows are every
    row below with a non-empty col A. Returns (rows, pct_col) for every
    such block on the sheet, in top-to-bottom order."""
    blocks = []
    maxcol = ws_val.max_column
    maxrow = ws_val.max_row
    r = 1
    while r <= maxrow:
        row_vals = [ws_val.cell(row=r, column=c).value for c in range(1, maxcol + 1)]
        a = row_vals[0]
        if isinstance(a, str) and a.strip():
            pct_col = None
            for idx, v in enumerate(row_vals, start=1):
                if isinstance(v, str) and "%" in v:
                    pct_col = idx
                    break
            if pct_col is not None:
                rr = r + 1
                rows = []
                while rr <= maxrow and ws_val.cell(row=rr, column=1).value not in (None, ""):
                    rows.append(rr)
                    rr += 1
                blocks.append((rows, pct_col))
                r = rr
                continue
        r += 1
    return blocks

for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    ws_val = wb_val[sheet_name]
    print(f"sheet: {sheet_name}")

    for rows, pct_col in find_blocks(ws_val):
        if not rows:
            continue
        maxcol = ws.max_column

        snap = []
        for r in rows:
            pct = to_number(ws_val.cell(row=r, column=pct_col).value)
            row_content = [ws.cell(row=r, column=c).value for c in range(1, maxcol + 1)]
            snap.append((pct, r, row_content))
        snap.sort(key=lambda x: x[0], reverse=True)

        for new_r, (pct, old_r, row_content) in zip(rows, snap):
            for c, v in enumerate(row_content, start=1):
                ws.cell(row=new_r, column=c).value = shift_formula(v, old_r, new_r)
        print(f"  sorted {len(rows)} rows by column {openpyxl.utils.get_column_letter(pct_col)}")

wb.save(dst)
print(f"saved: {dst}")
PYEOF
