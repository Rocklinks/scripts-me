#!/usr/bin/env bash
# fix_growth_region.sh — ONE job only: fix Growth.xlsx's region rollup
# formulas (D:L columns) when they end up with wrong ranges again
# (e.g. A3:A39, A6:A42, instead of the correct A4:A40).
#
# Matches each region row by its NAME in column A (VNR1, Tirunelveli, KVT1,
# Tuticorin, NGR1, TKS1) — not by row number — so it works even if the
# region rows or branch rows have been reordered/sorted.
#
# Usage:
#   ./fix_growth_region.sh Growth.xlsx
#   ./fix_growth_region.sh /path/to/Growth.xlsx

set -euo pipefail

IN="${1:-Growth.xlsx}"
if [[ ! -f "$IN" ]]; then
    echo "ERROR: file not found: $IN"
    exit 1
fi
ABS=$(cd "$(dirname "$IN")" && pwd)/$(basename "$IN")

command -v python3 >/dev/null || { echo "ERROR: python3 not found."; exit 1; }
python3 -c "import openpyxl" 2>/dev/null || { echo "ERROR: run: pip install --user openpyxl"; exit 1; }
command -v zip >/dev/null || { echo "ERROR: 'zip' not found."; exit 1; }
command -v unzip >/dev/null || { echo "ERROR: 'unzip' not found."; exit 1; }

WORK=$(mktemp -d)
trap 'rm -rf "$WORK"' EXIT

python3 - "$ABS" <<'PYEOF'
import re, os, sys
import openpyxl

path = sys.argv[1]
work = os.environ.get("WORK") or __import__("tempfile").mkdtemp()

def extract(xlsx_path, workdir):
    d = os.path.join(workdir, os.path.basename(xlsx_path) + ".d")
    os.makedirs(d, exist_ok=True)
    os.system(f'unzip -q -o "{xlsx_path}" -d "{d}"')
    return d

def repack(extract_dir, out_path):
    # out_path MUST be absolute — we cd into extract_dir before zipping
    os.system(f'cd "{extract_dir}" && rm -f "{out_path}" && zip -q -r -X "{out_path}" .')

d = extract(path, work)
sx = os.path.join(d, "xl", "worksheets", "sheet1.xml")
with open(sx, encoding="utf-8") as f:
    content = f.read()

wb = openpyxl.load_workbook(path, data_only=True)
ws = wb["Sheet1"]
cols = "DEFGHIJKL"

branches = {}
for r in range(4, 41):
    name = ws[f"A{r}"].value
    if name is None:
        continue
    branches[name] = {c: ws[f"{c}{r}"].value for c in cols}

region_branches = {
    "VNR1": ["Virudhunagar", "Virudhunagar-2", "Aruppukottai", "Aruppukottai -2", "Sivakasi"],
    "Tirunelveli": ["Tirunelveli-1", "Valliyur-1", "Ambasamudram-1", "Anjugramam-1"],
    "KVT1": ["Kovilpatti", "Ramnad", "Ramnad-2", "Paramakudi", "Sayalkudi-1",
             "Villathikullam", "Sattur-2", "Sankarankovil-1", "Kayathar-1"],
    "Tuticorin": ["Tuticorin-1", "Tuticorin-2", "Thiruchendur-1", "Thisayanvilai-1", "Eral-2", "Udankudi"],
    "NGR1": ["Nagercoil", "Marthandam", "Thuckalay-1", "Colachel-1", "Kulasekharam-1",
             "Monday Market", "Karungal-1"],
    "TKS1": ["Thenkasi", "Thenkasi-2", "Surandai-1", "Puliyankudi-1", "Rajapalayam", "Sengottai-1"],
}

region_rows, header_row, total_row = {}, None, None
for r in range(1, ws.max_row + 1):
    if ws[f"A{r}"].value == "Region":
        header_row = r
    if ws[f"A{r}"].value == "Total" and header_row is not None and r > header_row:
        total_row = r
        break

if header_row is None or total_row is None:
    print("ERROR: could not find 'Region' header row / 'Total' row in column A.")
    sys.exit(1)

for r in range(header_row + 1, total_row):
    name = ws[f"A{r}"].value
    if name in region_branches:
        region_rows[r] = name
    elif name is not None:
        print(f"WARN: row {r} has unrecognized region name {name!r} — skipped")

sums = {}
for row, region_name in region_rows.items():
    names = region_branches[region_name]
    sums[row] = {
        c: sum(branches[n][c] for n in names if n in branches and isinstance(branches[n][c], (int, float)))
        for c in cols
    }

for row, region_name in region_rows.items():
    names = region_branches[region_name]
    names_xml = ",".join(f"&quot;{n}&quot;" for n in names)
    for c in cols:
        newval = sums[row][c]
        newformula = f"SUM(SUMIF(A4:A40,{{{names_xml}}},{c}4:{c}40))"
        pattern = re.compile(rf'(<c r="{c}{row}"[^>]*>)<f[^>]*>.*?</f><v>.*?</v>(</c>)', re.S)
        m = pattern.search(content)
        if m:
            repl = f'{m.group(1)}<f>{newformula}</f><v>{newval}</v>{m.group(2)}'
            content = content[: m.start()] + repl + content[m.end():]
        else:
            print(f"WARN: cell {c}{row} not found/pattern mismatch — left as-is")

grand = {c: sum(sums[r][c] for r in sums) for c in cols}
for c in cols:
    newval = grand[c]
    pattern = re.compile(rf'<c r="{c}{total_row}"([^>]*)>(<f[^/]*?/>|<f[^>]*>[^<]*</f>)<v>[^<]*</v></c>')
    m = pattern.search(content)
    if m:
        repl = f'<c r="{c}{total_row}"{m.group(1)}>{m.group(2)}<v>{newval}</v></c>'
        content = content[: m.start()] + repl + content[m.end():]
    else:
        print(f"WARN: total cell {c}{total_row} not found/pattern mismatch — left as-is")

with open(sx, "w", encoding="utf-8") as f:
    f.write(content)
repack(d, path)

# sanity check — find the TOP table's "Total" row independently (it's the
# last row above the "Region" header whose column A says "Total")
top_row = None
for r in range(1, header_row):
    if ws[f"A{r}"].value == "Total":
        top_row = r

wb2 = openpyxl.load_workbook(path, data_only=True)
ws2 = wb2["Sheet1"]
print(f"Fixed region rows {sorted(region_rows)} (region total row {total_row})")
if top_row:
    top = [ws2[f"{c}{top_row}"].value for c in cols]
    reg = [ws2[f"{c}{total_row}"].value for c in cols]
    print(f"Top total (row {top_row}):    {top}")
    print(f"Region total (row {total_row}): {reg}")
    print("MATCH:", top == reg)
else:
    print("(could not locate top table's Total row for a sanity check — verify manually)")
PYEOF
