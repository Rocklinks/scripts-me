#!/bin/bash

cat > /tmp/merge_excel.py << 'EOF'
import openpyxl
from copy import copy, deepcopy
import os

source_dirs = {
    "/home/rocklin/sathya/Growth/Cur/": [
        "Branch_Sales.xlsx",
        "growthper.xlsx",
        "Growth.xlsx",
        "sales-ew.xlsx"
    ],
    "/home/rocklin/sathya/Ratio/Cur/": [
        "Ratio.xlsx"
    ],

        "/home/rocklin/sathya/bw_warr/": [
        "Warr.xlsx"
    ]
}

output_file = "/home/rocklin/Downloads/south1.xlsx"

wb_out = openpyxl.Workbook()
wb_out.remove(wb_out.active)

sheet_counter = 1

for source_dir, files in source_dirs.items():
    for fname in files:
        fpath = os.path.join(source_dir, fname)
        if not os.path.exists(fpath):
            print(f"Warning: {fpath} not found, skipping")
            continue
        
        wb_values = openpyxl.load_workbook(fpath, data_only=True)
        wb_format = openpyxl.load_workbook(fpath)
        
        for sheet_name in wb_format.sheetnames:
            ws_val = wb_values[sheet_name]
            ws_fmt = wb_format[sheet_name]
            
            # Use default Sheet1, Sheet2, Sheet3, Sheet4, Sheet5, Sheet6
            new_sheet_name = f"Sheet{sheet_counter}"
            sheet_counter += 1
            
            ws_out = wb_out.create_sheet(title=new_sheet_name)
            
            # Copy column widths EXACTLY from source
            for col_letter, col_dim in ws_fmt.column_dimensions.items():
                ws_out.column_dimensions[col_letter].width = col_dim.width
                ws_out.column_dimensions[col_letter].hidden = col_dim.hidden
                ws_out.column_dimensions[col_letter].bestFit = col_dim.bestFit
            
            # Copy row heights and cell data + formatting
            for row_idx, (row_val, row_fmt) in enumerate(zip(ws_val.iter_rows(), ws_fmt.iter_rows()), 1):
                if row_idx in ws_fmt.row_dimensions:
                    ws_out.row_dimensions[row_idx].height = ws_fmt.row_dimensions[row_idx].height
                    ws_out.row_dimensions[row_idx].hidden = ws_fmt.row_dimensions[row_idx].hidden
                
                for col_idx, (cell_val, cell_fmt) in enumerate(zip(row_val, row_fmt), 1):
                    value = cell_val.value
                    new_cell = ws_out.cell(row=row_idx, column=col_idx, value=value)
                    
                    # Copy formatting
                    if cell_fmt.has_style:
                        new_cell.font = copy(cell_fmt.font)
                        new_cell.border = copy(cell_fmt.border)
                        new_cell.fill = copy(cell_fmt.fill)
                        new_cell.number_format = copy(cell_fmt.number_format)
                        new_cell.protection = copy(cell_fmt.protection)
                        new_cell.alignment = copy(cell_fmt.alignment)
            
            # Copy merged cells
            for merged_range in ws_fmt.merged_cells.ranges:
                ws_out.merge_cells(str(merged_range))
            
            # Copy conditional formatting
            if hasattr(ws_fmt, 'conditional_formatting') and ws_fmt.conditional_formatting._cf_rules:
                ws_out.conditional_formatting = deepcopy(ws_fmt.conditional_formatting)
            
            print(f"Copied {fname} -> {sheet_name} as {new_sheet_name}")

wb_out.save(output_file)
print(f"Created {output_file} with {sheet_counter - 1} sheets")
EOF

python3 /tmp/merge_excel.py
