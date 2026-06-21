import openpyxl

wb = openpyxl.load_workbook("referensi/SWARA_MAIRCA_DETAIL_RUMUS.xlsx", data_only=True)
print("Sheets in workbook:", wb.sheetnames)
for sheetname in wb.sheetnames:
    ws = wb[sheetname]
    print(f"\nSheet: {sheetname}")
    print(f"Max row: {ws.max_row}, Max col: {ws.max_column}")
    # Print the first 5 rows and 5 columns
    for r in range(1, min(10, ws.max_row + 1)):
        row_vals = [ws.cell(row=r, column=c).value for c in range(1, min(10, ws.max_column + 1))]
        print(f"Row {r}: {row_vals}")
