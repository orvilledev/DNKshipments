"""Ad-hoc verification of the parser against a real Package Content List file."""

import sys

from packing_list import (
    build_dimensions,
    build_output,
    parse_packing_list,
    to_excel_bytes,
)

path = sys.argv[1] if len(sys.argv) > 1 else r"C:\Users\Administrator\Downloads\PMSH01 PO# FBA19HZZ856W.xlsx"

parsed = parse_packing_list(path)
print("orders:", parsed.order_numbers)
print("cartons parsed:", len(parsed.cartons), "reported:", parsed.reported_carton_count)
print("qty parsed:", parsed.total_quantity, "reported:", parsed.reported_total_quantity)
print("item lines:", len(parsed.lines))
print("warnings:", parsed.warnings)

print("\nbox numbers:", [c.box_number for c in parsed.cartons])

raw = build_output(parsed, combine_duplicates=False)
combined = build_output(parsed, combine_duplicates=True)
print("\nrows raw:", len(raw), "rows combined:", len(combined))
print("qty raw:", raw["Quantity"].sum(), "qty combined:", combined["Quantity"].sum())

print("\n-- box 1 raw --")
print(raw[raw["Box Number"] == 1].to_string(index=False))
print("\n-- box 1 combined --")
print(combined[combined["Box Number"] == 1].to_string(index=False))

print("\n-- box 2 raw (leading-zero check) --")
print(raw[raw["Box Number"] == 2].to_string(index=False))

print("\n-- box 17 raw (leading-zero check row 243) --")
print(raw[raw["Box Number"] == 17].to_string(index=False))

print("\n-- last box --")
print(raw[raw["Box Number"] == raw["Box Number"].max()].to_string(index=False))

print("\ndetail columns:", list(parsed.lines.columns))
print(parsed.lines.head(3).to_string(index=False))

dimensions = build_dimensions(parsed)
print("\n-- box dimensions --")
print("columns:", list(dimensions.columns), "dtypes:", dimensions.dtypes.to_dict())
print(dimensions.head(3).to_string(index=False))
print("rows:", len(dimensions))
print("unique L/W/H:", set(map(tuple, dimensions[["Length", "Width", "Height"]].values)))
assert len(dimensions) == 28
assert list(dimensions.columns) == ["Box Number", "Length", "Width", "Height"]
assert list(dimensions.loc[0, ["Length", "Width", "Height"]]) == [31, 19, 14]

no_box = build_dimensions(parsed, include_box_number=False)
assert list(no_box.columns) == ["Length", "Width", "Height"]

blob = to_excel_bytes(raw, dimensions, parsed.lines)
print("\nexcel bytes:", len(blob))
with open("_verify_output.xlsx", "wb") as fh:
    fh.write(blob)

import openpyxl

wb = openpyxl.load_workbook("_verify_output.xlsx")
print("sheets:", wb.sheetnames)
assert wb.sheetnames == ["Box Contents", "Box Dimensions", "Details"]
for name in ("Box Contents", "Box Dimensions"):
    ws = wb[name]
    print(f"\n[{name}] dims:", ws.dimensions)
    for row in ws.iter_rows(min_row=1, max_row=5, values_only=True):
        print(row)

ws = wb["Box Dimensions"]
assert ws["A1"].value == "Box Number"
assert [ws["B1"].value, ws["C1"].value, ws["D1"].value] == ["Length", "Width", "Height"]
assert [ws["B2"].value, ws["C2"].value, ws["D2"].value] == [31, 19, 14]
assert all(isinstance(ws[f"{c}2"].value, int) for c in "BCD"), "L/W/H must be numeric"
assert ws.max_row == 29, ws.max_row
