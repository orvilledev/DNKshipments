# Shipment Box Contents Formatter

Streamlit app that converts a Dansko **Package Content List** export (e.g.
`PMSH01 PO# FBA19HZZ856W.xlsx`) into a two-tab workbook.

Each `Carton No:` block in the export becomes one box, numbered from the
`Carton: N of M` label.

**Tab 1 — Box Contents:** every `Item` in that block, listed against its box.

| UPCs | Box Number | Quantity |
| --- | --- | --- |
| 250780202 | 1 | 1 |
| 3951320202 | 1 | 3 |
| 4144180200 | 1 | 4 |
| ... | ... | ... |

**Tab 2 — Box Dimensions:** the carton's `Dimensions: 31x19x14` label split into
its three numbers, in printed order, one row per box.

| Box Number | Length | Width | Height |
| --- | --- | --- | --- |
| 1 | 31 | 19 | 14 |
| 2 | 31 | 19 | 14 |
| ... | ... | ... | ... |

## Run it

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then upload the `.xlsx` packing list and use **Download Excel (both tabs)**, or
grab either tab on its own as CSV.

## Deploying

`requirements.txt` deliberately uses lower bounds rather than exact pins so the
host installs prebuilt wheels for whichever Python it runs. Exact pins such as
`pandas==2.1.3` have no wheels for Python 3.13, which forces pip to compile
pandas from source and makes a Streamlit Community Cloud deploy take many
minutes or fail outright.

## Options

- **Combine duplicate UPCs per box** (default off) — the packing list has one
  line per size, so the same UPC can appear several times in one carton. Left
  off, you get one row per printed line. Turn it on for one row per UPC per box
  with the quantities added together, which is what Amazon FBA box content
  expects.
- **Show Box Number on the dimensions tab** (default on) — turn it off for a tab
  with only the Length, Width and Height columns.
- **Add a 'Details' sheet** — appends size, description, unit of measure, carton
  number, dimensions, order number and the source row of the original file for
  auditing.

## Checks built in

The app reconciles its output against the totals printed inside the file and
shows the result before you download:

- box count vs `Total # of Cartons:`
- total units vs `Total Quantity:`
- each box's units vs its own `Carton Total:`
- every box has readable dimensions

## Notes

- Every numeric column (`UPCs`, `Box Number`, `Quantity`, `Length`, `Width`,
  `Height`) is written as a real number so Excel can sum and sort without the
  "number stored as text" warning. Item numbers that the packing list prints
  with a leading zero (for example `050020202`) are still stored as numbers,
  with a zero-padded Excel format so they display the way they were printed.
  Alphanumeric codes, when they appear, stay as text.
- All worksheets in the workbook are parsed, and a carton whose items continue
  onto another print page is kept as a single box.
- Dimensions are read as `length x width x height` in printed order. Spaces,
  `X`, `×`, `*`, decimals, comma decimals and trailing units are all accepted,
  and a box whose dimensions cannot be read is left blank and flagged rather
  than guessed.

## Files

- `app.py` — Streamlit UI (upload, preview, validation, downloads).
- `packing_list.py` — parsing and export logic, usable on its own.
- `check_parse.py` — command-line sanity check against a real file:
  `python check_parse.py "path\to\packing list.xlsx"`.
- `check_edge_cases.py` — synthetic workbooks covering page-split cartons,
  multi-sheet exports, mismatched totals and bad input.
- `check_app.py` — renders `app.py` through Streamlit's `AppTest` to confirm the
  full upload-to-download path works.
