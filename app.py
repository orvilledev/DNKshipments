"""Streamlit app: turn a Package Content List export into a box-contents table."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import streamlit as st

from packing_list import (
    PackingListError,
    ParsedPackingList,
    build_dimensions,
    build_output,
    parse_packing_list,
    to_excel_bytes,
    upc_display_widths,
)

st.set_page_config(page_title="Shipment Box Contents Formatter", page_icon="📦", layout="wide")


@st.cache_data(show_spinner=False)
def _parse(file_bytes: bytes) -> ParsedPackingList:
    return parse_packing_list(file_bytes)


def _validation(parsed: ParsedPackingList, output: pd.DataFrame) -> list[tuple[str, str]]:
    """Reconcile the parsed result against the totals printed in the file."""
    checks: list[tuple[str, str]] = []

    boxes = int(output["Box Number"].nunique())
    if parsed.reported_carton_count is None:
        checks.append(("warning", f"{boxes} boxes found (the file has no carton total to check against)."))
    elif boxes == parsed.reported_carton_count:
        checks.append(("ok", f"Box count matches the file: {boxes} of {parsed.reported_carton_count} cartons."))
    else:
        checks.append((
            "error",
            f"Box count mismatch: {boxes} boxes parsed but the file reports "
            f"{parsed.reported_carton_count} cartons.",
        ))

    quantity = int(output["Quantity"].sum())
    if parsed.reported_total_quantity is None:
        checks.append(("warning", f"{quantity} units total (the file has no quantity total to check against)."))
    elif quantity == parsed.reported_total_quantity:
        checks.append(("ok", f"Quantity matches the file: {quantity} units."))
    else:
        checks.append((
            "error",
            f"Quantity mismatch: {quantity} units parsed but the file reports "
            f"{parsed.reported_total_quantity}.",
        ))

    mismatched = [
        carton
        for carton in parsed.cartons
        if carton.reported_total is not None and carton.reported_total != carton.parsed_total
    ]
    if mismatched:
        checks.append((
            "error",
            "Per-box totals disagree with the file for box(es): "
            + ", ".join(str(carton.box_number) for carton in mismatched),
        ))
    else:
        checks.append(("ok", "Every box matches its printed 'Carton Total'."))

    missing_dimensions = [c.box_number for c in parsed.cartons if c.length is None]
    if missing_dimensions:
        checks.append((
            "warning",
            "No dimensions could be read for box(es): "
            + ", ".join(str(number) for number in missing_dimensions),
        ))
    else:
        checks.append(("ok", f"Dimensions read for all {len(parsed.cartons)} boxes."))

    return checks


st.title("📦 Shipment Box Contents Formatter")
st.write(
    "Upload a **Package Content List** export and get back a workbook with a "
    "**UPCs / Box Number / Quantity** tab and a **Length / Width / Height** tab, "
    "with each carton numbered as a box."
)

with st.sidebar:
    st.header("Options")
    combine_duplicates = st.toggle(
        "Combine duplicate UPCs per box",
        value=False,
        help=(
            "Off: one row per line exactly as printed on the packing list. On: one "
            "row per UPC per box, with sizes of the same UPC added together."
        ),
    )
    dimensions_box_number = st.toggle(
        "Show Box Number on the dimensions tab",
        value=True,
        help="Turn off for a tab with only the Length, Width and Height columns.",
    )
    include_details = st.toggle(
        "Add a 'Details' sheet to the Excel download",
        value=False,
        help="Includes size, description, carton number and source row for auditing.",
    )
    st.caption("Box numbers come from the 'Carton: N of M' label in the file.")

uploaded = st.file_uploader(
    "Upload the packing list (.xlsx)",
    type=["xlsx", "xlsm"],
    help="This is the 'Package Content List' export, e.g. PMSH01 PO# FBA19HZZ856W.xlsx",
)

if uploaded is None:
    st.info("Upload a file to get started.")
    with st.expander("What the app produces"):
        st.markdown(
            "**Tab 1 — Box Contents**\n"
            "- **UPCs** — the `Item` number from the packing list, kept exactly as "
            "printed (including any leading zeros).\n"
            "- **Box Number** — 1 for the first carton, 2 for the second, and so on.\n"
            "- **Quantity** — the units of that UPC in that box.\n\n"
            "**Tab 2 — Box Dimensions**\n"
            "- **Length / Width / Height** — split from each carton's "
            "`Dimensions: 31x19x14` label, so length 31, width 19, height 14, one "
            "row per box.\n\n"
            "The totals are reconciled against the *Carton Total*, *Total # of Cartons* "
            "and *Total Quantity* figures printed in the file, so you can see at a "
            "glance that nothing was dropped."
        )
    st.stop()

try:
    parsed = _parse(uploaded.getvalue())
except PackingListError as exc:
    st.error(str(exc))
    st.stop()

output = build_output(parsed, combine_duplicates=combine_duplicates)
dimensions = build_dimensions(parsed, include_box_number=dimensions_box_number)
upc_widths = upc_display_widths(parsed.lines["UPCs"])

col1, col2, col3, col4 = st.columns(4)
col1.metric("Order No.", ", ".join(parsed.order_numbers) or "—")
col2.metric("Boxes", int(output["Box Number"].nunique()))
col3.metric("Rows", len(output))
col4.metric("Total quantity", int(output["Quantity"].sum()))

checks = _validation(parsed, output)
if any(level == "error" for level, _ in checks):
    st.error("Some totals do not reconcile — please review before using the file.")
for level, message in checks:
    if level == "ok":
        st.success(message, icon="✅")
    elif level == "warning":
        st.warning(message, icon="⚠️")
    else:
        st.error(message, icon="🚨")

if parsed.warnings:
    with st.expander(f"Parser notes ({len(parsed.warnings)})"):
        for note in parsed.warnings:
            st.write("- " + note)

st.subheader("Formatted result")
contents_tab, dimensions_tab = st.tabs(["Box Contents", "Box Dimensions"])
with contents_tab:
    st.dataframe(
        output,
        width="stretch",
        hide_index=True,
        column_config={
            "UPCs": (
                st.column_config.NumberColumn("UPCs", format="%d")
                if pd.api.types.is_numeric_dtype(output["UPCs"])
                else st.column_config.TextColumn("UPCs")
            ),
            "Box Number": st.column_config.NumberColumn("Box Number", format="%d"),
            "Quantity": st.column_config.NumberColumn("Quantity", format="%d"),
        },
    )
with dimensions_tab:
    st.dataframe(dimensions, width="stretch", hide_index=True)
    st.caption(
        "Read from each carton's 'Dimensions' label — the first number is the "
        "length, the second the width, the third the height."
    )

stem = Path(uploaded.name).stem
detail = parsed.lines if include_details else None

excel, contents_csv, dimensions_csv = st.columns(3)
with excel:
    st.download_button(
        "⬇️ Download Excel (both tabs)",
        data=to_excel_bytes(output, dimensions, detail, upc_widths=upc_widths),
        file_name=f"{stem} - Box Contents.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        width="stretch",
    )
with contents_csv:
    st.download_button(
        "⬇️ Contents CSV",
        data=output.to_csv(index=False).encode("utf-8-sig"),
        file_name=f"{stem} - Box Contents.csv",
        mime="text/csv",
        width="stretch",
    )
with dimensions_csv:
    st.download_button(
        "⬇️ Dimensions CSV",
        data=dimensions.to_csv(index=False).encode("utf-8-sig"),
        file_name=f"{stem} - Box Dimensions.csv",
        mime="text/csv",
        width="stretch",
    )

with st.expander("Per-box summary"):
    summary = pd.DataFrame(
        [
            {
                "Box Number": carton.box_number,
                "Carton No": carton.carton_no,
                "Dimensions": carton.dimensions,
                "UPCs in box": len({line["UPCs"] for line in carton.lines}),
                "Quantity": carton.parsed_total,
                "Carton Total in file": carton.reported_total,
            }
            for carton in parsed.cartons
        ]
    )
    st.dataframe(summary, width="stretch", hide_index=True)

with st.expander("Line detail (size, description, carton number)"):
    st.dataframe(parsed.lines, width="stretch", hide_index=True)
