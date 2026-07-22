"""
app.py

Streamlit front-end for the UOP depreciation recalculation pipeline.

Upload the consolidated workbook (must contain "1. PAX BUDGET", "1.MPPA",
"MASTERFILE", "2.DATABASE", and "3.DEPKEY FROM SAP" sheets) and download a
single "PAX_Rate_Matrix_Test.xlsx" containing:
  - one rate-matrix sheet per airport (useful life x year, 2026-2069)
  - a DEPKEY_LOOKUP sheet (Rate -> DepKy, copied as-is from 3.DEPKEY FROM SAP)
  - one "<Airport>_ASSETS" sheet per airport with real Excel formulas
    (INDEX/MATCH, VLOOKUP, SUM) joining 2.DATABASE assets to their airport's
    rate matrix via MASTERFILE (BusA -> Airport)
  - a Summary sheet (depreciation by airport/year) and a TO NOTE sheet
    (assets with no matching New DepKy)

Nothing is written to disk server-side — everything happens in memory for
the duration of the request, and the result is streamed back as a download.
"""

import io

import streamlit as st

from pax_rate_matrix import build_rate_matrix_workbook
from extract_assets import add_asset_sheets

REQUIRED_SHEETS = ("1. PAX BUDGET", "1.MPPA", "MASTERFILE", "2.DATABASE", "3.DEPKEY FROM SAP")


def run_pipeline(src_bytes: bytes) -> tuple:
    """
    src_bytes: the uploaded consolidated workbook's raw bytes.
    Returns (output_bytesio, log_lines).
    """
    wb, airport_key_map, log = build_rate_matrix_workbook(src_bytes)
    log.append(f"\n{len(airport_key_map)} airports matched to a rate matrix.")

    log += add_asset_sheets(wb, src_bytes)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf, log


# ── Streamlit UI ──────────────────────────────────────────────────────────────

st.set_page_config(page_title="UOP Depreciation Recalculation", page_icon="✈️", layout="wide")

st.title("UOP Depreciation Recalculation")
st.caption(
    "Upload the consolidated workbook (must contain \"1. PAX BUDGET\", \"1.MPPA\", \"MASTERFILE\", "
    "\"2.DATABASE\", and \"3.DEPKEY FROM SAP\" sheets). "
    "Produces a combined PAX_Rate_Matrix_Test.xlsx with recalculated rates, DepKy lookups, "
    "and a cross-airport summary — all as live Excel formulas."
)

uploaded_file = st.file_uploader("Consolidated UOP workbook", type=["xlsx"])

if uploaded_file:
    st.write(f"Ready: {uploaded_file.name}")

    if st.button("Run recalculation", type="primary"):
        with st.spinner("Processing..."):
            src_bytes = uploaded_file.getvalue()
            wb_ro = None
            try:
                import openpyxl
                wb_ro = openpyxl.load_workbook(io.BytesIO(src_bytes), read_only=True)
                missing = [s for s in REQUIRED_SHEETS if s not in wb_ro.sheetnames]
            finally:
                if wb_ro is not None:
                    wb_ro.close()

            if missing:
                st.error(f"Missing required sheet(s): {missing}")
            else:
                output_buf, log_lines = run_pipeline(src_bytes)

                st.success("Done.")
                st.code("\n".join(log_lines))

                st.download_button(
                    "Download PAX_Rate_Matrix_Test.xlsx",
                    data=output_buf,
                    file_name="PAX_Rate_Matrix_Test.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
else:
    st.info("Upload the consolidated UOP workbook to get started.")
