"""
app.py

Streamlit front-end for the UOP depreciation recalculation pipeline.

Upload each airport's Excel file (must contain REVISE RATE, UOP, and RATE
sheets) and download a combined "Revised Rate.xlsx" containing, per airport:
  - a rate-matrix sheet (useful life x year, 2026-2069)
  - a RATE lookup sheet (Rate -> DepKy, copied as-is)
  - an asset-listing sheet with real Excel formulas (INDEX/MATCH, VLOOKUP, SUM)
plus a Summary sheet (depreciation by airport/year) and a divider tab.

Nothing is written to disk server-side — everything happens in memory for
the duration of the request, and the result is streamed back as a download.
"""

import io
import re

import openpyxl
import streamlit as st

from revised_rate import (
    YEARS, ALL_LIVES,
    read_pax, read_rate_lookup, build_rate_matrix,
    add_rate_sheet, add_rate_lookup_sheet,
    sheet_name_for, lookup_sheet_name_for,
)
from extract_assets import (
    read_asset_listing, build_full_headers, add_asset_sheet,
    add_divider_sheet, add_summary_sheet,
    DIVIDER_SHEET_NAME, SUMMARY_SHEET_NAME,
)

REQUIRED_SHEETS = ("REVISE RATE", "UOP", "RATE")


def label_for_filename(filename: str) -> str:
    stem = filename.rsplit(".", 1)[0]
    tokens = [t for t in stem.split(" ") if not re.match(r"^\d+\.$", t)]
    return tokens[0] if tokens else stem


def run_pipeline(files: list) -> tuple:
    """
    files: list of (filename, file-like) tuples, each file-like seekable and
    openable by openpyxl (e.g. a Streamlit UploadedFile).
    Returns (output_bytesio, log_lines).
    """
    log = []
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    rate_pass = []   # (label, filename, filelike) carried into the asset-listing pass
    for filename, filelike in files:
        label = label_for_filename(filename)
        filelike.seek(0)
        wb_ro = openpyxl.load_workbook(filelike, read_only=True, data_only=True)

        missing = [s for s in REQUIRED_SHEETS if s not in wb_ro.sheetnames]
        if missing:
            log.append(f"SKIPPED {filename}: missing sheet(s) {missing}")
            wb_ro.close()
            continue

        pax = read_pax(wb_ro["REVISE RATE"])
        rate_lookup = read_rate_lookup(wb_ro["RATE"])
        wb_ro.close()

        rates = build_rate_matrix(pax, set(ALL_LIVES))
        add_rate_sheet(wb, sheet_name_for(label), label, rates, ALL_LIVES, pax)
        add_rate_lookup_sheet(wb, lookup_sheet_name_for(label), airport_label=label, rate_lookup=rate_lookup)
        log.append(f"{label}: rate matrix built, {len(rate_lookup)} RATE lookup rows copied")
        rate_pass.append((label, filename, filelike))

    add_divider_sheet(wb, DIVIDER_SHEET_NAME)

    built_airports = []
    for label, filename, filelike in rate_pass:
        filelike.seek(0)
        wb_ro = openpyxl.load_workbook(filelike, read_only=True, data_only=True)
        asset_headers, rows = read_asset_listing(wb_ro["UOP"])
        wb_ro.close()

        full_headers = build_full_headers(asset_headers)
        assets_sheet_name = (label.replace(" ", "_") + "_ASSETS")[:31]
        add_asset_sheet(wb, assets_sheet_name, label, full_headers, rows,
                         sheet_name_for(label), lookup_sheet_name_for(label))
        built_airports.append((label, assets_sheet_name))
        log.append(f"{label}: {len(rows)} asset line items")

    add_summary_sheet(wb, SUMMARY_SHEET_NAME, built_airports)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf, log


# ── Streamlit UI ──────────────────────────────────────────────────────────────

st.set_page_config(page_title="UOP Depreciation Recalculation", page_icon="✈️", layout="wide")

st.title("UOP Depreciation Recalculation")
st.caption(
    "Upload each airport's Excel file (must contain REVISE RATE, UOP, and RATE sheets). "
    "Produces a combined Revised Rate.xlsx with recalculated rates, DepKy lookups, "
    "and a cross-airport summary — all as live Excel formulas."
)

uploaded_files = st.file_uploader(
    "Airport Excel files", type=["xlsx"], accept_multiple_files=True
)

if uploaded_files:
    st.write(f"{len(uploaded_files)} file(s) ready: " + ", ".join(f.name for f in uploaded_files))

    if st.button("Run recalculation", type="primary"):
        with st.spinner("Processing..."):
            files = [(f.name, f) for f in uploaded_files]
            output_buf, log_lines = run_pipeline(files)

        st.success("Done.")
        st.code("\n".join(log_lines))

        st.download_button(
            "Download Revised Rate.xlsx",
            data=output_buf,
            file_name="Revised Rate.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("Upload one or more airport Excel files to get started.")
