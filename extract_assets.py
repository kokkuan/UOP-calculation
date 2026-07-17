"""
extract_assets.py

Appends one "<label>_ASSETS" sheet per airport to the existing
Output\\Revised Rate.xlsx workbook (must be built first by revised_rate.py).

Each asset sheet has:
  - columns "Asset" .. "ORI. USEFUL LIFE" — copied as-is, no calculation
  - "New DepKy" — VLOOKUP of NEW RATE 2026 (rounded to 4dp) against the
    "<label>_RATE" lookup sheet, giving the SAP depreciation key for the
    earliest year only
  - columns "NEW RATE" .. "Difference" — REAL Excel formulas (INDEX/MATCH,
    SUM), referencing the "<label>" rate-matrix sheet that already lives
    in the same workbook.

Run revised_rate.py first — this script only appends to its output.
"""

import re
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from revised_rate import (
    OUTPUT_DIR, FILES, YEARS, ALL_LIVES,
    sheet_name_for, lookup_sheet_name_for,
)

# ── UOP sheet: 0-based column indices ────────────────────────────────────────
UOP_HEADER_ROW     = 4   # row holding column titles ("Asset", "SNo.", ...)
UOP_DATA_START_ROW = 6   # first asset data row

COL_FIRST = 1    # B  Asset
COL_LAST  = 13   # N  ORI. USEFUL LIFE
# (columns in between: SNo., Cap.date, Asset description, Acquis.val.,
#  Accum.dep., Book val., Crcy, BusA, APC, Class, DepKy)

COL_ACQUIS_IDX = 4    # index of "Acquis.val." within the extracted row tuple
COL_LIFE_IDX   = 12   # index of "ORI. USEFUL LIFE" within the extracted row tuple

ASSET_INFO_COLS = 13   # number of raw asset-info columns written first

# ── Output sheet column layout (1-based) ─────────────────────────────────────
# Block order: asset info | New DepKy | NEW RATE (all years) | Sum/Diff | Dep Amount (all years) | Total/Diff
N_YEARS         = len(YEARS)
NEWDEPKY_COL    = ASSET_INFO_COLS + 1
NEW_RATE_START  = NEWDEPKY_COL + 1
SUM_COL         = NEW_RATE_START + N_YEARS
DIFF_COL        = SUM_COL + 1
DEP_START       = DIFF_COL + 1
TOTAL_COL       = DEP_START + N_YEARS
DIFFTOTAL_COL   = TOTAL_COL + 1

LIFE_COL_LETTER   = get_column_letter(COL_LIFE_IDX + 1)     # M
ACQUIS_COL_LETTER = get_column_letter(COL_ACQUIS_IDX + 1)   # E

# ── Rate-matrix sheet layout (built by revised_rate.py's add_rate_sheet) ────
RATE_MATRIX_LAST_COL   = get_column_letter(1 + N_YEARS)        # AS
RATE_MATRIX_DATA_ROW1  = 4
RATE_MATRIX_DATA_ROWN  = 3 + len(ALL_LIVES)                    # 46

NUMERIC_2DP_HEADERS = {"Acquis.val.", "Accum.dep.", "Book val.", "Total Dep", "Diff (Total-Acquis)"}
PCT_HEADER_PREFIXES = ("NEW RATE", "Sum UOP Rates", "Diff (Sum-1)")

DIVIDER_SHEET_NAME = "> Asset List"
DIVIDER_FILL       = PatternFill("solid", fgColor="ED7D31")   # orange, stands out from the blue rate sheets

SUMMARY_SHEET_NAME = "Summary"
TOTAL_FILL         = PatternFill("solid", fgColor="D9E1F2")   # light blue for the grand-total row

# ── Styling (matches revised_rate.py) ────────────────────────────────────────
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
SUBHDR_FILL = PatternFill("solid", fgColor="2E75B6")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=10)
LABEL_FONT  = Font(bold=True, size=10)
THIN        = Side(style="thin", color="BFBFBF")
BORDER      = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _cell(ws, row, col, value=None, fill=None, font=None, align="center", num_fmt=None):
    c = ws.cell(row=row, column=col, value=value)
    if fill:    c.fill      = fill
    if font:    c.font      = font
    if num_fmt: c.number_format = num_fmt
    c.alignment = Alignment(horizontal=align, vertical="center")
    c.border    = BORDER
    return c


def read_asset_listing(ws_uop) -> tuple:
    """
    Read asset line items from the UOP sheet, columns Asset .. ORI. USEFUL LIFE.
    Returns (headers, rows) where each row is a tuple of raw cell values.
    """
    headers = [
        str(ws_uop.cell(row=UOP_HEADER_ROW, column=c + 1).value).strip()
        for c in range(COL_FIRST, COL_LAST + 1)
    ]

    rows = []
    for row in ws_uop.iter_rows(min_row=UOP_DATA_START_ROW, values_only=True):
        asset_no = row[COL_FIRST]
        if asset_no is None:
            continue
        rows.append(row[COL_FIRST:COL_LAST + 1])

    return headers, rows


def build_full_headers(asset_headers: list) -> list:
    headers = list(asset_headers)
    headers += ["New DepKy"]
    headers += [f"NEW RATE {y}" for y in YEARS]
    headers += ["Sum UOP Rates", "Diff (Sum-1)"]
    headers += [f"Dep Amount {y}" for y in YEARS]
    headers += ["Total Dep", "Diff (Total-Acquis)"]
    return headers


def write_asset_row(ws, row_idx: int, asset_row: tuple, rate_sheet: str, lookup_sheet: str):
    # ── Asset info columns (raw values, as-is) ───────────────────────────────
    for col_idx, val in enumerate(asset_row, start=1):
        align = "left" if isinstance(val, str) else "center"
        _cell(ws, row_idx, col_idx, val, align=align)

    life_ref   = f"${LIFE_COL_LETTER}{row_idx}"
    acquis_ref = f"${ACQUIS_COL_LETTER}{row_idx}"

    # ── New DepKy: VLOOKUP of NEW RATE 2026 (rounded) against the RATE lookup sheet ──
    new_rate_2026_ref = f"{get_column_letter(NEW_RATE_START)}{row_idx}"
    depkey_formula = (
        f"=IFERROR(VLOOKUP(ROUND({new_rate_2026_ref},4),"
        f"'{lookup_sheet}'!$A:$B,2,FALSE),\"\")"
    )
    _cell(ws, row_idx, NEWDEPKY_COL, depkey_formula)

    # ── NEW RATE per year: INDEX/MATCH against the rate-matrix sheet ────────
    new_rate_refs = []
    for i, y in enumerate(YEARS):
        col = NEW_RATE_START + i
        formula = (
            f"=IFERROR(INDEX('{rate_sheet}'!$B${RATE_MATRIX_DATA_ROW1}:"
            f"${RATE_MATRIX_LAST_COL}${RATE_MATRIX_DATA_ROWN},"
            f"MATCH({life_ref},'{rate_sheet}'!$A${RATE_MATRIX_DATA_ROW1}:"
            f"$A${RATE_MATRIX_DATA_ROWN},0),"
            f"MATCH({y},'{rate_sheet}'!$B$3:${RATE_MATRIX_LAST_COL}$3,0)),0)"
        )
        _cell(ws, row_idx, col, formula, num_fmt="0.0000%")
        new_rate_refs.append(f"{get_column_letter(col)}{row_idx}")

    # ── Sum of rates / difference from 1.0 ───────────────────────────────────
    sum_range = f"{get_column_letter(NEW_RATE_START)}{row_idx}:{get_column_letter(NEW_RATE_START + N_YEARS - 1)}{row_idx}"
    _cell(ws, row_idx, SUM_COL, f"=SUM({sum_range})", num_fmt="0.0000%")
    _cell(ws, row_idx, DIFF_COL, f"={get_column_letter(SUM_COL)}{row_idx}-1", num_fmt="0.0000%")

    # ── Dep amount per year = acquisition value * rate ───────────────────────
    for i in range(N_YEARS):
        col = DEP_START + i
        formula = f"={acquis_ref}*{new_rate_refs[i]}"
        _cell(ws, row_idx, col, formula, num_fmt="#,##0.00")

    # ── Total dep / difference from acquisition value ────────────────────────
    dep_range = f"{get_column_letter(DEP_START)}{row_idx}:{get_column_letter(DEP_START + N_YEARS - 1)}{row_idx}"
    _cell(ws, row_idx, TOTAL_COL, f"=SUM({dep_range})", num_fmt="#,##0.00")
    _cell(ws, row_idx, DIFFTOTAL_COL, f"={get_column_letter(TOTAL_COL)}{row_idx}-{acquis_ref}", num_fmt="#,##0.00")


def add_divider_sheet(wb_out, sheet_name: str):
    """A visual separator tab marking where the rate/lookup sheets end and the asset listings begin."""
    if sheet_name in wb_out.sheetnames:
        wb_out.remove(wb_out[sheet_name])
    ws = wb_out.create_sheet(title=sheet_name)

    _cell(ws, 2, 2, "Asset Listing", fill=DIVIDER_FILL, font=Font(bold=True, color="FFFFFF", size=18), align="left")
    ws.merge_cells(start_row=2, start_column=2, end_row=2, end_column=6)
    _cell(ws, 4, 2, "Sheets to the right list every asset line item with recalculated UOP rates.",
          font=Font(italic=True, size=10), align="left")
    ws.merge_cells(start_row=4, start_column=2, end_row=4, end_column=6)

    for cell_row in (2, 4):
        for c in range(2, 7):
            ws.cell(row=cell_row, column=c).border = Border()

    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = "ED7D31"
    ws.column_dimensions[get_column_letter(1)].width = 3
    for col_idx in range(2, 7):
        ws.column_dimensions[get_column_letter(col_idx)].width = 22
    ws.row_dimensions[2].height = 26


def add_asset_sheet(wb_out, sheet_name: str, airport_label: str, headers: list, rows: list, rate_sheet: str, lookup_sheet: str):
    if sheet_name in wb_out.sheetnames:
        wb_out.remove(wb_out[sheet_name])
    ws = wb_out.create_sheet(title=sheet_name)

    total_cols = len(headers)
    _cell(ws, 1, 1, airport_label, fill=HEADER_FILL, font=Font(bold=True, color="FFFFFF", size=11), align="left")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)

    for col_idx, h in enumerate(headers, start=1):
        _cell(ws, 2, col_idx, h, fill=SUBHDR_FILL, font=HEADER_FONT)

    for row_idx, asset_row in enumerate(rows, start=3):
        write_asset_row(ws, row_idx, asset_row, rate_sheet, lookup_sheet)

    for col_idx, h in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = max(11, len(h) + 2)

    ws.freeze_panes = "C3"


def add_summary_sheet(wb_out, sheet_name: str, airports: list):
    """
    Summary tab: rows = airport, columns = years 2026-2069, values = total
    depreciation for that airport/year, plus a per-airport row total and a
    grand-total row across all airports. Each cell is a real formula
    (full-column SUM against that airport's "<label>_ASSETS" sheet).
    airports: list of (label, assets_sheet_name)
    """
    if sheet_name in wb_out.sheetnames:
        wb_out.remove(wb_out[sheet_name])
    ws = wb_out.create_sheet(title=sheet_name, index=0)

    total_col   = 2 + N_YEARS   # one "Airport" col + N_YEARS year cols, then Total
    total_cols  = total_col

    _cell(ws, 1, 1, "Depreciation Summary by Airport (2026 - 2069)",
          fill=HEADER_FILL, font=Font(bold=True, color="FFFFFF", size=13), align="left")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)

    _cell(ws, 2, 1, "Airport", fill=SUBHDR_FILL, font=HEADER_FONT)
    for col_idx, y in enumerate(YEARS, start=2):
        _cell(ws, 2, col_idx, y, fill=SUBHDR_FILL, font=HEADER_FONT)
    _cell(ws, 2, total_col, "Total", fill=SUBHDR_FILL, font=HEADER_FONT)

    first_data_row = 3
    for row_idx, (label, assets_sheet_name) in enumerate(airports, start=first_data_row):
        _cell(ws, row_idx, 1, label, font=LABEL_FONT, align="left")
        for i, y in enumerate(YEARS):
            col = 2 + i
            dep_col_letter = get_column_letter(DEP_START + i)
            formula = f"=SUM('{assets_sheet_name}'!{dep_col_letter}:{dep_col_letter})"
            _cell(ws, row_idx, col, formula, num_fmt="#,##0.00")
        row_range = f"{get_column_letter(2)}{row_idx}:{get_column_letter(1 + N_YEARS)}{row_idx}"
        _cell(ws, row_idx, total_col, f"=SUM({row_range})", num_fmt="#,##0.00", font=LABEL_FONT)

    total_row = first_data_row + len(airports)
    _cell(ws, total_row, 1, "Total", fill=TOTAL_FILL, font=LABEL_FONT, align="left")
    for i in range(N_YEARS):
        col = 2 + i
        col_range = f"{get_column_letter(col)}{first_data_row}:{get_column_letter(col)}{total_row - 1}"
        _cell(ws, total_row, col, f"=SUM({col_range})", fill=TOTAL_FILL, num_fmt="#,##0.00", font=LABEL_FONT)
    grand_range = f"{get_column_letter(2)}{total_row}:{get_column_letter(1 + N_YEARS)}{total_row}"
    _cell(ws, total_row, total_col, f"=SUM({grand_range})", fill=TOTAL_FILL, num_fmt="#,##0.00", font=LABEL_FONT)

    ws.column_dimensions[get_column_letter(1)].width = 14
    for col_idx in range(2, total_cols + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 13
    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "B3"


# ── File processor ────────────────────────────────────────────────────────────

def label_for_file(filepath: Path) -> str:
    label_tokens = [t for t in filepath.stem.split(" ") if not re.match(r"^\d+\.$", t)]
    return label_tokens[0]


def process_file(filepath: Path) -> tuple:
    """Process one airport file. Returns (label, headers, rows)."""
    wb_ro = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    headers, rows = read_asset_listing(wb_ro["UOP"])
    wb_ro.close()

    label = label_for_file(filepath)
    print(f"{label:10s}: {len(rows)} asset line items")
    return (label, headers, rows)


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    summary_path = OUTPUT_DIR / "Revised Rate.xlsx"
    if not summary_path.exists():
        raise SystemExit(f"{summary_path} not found — run revised_rate.py first.")

    wb = openpyxl.load_workbook(summary_path)

    # Clean up any existing summary/divider/asset sheets first, so re-running
    # always rebuilds them at the end, in order, right after a fresh divider.
    if SUMMARY_SHEET_NAME in wb.sheetnames:
        wb.remove(wb[SUMMARY_SHEET_NAME])
    if DIVIDER_SHEET_NAME in wb.sheetnames:
        wb.remove(wb[DIVIDER_SHEET_NAME])
    for f in FILES:
        assets_sheet_name = (label_for_file(f).replace(" ", "_") + "_ASSETS")[:31]
        if assets_sheet_name in wb.sheetnames:
            wb.remove(wb[assets_sheet_name])

    add_divider_sheet(wb, DIVIDER_SHEET_NAME)

    built_airports = []
    for f in FILES:
        if not f.exists():
            print(f"WARNING: File not found — {f}")
            continue

        label, asset_headers, rows = process_file(f)

        rate_sheet   = sheet_name_for(label)
        lookup_sheet = lookup_sheet_name_for(label)
        if rate_sheet not in wb.sheetnames or lookup_sheet not in wb.sheetnames:
            print(f"WARNING: {label} — '{rate_sheet}'/'{lookup_sheet}' sheets not found in {summary_path.name}, skipping")
            continue

        full_headers = build_full_headers(asset_headers)
        assets_sheet_name = (label.replace(" ", "_") + "_ASSETS")[:31]
        add_asset_sheet(wb, assets_sheet_name, label, full_headers, rows, rate_sheet, lookup_sheet)
        built_airports.append((label, assets_sheet_name))

    add_summary_sheet(wb, SUMMARY_SHEET_NAME, built_airports)

    wb.save(summary_path)
    print(f"\nAsset listing (with formulas) added to: {summary_path.name}")
