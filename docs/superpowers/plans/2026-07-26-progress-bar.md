# Progress Bar for UOP Depreciation Pipeline Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Show a live progress bar and current-airport status text in the Streamlit UI while the pipeline builds rate-matrix and asset sheets, instead of a static spinner.

**Architecture:** Thread an optional `progress_cb(phase, i, total, label)` callback through the two per-airport loops in `pax_rate_matrix.build_rate_matrix_workbook()` and `extract_assets.add_asset_sheets()`. `app.py` supplies a closure that updates a single `st.progress` bar and `st.empty()` status text, computed as one combined fraction across both phases.

**Tech Stack:** Python, Streamlit, openpyxl. No test framework exists in this repo (no `tests/` dir, no pytest) — verification is manual, via a throwaway script run against the real source workbook already checked into the repo root (`WORKING FOR UOP CALCULATION MA S MASB 21072026.xlsx`).

## Global Constraints

- `progress_cb` must default to `None` everywhere it's added, and be a no-op when `None` — the CLI entry points (`pax_rate_matrix.main()`, `extract_assets.__main__` block) call these functions without a callback and must keep working unchanged.
- No change to existing log content, return values, sheet contents, or CLI printed output.
- One combined progress bar in the UI (not two separate bars per phase) — confirmed by user.
- Callback is called once per airport iterated in each phase's loop (28ish times per phase), not more granularly.

---

### Task 1: Add `progress_cb` to `build_rate_matrix_workbook`

**Files:**
- Modify: `pax_rate_matrix.py:209-238`

**Interfaces:**
- Produces: `build_rate_matrix_workbook(src=SRC, progress_cb=None)` — unchanged return value `(wb_out, airport_key_map, log)`. If provided, `progress_cb` is called as `progress_cb("rate_matrix", i, total, airport)` once per airport, 1-based `i`, `total = len(matched_airports)`, after that airport's sheet is added.

- [ ] **Step 1: Edit the function signature and loop**

Change:
```python
def build_rate_matrix_workbook(src=SRC):
```
to:
```python
def build_rate_matrix_workbook(src=SRC, progress_cb=None):
```

Change the loop body (currently):
```python
    log = []
    for airport in matched_airports:
        rates = build_rate_matrix(capped[airport], set(ALL_LIVES))
        sheet_name = sheet_name_for(airport)
        add_rate_sheet(wb_out, sheet_name, airport, rates, ALL_LIVES, capped[airport])

        was_capped = pax[airport][START_YEAR] != capped[airport][START_YEAR]
        if was_capped:
            wb_out[sheet_name].cell(row=PAX_ROW, column=YEAR_COL_START_CELL).fill = CAPPED_FILL
        log.append(f"Sheet added: {airport}{'  (2026 capped — highlighted)' if was_capped else ''}")

    return wb_out, airport_key_map, log
```
to:
```python
    log = []
    total = len(matched_airports)
    for i, airport in enumerate(matched_airports, start=1):
        rates = build_rate_matrix(capped[airport], set(ALL_LIVES))
        sheet_name = sheet_name_for(airport)
        add_rate_sheet(wb_out, sheet_name, airport, rates, ALL_LIVES, capped[airport])

        was_capped = pax[airport][START_YEAR] != capped[airport][START_YEAR]
        if was_capped:
            wb_out[sheet_name].cell(row=PAX_ROW, column=YEAR_COL_START_CELL).fill = CAPPED_FILL
        log.append(f"Sheet added: {airport}{'  (2026 capped — highlighted)' if was_capped else ''}")

        if progress_cb is not None:
            progress_cb("rate_matrix", i, total, airport)

    return wb_out, airport_key_map, log
```

- [ ] **Step 2: Verify the CLI entry point still works unchanged**

Run:
```bash
python pax_rate_matrix.py
```
Expected: same console output as before (list of "Sheet added: ..." lines, then "Rate matrix workbook saved: ..."), no errors. This confirms `progress_cb=None` default doesn't break the existing no-callback call path.

- [ ] **Step 3: Verify the callback fires correctly**

Run:
```bash
python -c "
from pax_rate_matrix import build_rate_matrix_workbook
calls = []
wb, key_map, log = build_rate_matrix_workbook(progress_cb=lambda phase, i, total, label: calls.append((phase, i, total, label)))
print('total calls:', len(calls))
print('first:', calls[0])
print('last:', calls[-1])
assert all(c[0] == 'rate_matrix' for c in calls)
assert calls[-1][1] == calls[-1][2] == len(calls)
print('OK')
"
```
Expected: `OK` printed, with `total calls` matching the number of "Sheet added" log lines, and the last call's `i == total == total calls`.

- [ ] **Step 4: Commit**

```bash
git add pax_rate_matrix.py
git commit -m "Add optional progress callback to build_rate_matrix_workbook"
```

---

### Task 2: Add `progress_cb` to `add_asset_sheets`

**Files:**
- Modify: `extract_assets.py:459-502`

**Interfaces:**
- Consumes: nothing new from Task 1 (independent function).
- Produces: `add_asset_sheets(wb, src=SRC, progress_cb=None)` — unchanged return value `log`. If provided, `progress_cb` is called as `progress_cb("assets", i, total, label)` once per airport in `sorted(grouped)`, 1-based `i`, `total = len(grouped)`, after that airport's iteration completes (whether its asset sheet was built or skipped with a warning).

- [ ] **Step 1: Edit the function signature and loop**

Change:
```python
def add_asset_sheets(wb, src=SRC):
```
to:
```python
def add_asset_sheets(wb, src=SRC, progress_cb=None):
```

Change the loop body (currently):
```python
    full_headers = build_full_headers()
    built_airports = []
    log = []
    for label in sorted(grouped):
        rows = grouped[label]
        rate_sheet = sheet_name_for(label)
        if rate_sheet not in wb.sheetnames:
            log.append(f"WARNING: {label} — rate-matrix sheet '{rate_sheet}' not found, skipping ({len(rows)} assets)")
            continue

        log.append(f"{label:20s}: {len(rows)} asset line items")
        assets_sheet_name = (label.replace(" ", "_") + "_ASSETS")[:31]
        add_asset_sheet(wb, assets_sheet_name, label, full_headers, rows, rate_sheet)
        built_airports.append((label, assets_sheet_name))

    add_summary_sheet(wb, SUMMARY_SHEET_NAME, built_airports)
```
to:
```python
    full_headers = build_full_headers()
    built_airports = []
    log = []
    airport_labels = sorted(grouped)
    total = len(airport_labels)
    for i, label in enumerate(airport_labels, start=1):
        rows = grouped[label]
        rate_sheet = sheet_name_for(label)
        if rate_sheet not in wb.sheetnames:
            log.append(f"WARNING: {label} — rate-matrix sheet '{rate_sheet}' not found, skipping ({len(rows)} assets)")
            if progress_cb is not None:
                progress_cb("assets", i, total, label)
            continue

        log.append(f"{label:20s}: {len(rows)} asset line items")
        assets_sheet_name = (label.replace(" ", "_") + "_ASSETS")[:31]
        add_asset_sheet(wb, assets_sheet_name, label, full_headers, rows, rate_sheet)
        built_airports.append((label, assets_sheet_name))

        if progress_cb is not None:
            progress_cb("assets", i, total, label)

    add_summary_sheet(wb, SUMMARY_SHEET_NAME, built_airports)
```

- [ ] **Step 2: Verify the CLI entry point still works unchanged**

Run:
```bash
python pax_rate_matrix.py
python extract_assets.py
```
Expected: `extract_assets.py` prints the same per-airport lines and "TO NOTE: N asset(s)..." as before, with `PAX_Rate_Matrix_Test.xlsx` saved successfully. Confirms the `progress_cb=None` path is unaffected. (First line runs `pax_rate_matrix.py` because `extract_assets.py` requires `Output/PAX_Rate_Matrix_Test.xlsx` to already exist.)

- [ ] **Step 3: Verify the callback fires correctly**

Run:
```bash
python -c "
import openpyxl
from pax_rate_matrix import build_rate_matrix_workbook
from extract_assets import add_asset_sheets

wb, _, _ = build_rate_matrix_workbook()
calls = []
log = add_asset_sheets(wb, progress_cb=lambda phase, i, total, label: calls.append((phase, i, total, label)))
print('total calls:', len(calls))
print('first:', calls[0])
print('last:', calls[-1])
assert all(c[0] == 'assets' for c in calls)
assert calls[-1][1] == calls[-1][2] == len(calls)
print('OK')
"
```
Expected: `OK` printed, with `total calls` matching `len(sorted(grouped))` (i.e. the number of distinct airports found in `2.DATABASE`), and the last call's `i == total == total calls`.

- [ ] **Step 4: Commit**

```bash
git add extract_assets.py
git commit -m "Add optional progress callback to add_asset_sheets"
```

---

### Task 3: Wire the progress bar into `app.py`

**Files:**
- Modify: `app.py:31-89`

**Interfaces:**
- Consumes: `build_rate_matrix_workbook(src, progress_cb=None)` from Task 1, `add_asset_sheets(wb, src, progress_cb=None)` from Task 2 — both called with `phase` one of `"rate_matrix"` / `"assets"`, 1-based `i`, `total`, and `label` (airport name).
- Produces: `run_pipeline(src_bytes, progress_cb=None)` — same return value `(output_buf, log_lines)` as before.

- [ ] **Step 1: Update `run_pipeline` to forward the callback**

Change:
```python
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
```
to:
```python
def run_pipeline(src_bytes: bytes, progress_cb=None) -> tuple:
    """
    src_bytes: the uploaded consolidated workbook's raw bytes.
    progress_cb: optional callback(phase, i, total, label), forwarded to
    both pipeline phases — see build_rate_matrix_workbook and
    add_asset_sheets for the exact contract.
    Returns (output_bytesio, log_lines).
    """
    wb, airport_key_map, log = build_rate_matrix_workbook(src_bytes, progress_cb=progress_cb)
    log.append(f"\n{len(airport_key_map)} airports matched to a rate matrix.")

    log += add_asset_sheets(wb, src_bytes, progress_cb=progress_cb)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf, log
```

- [ ] **Step 2: Replace the spinner with a progress bar + status text**

Change:
```python
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
```
to:
```python
    if st.button("Run recalculation", type="primary"):
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
            progress_bar = st.progress(0)
            status_text = st.empty()
            progress_state = {"rate_matrix_total": None, "assets_total": None}

            def progress_cb(phase, i, total, label):
                if phase == "rate_matrix":
                    progress_state["rate_matrix_total"] = total
                    assets_total_estimate = progress_state["assets_total"] or total
                    rate_matrix_total = total
                    done = i
                    phase_label = "Building rate matrix"
                else:
                    progress_state["assets_total"] = total
                    rate_matrix_total = progress_state["rate_matrix_total"] or total
                    assets_total_estimate = total
                    done = rate_matrix_total + i
                    phase_label = "Extracting assets"

                grand_total = rate_matrix_total + assets_total_estimate
                progress_bar.progress(min(done / grand_total, 1.0))
                status_text.text(f"{phase_label}: {label} ({i}/{total})")

            output_buf, log_lines = run_pipeline(src_bytes, progress_cb=progress_cb)

            progress_bar.progress(1.0)
            status_text.empty()
            st.success("Done.")
            st.code("\n".join(log_lines))

            st.download_button(
                "Download PAX_Rate_Matrix_Test.xlsx",
                data=output_buf,
                file_name="PAX_Rate_Matrix_Test.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
```

Note on the combined fraction: while still in the `"rate_matrix"` phase, `assets_total` isn't known yet, so it's estimated as equal to `rate_matrix_total` (both phases iterate over the same universe of MASTERFILE airports, so counts are normally equal or very close). Once the `"assets"` phase starts, the real `assets_total` replaces the estimate and the denominator becomes exact for the rest of the run. This means the bar may jump slightly at the phase boundary if the two counts differ — acceptable for a progress indicator, not something requiring precision.

- [ ] **Step 3: Manually verify in the running app**

Run:
```bash
streamlit run app.py
```
Then in the browser: upload `WORKING FOR UOP CALCULATION MA S MASB 21072026.xlsx` (already in the repo root) and click "Run recalculation".

Expected:
- Progress bar starts at 0 and advances smoothly to 100% without jumping backward.
- Status text shows `"Building rate matrix: <airport> (i/total)"` for each of the ~28 airports, then switches to `"Extracting assets: <airport> (i/total)"` for the second phase.
- After completion, the status text clears, "Done." success message and the existing log/download button appear exactly as before.
- Stop the app (Ctrl+C) once confirmed.

- [ ] **Step 4: Commit**

```bash
git add app.py
git commit -m "Show a live progress bar and per-airport status during pipeline runs"
```

---

## Self-Review Notes

- **Spec coverage:** Task 1 covers spec item 1 (`build_rate_matrix_workbook` callback), Task 2 covers item 2 (`add_asset_sheets` callback), Task 3 covers items 3-4 (`run_pipeline` forwarding + Streamlit UI). "Out of scope" items (no `find_new_depky_gaps` progress, single combined bar, no log/CLI changes) are respected — CLI verification steps in Tasks 1-2 explicitly confirm no behavior change.
- **Ambiguity resolved:** the spec's combined-fraction formula didn't address that `total_assets` is unknown during the `rate_matrix` phase. Task 3 resolves this explicitly: estimate `assets_total` as equal to `rate_matrix_total` until the real value is known, then switch to the exact value — documented inline as an acceptable approximation.
- **Type/signature consistency:** `progress_cb(phase, i, total, label)` signature is identical across Task 1, Task 2, and Task 3's consumption of it.
