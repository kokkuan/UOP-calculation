# Progress bar for the UOP depreciation pipeline

## Purpose

The Streamlit UI (`app.py`) currently shows a static `st.spinner("Processing...")`
for the whole pipeline run (building ~28 per-airport rate-matrix sheets, then
~28 per-airport asset sheets). The user wants visible progress: an actual
progress bar plus which airport/sheet is currently being processed.

## Approach

Thread an optional progress callback through the two library entry points so
the UI can observe per-airport progress without changing their return values
or CLI behavior.

1. **`pax_rate_matrix.build_rate_matrix_workbook(src, progress_cb=None)`** —
   inside the per-airport loop, call
   `progress_cb("rate_matrix", i, total, airport_label)` after each sheet is
   added (1-based `i`, `total = len(matched_airports)`).

2. **`extract_assets.add_asset_sheets(wb, src, progress_cb=None)`** — same
   pattern inside its per-airport loop:
   `progress_cb("assets", i, total, airport_label)` after each `_ASSETS`
   sheet is added (`total = len(grouped)`).

   Both callbacks default to `None` and are only invoked if provided, so
   `main()` / `if __name__ == "__main__"` blocks in both files are
   unaffected.

3. **`app.py` `run_pipeline(src_bytes, progress_cb=None)`** — forwards
   `progress_cb` to both calls above.

4. **Streamlit UI** — replace `with st.spinner("Processing..."):` with:
   - `st.progress(0)` bar
   - `st.empty()` placeholder for status text
   - a closure passed as `progress_cb` that:
     - computes overall fraction complete across both phases combined
       (`(done_in_phase + phase_offset) / grand_total`, where
       `grand_total = total_rate_matrix + total_assets`, and
       `phase_offset = total_rate_matrix` once in the assets phase)
     - updates the bar via `.progress(fraction)`
     - updates the text via `.text(f"{phase_label}: {label} ({i}/{total})")`,
       e.g. `"Building rate matrix: Kuching (14/28)"` or
       `"Extracting assets: Kuching_ASSETS (5/28)"`

   The existing `st.code("\n".join(log_lines))` summary after completion is
   unchanged.

## Out of scope

- No change to the log content, sheet contents, or CLI scripts' printed
  output.
- No separate progress bar per phase (one combined bar, per user choice).
- No progress reporting inside `find_new_depky_gaps` (a fast pass with no
  per-airport sheet writes to report against).

## Testing

Manual: run the Streamlit app, upload the consolidated workbook, confirm the
bar advances smoothly from 0 to 100% and the status text names each airport
as it's processed, across both phases. `main()` CLI runs for both modules
must still work unchanged (no `progress_cb` passed).
