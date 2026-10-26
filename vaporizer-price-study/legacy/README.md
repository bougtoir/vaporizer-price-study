# legacy/

`analyze_old.py` is the pre-publication pipeline kept for history only. It is
quarantined here because its `SCRIPT_DIR`-relative output paths now resolve to
`legacy/data`, `legacy/figures`, `legacy/papers` — so it can no longer
silently overwrite the published-path outputs (`data/combined_cleaned.csv`,
`data/statistics_summary.csv`, `figures/`). Do not run it for reproduction;
use `analyze.py`. Moved 2026-10-26 (VOR audit).
