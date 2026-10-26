# Stage 1 Adversarial Review

External-reviewer pass over the revised PR, 2026-10-26. Previous verdicts were
NOT carried forward; each claim was re-derived from code and execution logs.

| # | Attack | Result |
|---|---|---|
| 1 | Overbroad "every number reproduced" claim | **Defect found and fixed** — replaced with precise tiers in VOR_FINAL_QC_REPORT, PUBLISHED_RESULT_TRACEABILITY, VOR_ADVERSARIAL_REVIEW, PREVIOUS_AUDIT_RECONCILIATION, VOR_FINAL_HANDOFF, README. Correct statement: all principal published numerical results reproduced by original code; Table S1 numerically consistent. |
| 2 | Missing analytical scripts | None — 12/12 published analyses have implementations (see coverage table below). |
| 3 | Undocumented dependencies | `requirements.txt` present; statsmodels needed by sensitivity script is listed. PASS. |
| 4 | Machine-specific paths | `rg /home/|vaporizer_research` finds only test assertions forbidding them. PASS. |
| 5 | Statistical spec discrepancies | Kendall τ-b vs phase index, left-exclusive boundaries, quarterly count≥3/≥4 rule — all verified in code and match README/VOR. PASS. |
| 6 | Incorrect Kendall description | Corrected in VOR audit; README explicit (τ-b, ordinal 1–5, phase 1 empty). PASS. |
| 7 | Missing regulatory-phase definition | README table lists all 5 phases + left-exclusive rule + empty phase-1 note. PASS. |
| 8 | Legacy output collision | `legacy/analyze_old.py`: SCRIPT_DIR resolves under `legacy/`; standard workflow cannot clobber `data/`/`figures/`. Verified path constants derive from file location. PASS. |
| 9 | Numerical hard-coding | Scan found none; constants are seeds/dates/thresholds/spec choices only (STAGE1_COMPUTATIONAL_INTEGRITY). PASS. |
| 10 | Inconsistent published vs reproduced | All spot values exact at reported precision (STAGE1_DATA_TEXT_FIGURE_CONSISTENCY). PASS. |
| 11 | Misleading DiD/CITS interpretation | Docs explicitly label them sensitivity analyses with unmet causal assumptions (pre-trend p=0.001). PASS. |
| 12 | Incomplete/inaccurate audit docs | Table S1 downgraded to NUMERICALLY CONSISTENT; figure pixel-equality not claimed; ScienceDirect PDF gap recorded. PASS after Phase-2 corrections. |

## Coverage check (published analysis → implementation)

descriptives→`analyze.py`; MW/Welch/Cohen's d→`analyze.py`+`create_jca_tables.py`;
Spearman/Kendall/quarterly→`create_jca_tables.py`; bootstrap/power/DiD/CITS→
`revision_sensitivity_analysis.py` (executed, outputs identical to stored JSON).

## New defects introduced by this audit

None material. Minor: `DRAFT_REPLY_TO_DEGENFELLNER.txt` (prior-task deliverable)
still in `audit/` — retained intentionally; unsent, and README contains no
acknowledgment. Logged here for transparency.

## Verdict

PASS — one real documentation defect found and corrected; no analytical defects.
