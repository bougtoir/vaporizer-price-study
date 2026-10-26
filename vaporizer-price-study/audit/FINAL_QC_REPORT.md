# Phase 10 — Final QC Report

| Item | Verdict | Evidence |
|---|---|---|
| Baseline inventory | PASS | `audit/BASELINE_INVENTORY.md`, SHA, checksums, stdout |
| Portability (no machine-specific paths) | PASS | `rg` sweep empty; `test_no_machine_specific_paths_in_scripts` |
| Launch from external working directory | PASS | `test_scripts_run_from_external_cwd`; manual run from `/tmp` |
| Correct Kendall predictor (ordinal phase 1–5) | PASS | `create_jca_tables.py:108`; `test_trend_tests_use_correct_predictors` |
| Verified tau variant | PASS | scipy `variant='b'` default confirmed; tau-b documented |
| README phase boundaries | PASS | explicit 1–5 table, left-exclusive rule, anchors sourced from code |
| Spearman + quarterly specification | PASS | date-ordinal Spearman; quarter-index median rho documented |
| Descriptive statistics vs targets | PASS | N 158/321/554; means/SDs/pre/post n; % changes all exact |
| Trend tests vs targets | PASS | all 9 rho/tau + p-values reproduce to reported precision |
| Bootstrap | ~~NOT VERIFIED~~ → ORIGINAL-CODE REPRODUCED (VOR audit) | recovered `revision_sensitivity_analysis.py`; published CIs match exactly |
| Difference-in-differences | ~~NOT VERIFIED~~ → ORIGINAL-CODE REPRODUCED (VOR audit) | coef 0.01784, P=0.9463 = VOR |
| CITS | ~~NOT VERIFIED~~ → ORIGINAL-CODE REPRODUCED (VOR audit) | −0.8843/+0.4465 = VOR |

| Tables and figures | PASS | jca_table{1,2,s1}.docx regenerated; fig1–6 regenerated (rendering may differ by version; plotted data identical) |
| Computational integrity | PASS | `audit/COMPUTATIONAL_INTEGRITY_AUDIT.md`; no hard-coded results |
| Manuscript consistency | PASS | `audit/DATA_TEXT_FIGURE_CONSISTENCY.md`; all statements reproduce or match after rounding |
| Regression tests | PASS | 11/11 tests green (`pytest tests/`) |
| Adversarial review | PASS | `audit/FINAL_ADVERSARIAL_REVIEW.md`; 16 falsification attempts, 1 real hazard documented (`analyze_old.py` output clobber) |
| No unintended scientific changes | PASS | diff limited to path constants, README, requirements, tests, audit docs; estimands/methods/results untouched |

**Overall verdict: PASS WITH CONDITIONS**

Conditions:
1. `analyze_old.py` writes to the same generated filenames with a different
   pipeline; run order matters (README warns; outputs on this branch are the
   published-path versions).
2. ~~Bootstrap/DiD/CITS are NOT VERIFIED~~ — CORRECTED (VOR audit): all four
   sensitivity analyses are ORIGINAL-CODE REPRODUCED; see `VOR_FINAL_QC_REPORT.md`.
3. `statistics_summary.csv` regeneration can differ in last-ulp digits across
   scipy versions; reported precision is stable.
