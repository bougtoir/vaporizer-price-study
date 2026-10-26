# Stage 1 Final QC

Date: 2026-10-26. Branch `devin/1793007296-vaporizer-reproducibility-audit`
(PR #2, open, unmerged). Every verdict backed by logs/docs under `audit/`.

| Check | Result | Evidence |
|---|---|---|
| Published VOR provenance documented | PASS | `VERSION_OF_RECORD_PROVENANCE.md`, `vor_sources/ledger.json` |
| Previous audit corrected | PASS | `PREVIOUS_AUDIT_RECONCILIATION.md` + Phase-2 overclaim fixes |
| Principal analyses reproduced | PASS | `stage1_analyze_stdout.txt`, `stage1_jca_tables_stdout.txt` |
| Bootstrap reproduced | PASS | `stage1_sensitivity_stdout.txt` [−134.53, −10.70] |
| Power simulation reproduced | PASS | 0.284 / 0.5808 |
| DiD reproduced | PASS | 0.01784 / p=0.9463 |
| CITS reproduced | PASS | −0.8843 / +0.4465 |
| Kendall specification accurate | PASS | `KENDALL_SPECIFICATION_RECONCILIATION.md`, code-verified |
| README technically complete | PASS | all 11 required items; sensitivity run command added; no acknowledgment |
| Portable execution confirmed | PASS | scripts ran from `/tmp`; `test_scripts_run_from_external_cwd` |
| Legacy output collision prevented | PASS | `legacy/` quarantine; paths resolve under `legacy/` |
| Computational integrity | PASS | `STAGE1_COMPUTATIONAL_INTEGRITY.md` |
| Data–text–figure consistency | PASS | `STAGE1_DATA_TEXT_FIGURE_CONSISTENCY.md` (Table S1 & figures: NUMERICALLY CONSISTENT) |
| Regression tests | PASS | 13/13 (`stage1_pytest.log`) |
| Adversarial review | PASS | `STAGE1_ADVERSARIAL_REVIEW.md` (1 doc defect found+fixed) |
| Published results unchanged | PASS | no estimator/spec/result modified; additive changes only |
| PR technically ready to merge | **YES** | maintainer decision; NOT merged |

## Overall verdict: **PASS**

All principal published numerical results, including the reported sensitivity
analyses, were reproduced using the original analytical code. Supplementary
Table S1 and figures were verified for numerical consistency. The
publisher-hosted PDF was not retrievable (PMC text used); environment differs
from the paper's stated versions without affecting outputs.

## Merge recommendation: **MERGE READY**

Remaining limitations (documented, not defects): ScienceDirect PDF undiffed;
package-version drift; DiD/CITS are non-confirmatory sensitivity analyses.
