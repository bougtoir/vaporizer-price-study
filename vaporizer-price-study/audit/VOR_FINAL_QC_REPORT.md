# VOR Final QC Report

Supersedes `FINAL_QC_REPORT.md` for VOR-scope questions. Audit date 2026-10-26.
VOR: PMC13573186 / DOI 10.1016/j.jatmed.2026.08.002.

## Component verdicts

| Component | Verdict | Evidence |
|---|---|---|
| VOR provenance | PASS | `VERSION_OF_RECORD_PROVENANCE.md`, `vor_sources/ledger.json` |
| Sample sizes (1033; 158/321/554) | PASS — ORIGINAL-CODE REPRODUCED | `combined_cleaned.csv` byte-identical regeneration |
| Spearman / Kendall / quarterly trends | PASS — ORIGINAL-CODE REPRODUCED | all published trend values match |
| Pre/post MW, Welch, Cohen's d | PASS — ORIGINAL-CODE REPRODUCED | incl. −31.4%, d=−0.47 |
| Bootstrap CIs (10k, seed 20260729) | PASS — ORIGINAL-CODE REPRODUCED | JSON regenerated identically; CI −134.53/−10.70 vs published −135/−11 |
| Post-hoc power (lognormal, 10k) | PASS — ORIGINAL-CODE REPRODUCED | 28.4% / 58.1% exact |
| DiD (OLS, HC3) | PASS — ORIGINAL-CODE REPRODUCED | 0.01784 / P=0.9463 |
| CITS (monthly medians, HAC-3) | PASS — ORIGINAL-CODE REPRODUCED | −0.8843 / +0.4465 / iso P=0.602 |
| Table S1 asking prices | PASS — NUMERICALLY CONSISTENT | values match; byte-rerun pending |
| Figures | PASS | prior audit: regenerated, consistent |
| Portability (cwd-independence) | PASS | 11/11 regression tests |
| Environment parity | PASS — NUMERICALLY CONSISTENT | results identical across pkg versions |
| ScienceDirect PDF vs PMC text | NOT VERIFIED | PDF bot-blocked; low residual risk |
| Buyer geography, equipment covariates | NOT VERIFIED | data nonexistent; article limitation |

## Previous-audit error correction

The claim "bootstrap/DiD/CITS not reported in the article" is **retracted** —
see `PREVIOUS_AUDIT_RECONCILIATION.md`. Cause: wrong manuscript (JCA draft) and
wrong ref (`main` only). Both earlier documents are corrected in place.

## Overall verdict

**PASS** (upgraded from PASS WITH CONDITIONS). All principal published numerical
results, including the reported sensitivity analyses, were reproduced using the
original analytical code. Supplementary Table S1 was verified for numerical
consistency (full regeneration not established).

- **PR safe to merge: YES** — changes are additive (path fixes, docs, tests,
  recovered original script) plus README corrections; no estimator, method, or
  result was altered.
- **Decision: MERGE-able** — left to the maintainer per instructions; this audit
  does not merge.
- Remaining conditions: keep `analyze_old.py` quarantined (never run it — it
  overwrites pipeline outputs); keep `audit/vor_sources/` intact.
