# VOR Adversarial Review

Adversarial pass against the VOR audit itself: ten challenges an hostile reviewer
could raise, and the evidence either way.

| # | Challenge | Finding |
|---|---|---|
| 1 | The "VOR" used was the PMC deposit, not the ScienceDirect PDF — maybe they differ. | PMC record is the publisher-deposited full text (PMC13573186.1); EuropePMC XML and PMC HTML agree. ScienceDirect PDF unreachable to curl (HTML interstitial) — residual risk documented, low. |
| 2 | Sensitivity results could have been *written into* `revision_sensitivity.json` by hand rather than computed. | Re-running `revision_sensitivity_analysis.py` regenerates the JSON identically from `combined_cleaned.csv`; seeded RNG makes it deterministic. Not hand-written. |
| 3 | Seed was cherry-picked to match the article. | JSON on R1 branch predates this audit; the stored values already matched VOR before any code ran in this session. No tuning performed. |
| 4 | CITS monthly panel includes months with count=1 — fragile medians. | True (nobs=111 ≈ 37 months × 3 agents); documented as the published spec, not altered. Flagged as a methodological caveat, not a reproducibility failure. |
| 5 | DiD parallel-trends assumption is violated (pre-trend diff P=0.001). | Correct — the VOR itself reports the pre-trend differential and labels DiD/CITS "sensitivity, not confirmatory". Consistent, not hidden. |
| 6 | Different package versions could change p-values. | Audit env (scipy 1.15.3, statsmodels 0.15.0) differs from VOR (1.14 / 0.14) yet all principal published numerical results reproduced exactly (Table S1 verified for numerical consistency). Version drift is non-fatal here. |
| 7 | `analyze_old.py` could silently overwrite outputs with a different pipeline. | Real hazard, confirmed in Phase 0–11. Outputs restored; README warns; the QC report marks it quarantined-not-deleted (history preservation). |
| 8 | Kendall on a 5-level phase index where phase 1 is empty is effectively 4 phases — is the VOR misleading? | Resolved in `KENDALL_SPECIFICATION_RECONCILIATION.md`: VOR's "ordered regulatory phases" is consistent with the code; the "four phases" wording was a draft artifact. |
| 9 | Asking-price Table S1 uses Terapeak listings data not in the repo — unverifiable. | `create_jca_tables.py` reads an asking-price dataset that is in `data/`; values verified consistent. Marked NUMERICALLY CONSISTENT, not ORIGINAL-CODE REPRODUCED, pending byte-identical rerun — honest classification. |
| 10 | The audit branch changed repo layout (audit/, tests/) — does anything still run from the documented entry points? | Regression tests execute `analyze.py` and `create_jca_tables.py` from an external cwd; 11 tests pass; `revision_sensitivity_analysis.py` executed from `/tmp` successfully. |

## Net verdict

No challenge overturns the reproduction. The strongest residual caveats:
(1) ScienceDirect PDF not diffed against PMC text; (2) DiD/CITS assumptions
are weak by the article's own admission — a methods caveat, not a
reproducibility defect.
