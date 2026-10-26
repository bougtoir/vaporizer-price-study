# Reconciliation of the Previous Audit's False Claim

## The false claim

The Phase 0–11 audit (PR #2, first revision) stated:

- README: "The published article does **not** report bootstrap confidence intervals,
  difference-in-differences, or interrupted time-series models."
- FINAL_QC_REPORT.md: "Bootstrap / DiD / CITS — NOT VERIFIED — no implementation in
  repo; not reported in article."
- Verdict: PASS WITH CONDITIONS.

**This was wrong.** The Version of Record (PMC13573186) reports all three
sensitivity analyses plus a post-hoc power simulation, and the original code that
produced them exists in this repository's git history.

## Root cause — two scope errors

1. **Wrong document treated as the manuscript.** The audit compared outputs against
   the in-repo `papers/jca_manuscript.*` — a pre-revision draft prepared for a
   different journal — rather than the published JATM article. The sensitivity
   analyses were added during the JATM revision (R1) and appear only in the VOR.
2. **Wrong ref searched for code.** Only `main` was inventoried. The sensitivity
   implementation lived on the unmerged branch `origin/devin/jatm-minor-revision-r1`
   (`revision_sensitivity_analysis.py` + its stored outputs `revision_sensitivity.json`,
   `revision_sensitivity_its.csv`), created for the revision response and never merged.

## Which previous conclusions are affected

| Previous statement | Status now |
|---|---|
| "bootstrap/DiD/CITS not in article or repo" | **Retracted** — false. Corrected here and in README/FINAL_QC_REPORT. |
| "NOT VERIFIED, absent by design" | **Replaced** — all four sensitivity analyses are now ORIGINAL-CODE REPRODUCED (see `PUBLISHED_RESULT_TRACEABILITY.md`). |
| Portability fixes, checksum reproduction of `combined_cleaned.csv`, N=158/321/554, primary statistics, figure consistency | **Unaffected** — those were verified against data and code, not against the wrong manuscript. |
| "methods text says four phases but code uses five" | **Refined** — see `KENDALL_SPECIFICATION_RECONCILIATION.md`; the VOR wording ("ordered regulatory phases", milestones Apr 2022 → Jan 2026) is consistent with a 5-level coding where phase 1 is empty. |
| Overall verdict PASS WITH CONDITIONS | **Upgraded** — with original sensitivity code recovered and all published numbers reproduced, the verdict becomes PASS (see `VOR_FINAL_QC_REPORT.md`). The "conditions" now reduce to: keep `analyze_old.py` quarantined and the VOR sources ledger. |

## Lesson

An audit of "what the article reports" must retrieve the actual Version of Record
first. Repository drafts are not the article. This error class is now covered by the
required `VERSION_OF_RECORD_PROVENANCE.md` step in this audit.
