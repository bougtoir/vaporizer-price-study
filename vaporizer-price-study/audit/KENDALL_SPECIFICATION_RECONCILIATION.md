# Kendall Specification Reconciliation

## Question

Does the VOR's Kendall τ spec ("across ordered regulatory phases", milestones
Apr 2022 → Jan 2026) match the code's Kendall test?

## Code facts (`create_jca_tables.py:compute_trend`, `analyze.py`)

- `kendalltau(period_num, price_usd)` — scipy default **τ-b** (tie-corrected).
- `period_num` ∈ {1,2,3,4,5}: five phases with left-exclusive boundaries
  2022-04-05, 2023-10-05, 2024-02-07, 2026-01-01.
- Phase 1 (before 2022-04-05) is **empty** — data start Mar 2023. So only four
  phases contain observations (2,3,4,5), which is why an earlier manuscript draft
  described "four phases".
- Two sales dated exactly 2023-10-05 fall into phase 3 (left-exclusive).

## VOR wording vs code

The VOR lists five regulatory milestones (proposal Apr 2022, Parliament Mar 2023,
trilogue Oct 2023, adoption Feb 2024, restriction Jan 2026) and says Kendall τ was
computed "across ordered regulatory phases". The code implements a 5-level phase
index aligned to those milestones, with the pre-proposal phase empty by
construction of the data window (Mar 2023 onward).

**Reconciliation**: the VOR wording is consistent with the code. The earlier
"four vs five phases" discrepancy concerned a *draft* manuscript, not the VOR.
The τ-b predictor is the **phase index**, not dates — a genuine and now
documented spec choice: τ is invariant to the within-phase date distribution and
treats phases as ordinal categories.

## Residual caveat (documented, not resolved)

- τ-b vs τ-a: ties in `price_usd` make τ-b the right choice and it is scipy's
  default; verified.
- Phase-index coding is ordinal: equal spacing of phases 2–5 is assumed. This is
  stated in the README spec section; the VOR does not spell it out, marked here
  as Recoverable-from-code rather than article-Stated.
