# Sensitivity Analysis Recovery

## What was recovered

From branch `origin/devin/jatm-minor-revision-r1` (never merged to `main`):

| File | Content |
|---|---|
| `vaporizer-price-study/revision_sensitivity_analysis.py` | Original implementation: bootstrap CIs, post-hoc power, DiD, CITS (235 lines) |
| `vaporizer-price-study/data/revision_sensitivity.json` | Stored outputs of that script |
| `vaporizer-price-study/data/revision_sensitivity_its.csv` | Monthly median panel used by CITS |

Recovery commits on that branch: `2dd3d42`, `f306158`, `f3655e3`, `1cd2a96`,
`9809a8e`, `645bb8f`. The same branch also carries the JATM R1 manuscript sources
(`create_jatm_*.py`, `papers/jatm_submission_revision_20260729.zip`).

## Recovered specification (from code, not inferred)

```python
BAN_DATE = pd.Timestamp('2026-01-01')
N_BOOT = 10000; N_SIM = 10000
RNG = np.random.default_rng(20260729)   # single seeded generator
```

- **Bootstrap**: resample pre and post groups separately with replacement;
  per iteration record mean diff, median diff, MW p, Welch p; report 2.5/97.5
  percentiles and median p.
- **Power**: fit lognormal (mu, sigma) by MLE on `log(x)` per group; simulate
  10,000 pre/post pairs at observed n; power = fraction with p < 0.05.
- **DiD**: transaction-level OLS on `log_price` with agent-specific post and
  trend interactions; `cov_type='HC3'`; sevoflurane reference.
- **CITS**: monthly median panel (all months, incl. count=1); OLS on
  `log_median` with level/slope intervention terms per non-reference agent;
  `cov_type='HAC', maxlags=3`.

## Execution verification (2026-10-26)

The recovered script was run unmodified against `data/combined_cleaned.csv` and
regenerated `revision_sensitivity.json` **identically** to the stored R1 copy
(bootstrap/paths deterministic via the seeded RNG; DiD/CITS deterministic).
Every published VOR number matches (see `PUBLISHED_RESULT_TRACEABILITY.md`).

## Why independent reimplementation was not needed

Task Phase 5 requires reimplementation only "if spec identifiable" and original
code unavailable. The original code was recovered and reproduces the published
numbers exactly, so all four analyses classify **ORIGINAL-CODE REPRODUCED**. No
reimplementation, no post-hoc tuning, no hard-coded values.
