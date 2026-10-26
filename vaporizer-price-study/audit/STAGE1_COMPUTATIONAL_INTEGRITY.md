# Stage 1 Computational Integrity Audit

Scope: all scripts on the published analysis path — `analyze.py`,
`create_jca_tables.py`, `analyze_asking_prices.py`,
`revision_sensitivity_analysis.py` — plus a scan of every `*.py` in the project
dir. Date: 2026-10-26.

## Method

- Regex/AST-style scans for hard-coded empirical values: p-value literals,
  rho/tau/d literals, CI bound literals (−134.53, −10.70, −124, 0.284, 0.5808,
  0.0178, 0.884), and patterns like `p = 0.0x` in production code.
- Manual read of every constant assignment in the three principal scripts.
- Check that scripts recompute from raw CSVs rather than reading back their own
  stored outputs (`combined_cleaned.csv` is a legitimate intermediate: produced
  by `analyze.py` from the raw Terapeak files and consumed downstream as the
  documented pipeline input — including by `revision_sensitivity_analysis.py`).

## Findings

| Pattern searched | Result |
|---|---|
| Hard-coded p-values / effect sizes / CIs in analysis scripts | NONE found |
| Empirical constants (−134.53, 0.284, −0.884, …) anywhere in production code | NONE found |
| Scripts copying stored outputs instead of recalculating | NONE — all compute from CSV inputs |
| `np.random` / seeds outside sensitivity script | NONE; the only RNG is `default_rng(20260729)` in `revision_sensitivity_analysis.py` |

## Legitimate fixed parameters (verified, not empirical results)

| Constant | Location | Justification |
|---|---|---|
| `default_rng(20260729)` | revision_sensitivity_analysis.py:30 | published simulation seed |
| `N_BOOT = N_SIM = 10000` | :28–29 | published iteration counts |
| `BAN_DATE = 2026-01-01` | :27 | regulatory intervention date |
| Phase cut dates (2022-04-05 … 2026-01-01) | analyze.py / create_jca_tables.py | regulatory milestones |
| `count >= 3`, `len(qm) >= 4` | create_jca_tables.py:113 | quarterly-trend eligibility rule |
| `alpha`/`p < 0.05` | power simulation | significance threshold |
| `cov_type='HC3'`, `maxlags=3` | DiD / CITS | published variance estimators |
| Effect-size SE formula | create_jca_tables.py | standard Cohen's d SE |

## Regression-test reference values

`tests/test_reproducibility.py` asserts published values (−0.277, −0.121,
28.4%, …) as *expected targets* recomputed from raw data — this is the correct
use of published numbers: they are checked, never substituted into production
calculations.

## Verdict

PASS — no hard-coded empirical results, no output-copying shortcuts, no
unjustified constants on the published path.
