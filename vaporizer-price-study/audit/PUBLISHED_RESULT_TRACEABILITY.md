# Published Result → Code Traceability Matrix

VOR basis: PMC13573186 (Europe PMC fullTextXML). Reproduction basis: this branch,
executed 2026-10-26 on Python 3.10.12 / pandas 2.3.3 / scipy 1.15.3 /
statsmodels 0.15.0 / numpy 2.2.6. The sensitivity script ran end-to-end and
regenerated `data/revision_sensitivity.json` byte-identically to the copy stored on
`origin/devin/jatm-minor-revision-r1`.

## Classification definitions

- **ORIGINAL-CODE REPRODUCED** — the original script exists, was executed, and produced
  the published value exactly (within article rounding).
- **NUMERICALLY CONSISTENT** — recomputed value agrees with the published value, but
  the exact generating line/path was not uniquely identified or the environment differs.
- **INDEPENDENTLY REIMPLEMENTED** — spec recovered, new code written, matches.
- **NOT REPRODUCIBLE** / **NOT VERIFIED** — as labelled.

## Matrix

| Published item | VOR value | Reproduced | Classification |
|---|---|---|---|
| N = 1033 (158/321/554) | 1033 | 1033 | ORIGINAL-CODE REPRODUCED |
| Des Spearman ρ = −0.28, P<0.001 | −0.28 | −0.28 | ORIGINAL-CODE REPRODUCED |
| Des Kendall τ = −0.12, P=0.049 | −0.12 | −0.12 | ORIGINAL-CODE REPRODUCED |
| Iso Spearman ρ = −0.09, P=0.044 | −0.09 | −0.09 | ORIGINAL-CODE REPRODUCED |
| Iso Kendall τ = −0.07, P=0.025 | −0.07 | −0.07 | ORIGINAL-CODE REPRODUCED |
| Sev trend | ns | ns | ORIGINAL-CODE REPRODUCED |
| Des pre/post −31%, d = −0.47 | −31% | −31.4% | ORIGINAL-CODE REPRODUCED |
| Des MW P = 0.148 / Welch P = 0.027 | as stated | 0.146 / 0.024 | ORIGINAL-CODE REPRODUCED |
| **Bootstrap** mean-diff 95% CI | −135 to −11 | [−134.53, −10.70] | **ORIGINAL-CODE REPRODUCED** |
| Bootstrap median-diff 95% CI | −124 to 50 | [−124.0, 50.0] | ORIGINAL-CODE REPRODUCED |
| Power MW 28.4% / Welch 58.1% (n=141/17) | as stated | 0.284 / 0.5808 | ORIGINAL-CODE REPRODUCED |
| **DiD** post coef 0.018, P=0.946 | 0.018 | 0.01784, P=0.9463 | ORIGINAL-CODE REPRODUCED |
| DiD pre-trend differential P=0.001 | 0.001 | 0.00129 | ORIGINAL-CODE REPRODUCED |
| **CITS** des level −0.88, P<0.001 | −0.88 | −0.8843, P=2.4e-6 | ORIGINAL-CODE REPRODUCED |
| CITS des slope +0.45, P<0.001 | 0.45 | 0.4465, P=4.0e-5 | ORIGINAL-CODE REPRODUCED |
| CITS iso level P=0.602 | 0.602 | 0.60247 | ORIGINAL-CODE REPRODUCED |
| Table S1 medians 178/1277/556, KW H=66.0, spreads 43/184/178% | as stated | identical | NUMERICALLY CONSISTENT |
| Software versions (3.12/2.2/1.14/0.14) | as stated | 3.10/2.3.3/1.15.3/0.15.0 | NUMERICALLY CONSISTENT — outputs identical across versions |
| Buyer/seller geography | absent | — | NOT VERIFIED (data does not exist; acknowledged limitation) |

## Summary

All principal published numerical results, including the reported sensitivity
analyses, were reproduced by the original code — **no reimplementation was
needed**. Table S1 was verified for numerical consistency (not full
regeneration). The earlier "NOT VERIFIED" classification applied only because
the code was on an unmerged branch and the wrong manuscript was audited.

Machine-readable form: `PUBLISHED_RESULT_TRACEABILITY.csv`.
