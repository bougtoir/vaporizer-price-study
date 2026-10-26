# Published Analysis Inventory (VOR, DOI 10.1016/j.jatmed.2026.08.002)

Inventory basis: Europe PMC fullTextXML of PMC13573186 (authoritative full text).
Spec status legend: **Stated** = explicit in article; **Recoverable** = explicit in the
recovered original code (`revision_sensitivity_analysis.py`, branch
`devin/jatm-minor-revision-r1`, now merged into this audit branch); **Inferred** = deduced
from outputs; **Unknown** = not determinable from text or code.

## 1. Sample / data
- Source: eBay Terapeak completed sales, Mar 2023 – Mar 2026 (Stated).
- N = 1033 completed sales: 158 desflurane, 321 sevoflurane, 554 isoflurane (Stated).
- Filter: `price_usd > 0` (Recoverable — `load_data()`).

## 2. Primary trend analyses (per agent)
| Analysis | IV / DV | Estimator | Spec status |
|---|---|---|---|
| Spearman rank correlation | DV `price_usd`; IV = sale date (ordinal day) | `scipy.stats.spearmanr` | Stated (method) + Recoverable (predictor = date, in `create_jca_tables.py:compute_trend`) |
| Kendall τ across ordered regulatory phases | DV `price_usd`; IV = phase index 1–5 | `scipy.stats.kendalltau` → τ-b | Recoverable (code) — see `KENDALL_SPECIFICATION_RECONCILIATION.md` |
| Quarterly median trend | DV = quarterly median price; IV = quarter index | `spearmanr` on quarters with count ≥ 3, needs ≥ 4 quarters | Recoverable |
| Phase boundaries | five phases, left-exclusive cut dates 2022-04-05, 2023-10-05, 2024-02-07, 2026-01-01 | hard-coded in `analyze.py`/`create_jca_tables.py` | Recoverable |

## 3. Pre-/post-ban comparison
- Ban date: 2026-01-01 (Stated in article; `BAN_DATE` in code).
- Mann–Whitney U + Welch t + Cohen's d (Stated). Effect-size SE formula
  `sqrt((n1+n2)/(n1*n2) + d²/(2(n1+n2)))` (Recoverable).

## 4. Sensitivity analyses (in VOR, previously declared absent — WRONG)
| Analysis | Spec | Status |
|---|---|---|
| Bootstrap pre/post CIs | 10,000 resamples with replacement within pre and post groups; percentile 95% CIs for mean and median differences; per-iteration MW and Welch p, medians reported; seed `np.random.default_rng(20260729)` | Stated (iterations) + Recoverable (resampling scheme, seed, estimator) |
| Post-hoc power simulation | MLE lognormal (mu, sigma) per group on log prices; N_SIM = 10,000 simulated pairs at observed n (141/17 des); fraction with p < 0.05 | Stated ("lognormal") + Recoverable (MLE + seed) |
| Difference-in-differences | OLS on `log_price`, transaction level: `~ C(agent_type) + months_since_start + post + post_desflurane + post_isoflurane + trend_desflurane + trend_isoflurane`; HC3 robust SE; sevoflurane reference | Recoverable (formula + cov_type in code); article reports coef/P only |
| CITS | OLS on monthly `log_median`: `~ C(agent_type) + time_months + post + time_after + post_desflurane + post_isoflurane + trend_after_desflurane + trend_after_isoflurane`; HAC (Newey–West maxlags=3) SE; monthly medians over all months incl. count=1 | Recoverable; article reports coef/P only |

## 5. Supplementary Table S1
- Asking-price medians $178 / $1,277 / $556 (des/sev/iso), KW H = 66.0 (P<0.001),
  bid–ask spreads 43%/184%/178% (Stated; generator script = `create_jca_tables.py`
  asking-price section — verified consistent in prior audit).

## 6. Software environment
- VOR: Python 3.12, pandas 2.2, scipy 1.14, statsmodels 0.14 (Stated).
- Audit environment: Python 3.10.12, pandas 2.3.3, scipy 1.15.3, statsmodels 0.15.0 —
  all outputs reproduced identically despite version drift (see traceability matrix).

## Items genuinely absent from both VOR and code
- Item-level geographic buyer/seller data (article acknowledges as limitation).
- Equipment age/condition controls (acknowledged limitation).
- Nothing else identified.
