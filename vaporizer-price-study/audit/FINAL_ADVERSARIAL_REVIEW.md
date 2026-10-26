# Phase 9 — Adversarial Review

Concrete falsification attempts against the portability fix and the
reproduction claims. Environment per `BASELINE_INVENTORY.md`.

| # | Attempted falsification | Outcome |
|---|---|---|
| 1 | Delete the `/home/ubuntu/vaporizer_research` symlink used for baseline accommodation, then run `analyze.py` + `create_jca_tables.py` from `/tmp` | PASS — both complete and write into the repo; no absolute path dependency remains |
| 2 | Search all `*.py` for any remaining local absolute paths (`/home/`, `vaporizer_research`, `C:\`, `/Users/`) | PASS — none remain; enforced by `tests/test_reproducibility.py::test_no_machine_specific_paths_in_scripts` |
| 3 | Check whether `combined_cleaned.csv` in the repo was stale vs raw inputs | PASS — byte-identical after genuine regeneration from `terapeak_*.csv` |
| 4 | Check whether `statistics_summary.csv` was stale | PASS with note — identical except last-ulp float digits in `t_pval` (scipy 1.15.3 vs the version that produced the committed file); all reported digits unchanged |
| 5 | Hunt hard-coded results: grep numeric literals matching published stats inside scripts; inspect write_* emitters | PASS — every printed/written statistic is an interpolated computed value; only prespecified dates, labels, keywords, styling are literal |
| 6 | Kendall independent variable: swap `period_num` for `date_ord` to see whether targets were produced by dates | Confirmed code uses `period_num` (1–5); using dates would *not* reproduce τ=−0.121/p=0.049 — phase coding is genuinely what was published |
| 7 | Kendall variant ambiguity | Verified `scipy.stats.kendalltau` signature default `variant='b'` on scipy 1.15.3 (and documented tau-b in README); p=0.0492 target reproduced exactly under tau-b |
| 8 | Phase-boundary edge case: 2 sales dated exactly 2023-10-05 | They fall in phase 3 under the `d < boundary` rule (left-exclusive). This is the implemented (and now documented) convention; counts reproduce |
| 9 | Phase 1 empty (data start 2023-03-28 > 2022-04-05): does tau see only 4 distinct levels? | Yes — `period_num` ∈ {2,3,4,5}; tau-b handles ties/levels correctly; documented in README |
| 10 | Quarterly test uses same IV as Spearman? | No — it uses the quarter *sequence index* over quarters with `count>=3` (≥4 quarters required), not sale dates. Documented; values reproduce |
| 11 | Hidden seeds / stochastic components in published path | None — no `np.random`, bootstrap, or resampling calls exist anywhere in the repo (`rg` audit); the pipeline is fully deterministic |
| 12 | Missing files / optional inputs: run `create_jca_tables.py` without `ebay_asking_prices.csv` | Table S1 skipped gracefully; tables 1–2 unaffected |
| 13 | `analyze_old.py` output-clobber hazard | REAL ISSUE — it writes the same `combined_cleaned.csv`/`statistics_summary.csv`/`figures` filenames with a *different* pipeline. Mitigated by README warning + retained for provenance; outputs restored to published-path versions on this branch |
| 14 | `write_papers.py` (legacy BMJ generator) embeds figures and emits docx far larger than committed versions | Confirmed legacy drift — committed `vaporizer_paper_*.docx` were generated under an older figure set; regenerated outputs were NOT committed (kept minimal diff). Legacy script still runs from any cwd |
| 15 | Dependency claims | `requirements.txt` bounds are based on actually-tested versions (this VM); no claim of pandas-3.x support is made in-code, only that an external rerun reportedly succeeded on pandas 3.1/Python 3.12 (documented, unverified here) |
| 16 | Manuscript-vs-code mismatch hunting (Phase 6) | One wording quirk found: Methods says "four phases" while code uses a 1–5 scheme with an empty phase 1; no numeric impact. Reported, not rewritten |

## Limitations of this review

- Figure pixel-level equivalence was not asserted (matplotlib version rendering
  differences are expected); plotted data are identical.
- The live scrapers were not re-executed (require live marketplace access);
  they are out of the published analysis path.
- Bootstrap / DiD / CITS remain NOT VERIFIED — absent from repo and article.
