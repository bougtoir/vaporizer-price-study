# Stage 1 Baseline — Current State Record

Recorded 2026-10-26 (JST) on branch `devin/1793007296-vaporizer-reproducibility-audit`
(PR #2, open, unmerged).

## Commits

| Item | SHA |
|---|---|
| Base (merge-base with `origin/main`) | `cac26f92ab61d41b630f4e85465db27988e2632f` |
| PR head at baseline | `a79c6c4a41801ec98f364e83f9b57d23216bc2cc` |

## Repository structure (relevant parts)

```
vaporizer-price-study/           repo root
├── README.md, requirements.txt, FINAL_HANDOFF.txt, VOR_FINAL_HANDOFF.txt, .gitignore
└── vaporizer-price-study/       project dir
    ├── analyze.py               main pipeline (published-path)
    ├── create_jca_tables.py     tables + trend tests (Spearman/Kendall/quarterly)
    ├── revision_sensitivity_analysis.py  recovered sensitivity code (bootstrap/power/DiD/CITS)
    ├── data/                    raw Terapeak CSVs + generated outputs + revision_sensitivity.{json,its.csv}
    ├── figures/  papers/  tests/  audit/
    ├── legacy/analyze_old.py    quarantined legacy pipeline
    └── other generation/scraper scripts (provenance only)
```

## Changed files vs base (51 total, abridged)

- New: `requirements.txt`, `.gitignore`, `FINAL_HANDOFF.txt`, `VOR_FINAL_HANDOFF.txt`,
  `tests/test_reproducibility.py`, `revision_sensitivity_analysis.py`,
  `data/revision_sensitivity.json`, `data/revision_sensitivity_its.csv`,
  `legacy/{analyze_old.py,README.md}` (rename R097 from repo root),
  `audit/` (19 evidence files incl. `vor_sources/` with VOR text + SHA-256 ledger).
- Modified: 17 analysis/generation scripts — `SCRIPT_DIR`-relative path handling only.
- Modified: `README.md` — spec documentation + corrected sensitivity claims.

## Verification of preserved improvements (confirmed, not assumed)

- `revision_sensitivity_analysis.py` present on branch: seed
  `np.random.default_rng(20260729)`, `N_BOOT = 10000`, `N_SIM = 10000`,
  `cov_type='HC3'` (DiD), `cov_type='HAC', maxlags=3` (CITS) — all confirmed in code.
- `legacy/analyze_old.py` exists under `legacy/`; its `SCRIPT_DIR`-derived paths
  resolve under `legacy/` and cannot collide with `data/`/`figures/`/`papers/`.
- No `/home/ubuntu/vaporizer_research` references remain in scripts.
- `tests/test_reproducibility.py`: 13 tests present.

## Outstanding issues at baseline

- Three prior-audit docs still contain overbroad wording ("every number in the
  published article is reproduced") — requires precision correction (Phase 2).
- `DRAFT_REPLY_TO_DEGENFELLNER.txt` exists in `audit/` from the previous task —
  retained as prior-task deliverable; nothing was sent and no acknowledgment was
  added to README (Stage-1 scope forbids new correspondence work).
