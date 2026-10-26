# Version of Record — Provenance

## Article identity

| Field | Value |
|---|---|
| Title | Association between the EU desflurane phase-out and secondary market vaporizer prices: An observational time-series analysis |
| Journal | Journal of Anesthesia and Translational Medicine (J Anesth Transl Med), ISSN 2957-3912 |
| DOI | 10.1016/j.jatmed.2026.08.002 |
| PII | S2957-3912(26)00034-3 |
| PMC ID | PMC13573186 (version PMC13573186.1) |
| Type | Research Article |
| Published | 2026-09-08 (issue 2026-12); received 2026-07-13, revised 2026-08-17, accepted 2026-09-15 per PMC dates |
| License | CC BY-NC-ND 4.0 |
| Authors | Tatsuki Onishi (Shiga University), Tatsuyoshi Ikenoue (Shiga University / University of Miyazaki) |

## Sources retrieved

| Source | URL | Retrieved (UTC) | SHA-256 | Status |
|---|---|---|---|---|
| Europe PMC fullTextXML | https://www.ebi.ac.uk/europepmc/webservices/rest/PMC13573186/fullTextXML | 2026-10-26 | 8b37d08a…9d37 | OK — authoritative full text used |
| PMC full-text HTML | https://pmc.ncbi.nlm.nih.gov/articles/PMC13573186/ | 2026-10-26 | 01be69b5…69b2 | OK — cross-checked against XML |
| ScienceDirect VOR PDF | https://www.sciencedirect.com/science/article/pii/S2957391226000343/pdfft | 2026-10-26 | — | NOT OBTAINED — bot-blocked (HTML interstitial) |
| Unpaywall metadata | https://api.unpaywall.org/v2/10.1016/j.jatmed.2026.08.002 | 2026-10-26 | — | is_oa=true, CC BY-NC-ND, best location = ScienceDirect |

Files persisted under `audit/vor_sources/` with `ledger.json` (machine-readable ledger).

## Authoritative basis for this audit

The Europe PMC `fullTextXML` of PMC13573186 is used as the authoritative text of the
published article. It is the publisher-deposited full text (JATS XML), not the earlier
in-repo JCA-format draft (`papers/jca_manuscript.*`) that a previous audit mistreated as
the article of record. The PMC HTML rendering was cross-checked for identical numerical
content. Supplementary material (Table S1 values) was read from the same PMC record.

## Key distinction established by retrieval

The published article **does** contain sensitivity analyses (bootstrap CIs, post-hoc
power simulation, difference-in-differences, and comparative interrupted time series)
and a Table S1 of asking-price summaries. A prior audit concluded these were "not
reported in the article"; that conclusion was wrong because it audited the wrong
document (the pre-revision JCA draft) and the wrong branch (`main` only — the
sensitivity code lived on the unmerged `devin/jatm-minor-revision-r1` branch).
See `PREVIOUS_AUDIT_RECONCILIATION.md`.
