# Correspondence Implications

What this VOR audit changes for external communication about the repository.

## Facts that bear on the correspondence

1. An external researcher (see prior correspondence) attempted an independent
   rerun and could not locate the bootstrap / DiD / CITS implementations.
   **He was right about `main`**: the code genuinely was not on the default
   branch — it sat unmerged on `devin/jatm-minor-revision-r1`. His inability to
   find it was a repository-organization defect, not a failure on his part.
2. The code did exist and does produce the published numbers exactly. So the
   article's reported analyses are real and reproducible — but only after this
   audit merged the script into the public branch.
3. The previous audit's statement that the article lacks these analyses was
   wrong and is now retracted in-repo.
4. The Kendall spec (τ-b on a five-level phase index, phase 1 empty) was
   under-documented in the repo and is now spelled out — answering his most
   specific technical question.
5. PR #2 is open but **not merged**; correspondence must not imply it is.

## Recommended posture for the reply

- Thank him; his report directly caused the discovery that the sensitivity
  code was unreachable from `main`.
- State plainly: the analyses exist, code is now in the public repo (pending
  merge of the open PR), specs documented.
- Explain the Kendall spec precisely (τ-b, ordinal phase index 1–5,
  left-exclusive boundaries, phase 1 empty).
- Do not criticize him for failing to find code that was unreachable.
- Do not overclaim: ScienceDirect PDF not diffed; environment versions differ
  from the paper's stated ones (results nonetheless identical).
- Thank him for the Journal of Reproducible Health Sciences invitation;
  express interest in computational verification as a model; do not commit to
  submission.
- Request his run log and software environment — useful for documenting
  cross-environment reproduction.
