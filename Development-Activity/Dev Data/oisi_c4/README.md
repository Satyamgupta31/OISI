# OISI Component 4 — extraction and scoring pipeline

Window: 1 Oct 2025 00:00 UTC (inclusive) to 1 Oct 2026 00:00 UTC (exclusive). All rules live in `rules.py`.

## Files

| File | What it is |
| --- | --- |
| `repos.csv` | All 112 list entries: decision, primary repo(s), upstream (forks), issue tracker, `issue_repo`, `scored` yes/no |
| `rules.py` | Window, cohort cut-offs, bot list and patterns, identity-resolution rules R1–R6 |
| `git_metrics.py` | Metrics 1–7 + diagnostics from git history (any host) |
| `gh_extract.py` | Metrics 8-13 + responsiveness diagnostics from GitHub GraphQL, plus star history from GitHub's REST star-history endpoint used by Star History |
| `score.py` | Normalisation, weights, missing-data rule → `scores.csv` |
| `analysis.py` | Weight sensitivity and Spearman correlation |
| `identity_map.csv` | Manual identity merges (repo, email, canonical_id, evidence) — empty template |
| `manual_metrics.csv` | 16 repos whose issues are not on GitHub (Jira, GitLab, none): fill by hand or leave blank |
| `git_metrics.csv` | Commit metrics already computed for 90 of 96 scored repos |
| `scores_git_only.csv` | Partial scores (60% of weight); `component4_score` is blank by design |
| `git_only_*` | Sensitivity and correlation results on the commit metrics |
| `test_gh_logic.py` | Offline tests for the extractor's metric logic |

## Run it (on your own machine)

Requirements: Python 3.9+, git, internet. No extra packages.

```bash
# 1. commit metrics for all scored repos (also reaches GitLab / Gitea / Bitbucket, ~15 min)
python git_metrics.py repos.csv clones git_metrics.csv

# 2. GitHub issues, PRs, stars, forks (~30-60 min; cached in raw/ so re-runs are cheap)
export GITHUB_TOKEN=github_pat_...   # fine-grained token, public repositories read-only
python gh_extract.py repos.csv gh_metrics.csv

# 3. score + robustness checks
python score.py git_metrics.csv gh_metrics.csv manual_metrics.csv scores.csv
python analysis.py git_metrics.csv gh_metrics.csv final
```

Windows PowerShell: use `$env:GITHUB_TOKEN="github_pat_..."` instead of `export`.

## Decisions still open

1. **Oasisbr** — default branch has 0 own commits in the window; 88 sit on `feat/migration-vufind7-to-11`. Set `branch` in `repos.csv` if you want that branch scored.
2. **D4Science** — choose a core repository or leave excluded.
3. **Crossref, DMP Tool, Europe PMC, ROR issues, Dryad issues, fcrepo issues** — marked "verify" in `repos.csv`.
4. **Redundant metrics** — commits ~ active weeks (ρ 0.87) and contributors ~ effective contributors (ρ 0.88) among active repos. Consider dropping one of each pair after the full run.
