# OISI Component 4 Scoring Methodology and Data Provenance

Generated from the current Component 4 pipeline in `Dev Data/oisi_c4`.

## 1. Scope

This document explains how each Component 4 parameter is scored and where the data comes from. It matches the current pipeline files:

- `repos.csv`: master repository list, resolution decisions, issue tracker mapping, resolved URL markers.
- `git_metrics.py` -> `git_metrics.csv`: git-derived metrics 1-7 and git diagnostics.
- `gh_extract.py` -> `gh_metrics.csv`: GitHub issue, PR, star, fork metrics 8-13 and API diagnostics.
- `manual_metrics.csv`: optional override/fill file for non-GitHub trackers or manually collected values.
- `score.py` -> `scores.csv`: normalized metric scores, availability, final Component 4 score.
- `OISI_Component4_Final_Master_With_Provenance.xlsx`: final joined workbook containing all 112 main-list entries and provenance columns.

## 2. Observation Window and Common Rules

All time-windowed metrics use the same window:

```text
Start: 2025-10-01 00:00 UTC, inclusive
End:   2026-10-01 00:00 UTC, exclusive
```

Common processing rules:

- Git metrics use the repository default branch unless `repos.csv` specifies a `branch`.
- Commit dates use git author date.
- Merge commits are excluded from authorship metrics, but counted separately as diagnostics.
- Bots are removed before human commit, contributor, issue, and PR response calculations.
- Bot detection comes from `rules.py`: GitHub user type `Bot`, named bot accounts, and bot-name/email patterns.
- Identities are resolved within a repository using `.mailmap`, identical email, GitHub no-reply login, GitHub login where available, normalized full-name matches, and optional `identity_map.csv`.
- Multi-repo products in `repos.csv` use semicolon-separated repositories; commits are de-duplicated by commit hash across the set.
- Fork-specific inherited upstream commits can be excluded when `repos.csv` provides an `upstream`.

## 3. Score Eligibility and Missing Data

Each metric is normalized to a 0-100 score. The final Component 4 score is computed only if the row has enough available data.

Availability rule from `score.py`:

- At least 75% of total metric weight must be available.
- At least half of each dimension's weight must be available.
- If the rule is met, `component4_score = score_on_available`.
- If the rule is not met, `component4_score` is blank and `score_status = insufficient data`.

Dormant repositories are not automatically missing. For git metrics, zero activity scores as zero where appropriate. In `score.py`, `days_since_last` and `eff_contributors` are treated as zero-score values for git-complete dormant repositories.

## 4. Dimensions and Weights

| Dimension | Metrics | Total Weight |
| --- | --- | ---: |
| Continuity | m1, m2, m3, m4 | 35% |
| Contributors | m5, m6, m7 | 25% |
| Responsiveness | m8, m9, m10, m11 | 35% |
| Adoption | m12, m13 | 5% |

Total weight is 100%.

## 5. Metric-by-Metric Documentation

### m1 Human Commits

| Field | Value |
| --- | --- |
| Weight | 10% |
| Dimension | Continuity |
| Raw input column | `git_metrics.csv::human_commits` |
| Score output column | `scores.csv::m1_commits` |
| Data source | Local bare git clone made by `git_metrics.py` from `repos.csv::repo` |
| Raw definition | Count of non-merge, non-bot commits in the observation window |
| Score formula | `100 * min(1, log10(1 + c) / log10(1001))` |
| Cap | 1000 commits gives score 100 |

Notes:

- Merge commits are excluded.
- Bot commits are excluded and reported separately as `bot_commits`.
- For multi-repo products, commit hashes are de-duplicated across repositories.

### m2 Active Weeks

| Field | Value |
| --- | --- |
| Weight | 15% |
| Dimension | Continuity |
| Raw input column | `git_metrics.csv::active_weeks` |
| Score output column | `scores.csv::m2_active_weeks` |
| Data source | Local git history via `git_metrics.py` |
| Raw definition | Number of distinct weeks, out of 52, with at least one human non-merge commit |
| Score formula | `active_weeks / 52 * 100` |

Notes:

- Weeks are counted from the window start.
- A repository active every week scores 100.

### m3 Recency

| Field | Value |
| --- | --- |
| Weight | 5% |
| Dimension | Continuity |
| Raw input column | `git_metrics.csv::days_since_last` |
| Score output column | `scores.csv::m3_recency` |
| Data source | Local git history via `git_metrics.py` |
| Raw definition | Days from the latest human non-merge commit in the window to the cutoff date |
| Score formula | 100 if <= 30 days; linear decay to 0 at 365 days |

Formula:

```text
if days_since_last <= 30: score = 100
else: score = max(0, (365 - days_since_last) / 335 * 100)
```

### m4 Release Activity

| Field | Value |
| --- | --- |
| Weight | 5% |
| Dimension | Continuity |
| Raw input column | `git_metrics.csv::tags_in_window` |
| Score output column | `scores.csv::m4_releases` |
| Data source | Local git tags via `git_metrics.py` |
| Raw definition | Number of git tags whose creator date falls inside the observation window |
| Score formula | Bucketed score |

Buckets:

| Tags in Window | Score |
| ---: | ---: |
| 0 | 0 |
| 1 | 50 |
| 2-3 | 75 |
| >= 4 | 100 |

Notes:

- This uses git tags, not necessarily GitHub Releases.
- Repositories that release without tags may be undercounted.

### m5 Active Contributors

| Field | Value |
| --- | --- |
| Weight | 10% |
| Dimension | Contributors |
| Raw input column | `git_metrics.csv::contributors` |
| Score output column | `scores.csv::m5_contributors` |
| Data source | Local git history via `git_metrics.py` |
| Raw definition | Number of resolved human commit identities with at least one non-merge commit in the window |
| Score formula | `0 if n <= 1 else 100 * min(1, log2(n) / 5)` |
| Cap | 32 contributors gives score 100 |

Notes:

- Identity resolution is local to each repository or product set.
- Bots are excluded.

### m6 Effective Contributors

| Field | Value |
| --- | --- |
| Weight | 10% |
| Dimension | Contributors |
| Raw input column | `git_metrics.csv::eff_contributors` |
| Score output column | `scores.csv::m6_effective_contributors` |
| Data source | Local git history via `git_metrics.py` |
| Raw definition | Effective contributor count using inverse concentration of commit shares |
| Score formula | `0 if E <= 1 else 100 * min(1, log2(E) / 3)` |
| Cap | 8 effective contributors gives score 100 |

Raw formula:

```text
E = 1 / sum(contributor_commit_share^2)
```

Notes:

- This rewards distributed contribution rather than only the number of contributors.
- A project with many contributors but one dominant committer can still score modestly.

### m7 Contributor Retention

| Field | Value |
| --- | --- |
| Weight | 5% |
| Dimension | Contributors |
| Raw input column | `git_metrics.csv::retained` |
| Score output column | `scores.csv::m7_retention` |
| Data source | Local git history via `git_metrics.py` |
| Raw definition | Count of contributors active in both halves of the observation window |
| Score formula | Bucketed score |

Buckets:

| Retained Contributors | Score |
| ---: | ---: |
| 0 | 0 |
| 1 | 25 |
| 2 | 50 |
| 3-4 | 75 |
| >= 5 | 100 |

### m8 Issue Response Rate

| Field | Value |
| --- | --- |
| Weight | 10% |
| Dimension | Responsiveness |
| Raw input column | `gh_metrics.csv::m8_issue_response_rate` or `manual_metrics.csv::m8_issue_response_rate` |
| Score output column | `scores.csv::m8_issue_response` |
| Data source | GitHub GraphQL API via `gh_extract.py`; manual file for unavailable external trackers |
| Raw definition | Share of eligible issues receiving a qualifying response within 14 days |
| Score formula | `rate * 100` |
| Minimum data rule | At least 10 eligible issues required; otherwise metric is unavailable |

Eligibility:

- Issue author is not a bot.
- Issue was created within the observation window but early enough to observe the full 14-day response period.
- The code first uses non-maintainer issues. If fewer than 10 non-maintainer issues exist, it falls back to all issues and records this in `m8_basis`.

Qualifying response events:

- Non-author, non-bot issue comment.
- Non-author, non-bot label event.
- Non-author, non-bot assignment event.
- Non-author, non-bot close event.

Related diagnostic columns:

- `m8_n`
- `m8_responded`
- `m8_basis`
- `diag_issue_ttfr_median_h`
- `diag_issue_ttfr_n`
- `diag_issues_created`
- `diag_issues_by_maintainers`

### m9 PR Response Rate

| Field | Value |
| --- | --- |
| Weight | 10% |
| Dimension | Responsiveness |
| Raw input column | `gh_metrics.csv::m9_pr_response_rate` |
| Score output column | `scores.csv::m9_pr_response` |
| Data source | GitHub GraphQL API via `gh_extract.py` |
| Raw definition | Share of eligible non-bot PRs receiving a qualifying response within 7 days |
| Score formula | `rate * 100` |
| Minimum data rule | At least 10 eligible PRs required; otherwise metric is unavailable |

Eligibility:

- PR author is not a bot.
- PR was created within the observation window but early enough to observe the full 7-day response period.

Qualifying response events:

- Non-author, non-bot PR comment.
- Non-author, non-bot review.
- Merge by someone other than the author.
- Close by someone other than the author.

Related diagnostic columns:

- `m9_n`
- `m9_responded`
- `diag_pr_ttfr_median_h`
- `diag_pr_ttfr_n`

### m10 Issue Closure Within 90 Days

| Field | Value |
| --- | --- |
| Weight | 7.5% |
| Dimension | Responsiveness |
| Raw input column | `gh_metrics.csv::m10_issue_close90_rate` or `manual_metrics.csv::m10_issue_close90_rate` |
| Score output column | `scores.csv::m10_issue_close90` |
| Data source | GitHub GraphQL API via `gh_extract.py`; manual file for unavailable external trackers |
| Raw definition | Share of eligible human-authored issues closed within 90 days |
| Score formula | `rate * 100` |
| Minimum data rule | At least 10 eligible issues required; otherwise metric is unavailable |

Eligibility:

- Issue author is not a bot.
- Issue was created within the observation window but early enough to observe the full 90-day closure period.

Related diagnostic columns:

- `m10_n`
- `m10_closed`
- `diag_open_issues_over_180d_at_cutoff`

### m11 PR Decision Within 30 Days

| Field | Value |
| --- | --- |
| Weight | 7.5% |
| Dimension | Responsiveness |
| Raw input column | `gh_metrics.csv::m11_pr_decision30_rate` |
| Score output column | `scores.csv::m11_pr_decision30` |
| Data source | GitHub GraphQL API via `gh_extract.py` |
| Raw definition | Share of eligible non-bot PRs merged or closed within 30 days |
| Score formula | `rate * 100` |
| Minimum data rule | At least 10 eligible PRs required; otherwise metric is unavailable |

Eligibility:

- PR author is not a bot.
- PR was created within the observation window but early enough to observe the full 30-day decision period.

Related diagnostic columns:

- `m11_n`
- `m11_decided`
- `diag_prs_created`
- `diag_prs_merged`
- `diag_prs_closed_unmerged`
- `diag_prs_open_at_cutoff`
- `diag_prs_self_merged`
- `diag_acceptance_ratio`

### m12 Stars Gained

| Field | Value |
| --- | --- |
| Weight | 2.5% |
| Dimension | Adoption |
| Raw input column | `gh_metrics.csv::m12_stars_gained` |
| Score output column | `scores.csv::m12_stars` |
| Data source | GitHub REST star-history endpoint via `gh_extract.py` |
| Raw cache | `raw/<owner>__<repo>/stars_history.json` |
| Raw definition | Count of stars gained inside the observation window |
| Score formula | `100 * min(1, log10(1 + s) / 3)` |
| Cap | 999 stars approximately reaches score 100 |

Notes:

- This uses GitHub's privacy-safe star-history endpoint, not current star totals from the live audit workbook.
- For multi-repo product sets, stars are summed across repositories.
- Non-GitHub repositories do not receive this metric unless filled manually in a compatible future process.

### m13 Forks Created

| Field | Value |
| --- | --- |
| Weight | 2.5% |
| Dimension | Adoption |
| Raw input column | `gh_metrics.csv::m13_forks_created` |
| Score output column | `scores.csv::m13_forks` |
| Data source | GitHub GraphQL API via `gh_extract.py` |
| Raw cache | `raw/<owner>__<repo>/forks.json` |
| Raw definition | Count of forks created inside the observation window |
| Score formula | `100 * min(1, log10(1 + f) / 2)` |
| Cap | 99 forks approximately reaches score 100 |

Notes:

- Forks are weighted lightly because they are an adoption/popularity signal rather than direct maintenance evidence.
- For multi-repo product sets, forks are summed across repositories.

## 6. Final Score Calculation

For each available metric:

```text
weighted contribution = metric_score * metric_weight / 100
```

Then:

```text
weighted_points = sum(weighted contribution for available metrics)
available_weight_pct = sum(weights for available metrics)
score_on_available = weighted_points / available_weight_pct * 100
component4_score = score_on_available, only if missing-data rule is satisfied
```

If the missing-data rule is not satisfied:

```text
component4_score = blank
score_status = insufficient data
```

## 7. Repository Resolution and Provenance

Repository provenance is stored primarily in `repos.csv` and carried into the workbook.

Important columns:

| Column | Meaning |
| --- | --- |
| `oisi_id` | Stable Component 4 row identifier |
| `infrastructure` | Open infrastructure name |
| `original_url` | URL from the original/main sheet |
| `decision` | Current scoring decision |
| `host` | Platform used by the extractor: GitHub, GitLab, Gitea, Bitbucket, etc. |
| `repo` | Repository or semicolon-separated product set used for scoring |
| `branch` | Optional branch override |
| `upstream` | Optional upstream repo for fork-deduplication |
| `scored` | `yes` when included in extraction/scoring; `no` when excluded |
| `classification` | `software`, `standard`, `not-assessable`, `duplicate`, or `needs-decision` |
| `issue_tracker` | Human-readable issue tracker decision |
| `issue_repo` | GitHub repo used for issue metrics when different from code repo |
| `notes` | Audit notes and rationale |

Resolved URL marker columns:

| Column | Meaning |
| --- | --- |
| `resolved_url_marker` | `yes` when the row appeared in the user-provided resolved URL list |
| `resolved_url_candidate_repo` | Candidate canonical repository from that list |
| `resolved_url_platform` | Candidate platform from that list |
| `resolved_url_action` | `applied`, `already_current`, or `marker_only` |
| `resolved_url_note` | Why the candidate was applied or only recorded |

Recovery report:

- `resolved_url_recovery_report.csv` compares before/after status for the 22 resolved URL rows.
- It records whether git data became available, available weight before/after, and remaining gaps.

## 8. Raw Data Cache Provenance

GitHub API data is cached under `raw/<owner>__<repo>/`:

| File | Used For |
| --- | --- |
| `prs.json` | PR response rate, PR decision rate, PR diagnostics |
| `issues.json` | Issue response rate, issue closure rate, issue diagnostics |
| `forks.json` | Forks created in the observation window |
| `stars_history.json` | Stars gained in the observation window |

Git clone data is stored under `clones/` as bare repositories. `git_metrics.py` reads these clones and writes the summarized metrics to `git_metrics.csv`.

## 9. Diagnostic-Only Fields

The following are useful for interpretation but are not directly weighted in the final score:

- `active_months`
- `bf50`
- `bf80`
- `top_share`
- `bot_commits`
- `bot_share`
- `merge_commits`
- `merged_identity_groups`
- `diag_issue_ttfr_median_h`
- `diag_issue_ttfr_n`
- `diag_pr_ttfr_median_h`
- `diag_pr_ttfr_n`
- `diag_acceptance_ratio`
- `diag_open_issues_over_180d_at_cutoff`
- `diag_issues_by_maintainers`
- `diag_prs_self_merged`

These fields are included to support audit review, but they do not contribute to `component4_score`.

## 10. Known Limitations

- GitHub issue/PR metrics are only directly available for GitHub repositories or GitHub issue tracker mappings in `issue_repo`.
- GitLab, Gitea, Jira, Bitbucket, and no-public-tracker projects need manual metrics or platform-specific extractors for m8, m10, and sometimes m9/m11.
- Git tags may not capture all release practices.
- Stars and forks are adoption signals and intentionally low-weighted.
- The 10-item minimum for responsiveness metrics prevents tiny samples from producing misleading scores.
- Some standards/documentation repositories can be measured technically, but they may not be comparable to software codebases; the `classification` and `decision` columns preserve that context.

