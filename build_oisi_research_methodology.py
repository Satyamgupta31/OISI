from __future__ import annotations

from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
from xml.sax.saxutils import escape


OUT = Path("OISI_Research_Methodology.docx")


def p(text: str = "", style: str = "Normal") -> str:
    return f'<w:p><w:pPr><w:pStyle w:val="{style}"/></w:pPr><w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>'


def tbl(headers: list[str], rows: list[list[str]], widths: list[int] | None = None) -> str:
    widths = widths or [9000 // len(headers)] * len(headers)
    out = ['<w:tbl><w:tblPr><w:tblStyle w:val="GridTable"/><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/><w:tblBorders><w:top w:val="single" w:sz="5" w:color="B7C9D6"/><w:left w:val="single" w:sz="5" w:color="B7C9D6"/><w:bottom w:val="single" w:sz="5" w:color="B7C9D6"/><w:right w:val="single" w:sz="5" w:color="B7C9D6"/><w:insideH w:val="single" w:sz="4" w:color="D7E0E7"/><w:insideV w:val="single" w:sz="4" w:color="D7E0E7"/></w:tblBorders></w:tblPr>']
    for ri, row in enumerate([headers] + rows):
        out.append('<w:tr>')
        for ci, value in enumerate(row):
            head = ri == 0
            fill = '<w:shd w:fill="1F4E78"/>' if head else ''
            color = 'FFFFFF' if head else '243746'
            bold = '<w:b/>' if head else ''
            out.append(f'<w:tc><w:tcPr><w:tcW w:w="{widths[ci]}" w:type="dxa"/>{fill}<w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:spacing w:after="35" w:before="25"/></w:pPr><w:r><w:rPr>{bold}<w:color w:val="{color}"/><w:sz w:val="16"/></w:rPr><w:t xml:space="preserve">{escape(str(value))}</w:t></w:r></w:p></w:tc>')
        out.append('</w:tr>')
    out.append('</w:tbl>')
    return ''.join(out)


def section(title: str, subtitle: str, blocks: list[str], last: bool = False) -> str:
    body = p(title, 'Heading1') + p(subtitle, 'Subtitle') + ''.join(blocks)
    return body if last else body + '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'


def build() -> str:
    pages: list[str] = []
    pages.append(section('OISI Research Methodology', 'Data sources, metric development, GitHub extraction, scoring, and equal-metric analysis | Snapshot: 14 October 2026', [
        p('Abstract', 'Heading2'),
        p('This study develops and applies the Open Infrastructure Sustainability Index (OISI) to a master set of 112 open infrastructures. Three components—Community Engagement, Governance Practices, and Organizational Capacity—were operationalized from the infrastructure profile data associated with the Zenodo record cited below. A fourth component, Development Activity, was assembled from repository histories and GitHub API data. Evidence was mapped to defined indicators, normalized to common 0-100 scales, and aggregated using transparent coverage rules. Every scored, applicable metric receives an equal share within its component, and each of the four component scores receives an equal 25% share in the overall index.'),
        p('Research aim and unit of analysis', 'Heading2'),
        p('The unit of analysis is an infrastructure/service entry in the curated 112-entry master set. The study asks how consistently documented community, governance, organizational, and recent development signals can be summarized for comparative description. The index is not a direct estimate of future survival, software quality, social impact, or user satisfaction. Repository activity is only one observable part of a broader socio-technical system; this limitation is recognized in open-source sustainability and community-health research [3, 4].'),
        p('Design principles', 'Heading2'),
        tbl(['Principle', 'Application'], [
            ['Traceability', 'Preserve source fields, repository mapping decisions, derived values, statuses, and missingness notes.'],
            ['Construct coverage', 'Use four complementary components rather than treating repository activity as the whole of sustainability.'],
            ['Comparable scales', 'Convert component indicators to bounded 0-1 or 0-100 values before aggregation.'],
            ['Coverage safeguards', 'Do not publish a component score unless its minimum data and dimension-coverage conditions are met.'],
            ['Equal influence', 'Give every scored, applicable metric the same share within its component and every component the same 25% share in the overall index.'],
        ], [2300, 6700]),
        p('Terminology', 'Heading2'),
        p('“Metric” means a scored indicator in the component framework. “Raw field” means the source data item or API observation used to calculate it. “Score anchor” is the rule that maps raw evidence to a normalized score. “Missing” means that an observation was unavailable or unknown; it is not automatically a measured zero. “Not applicable” is a structural condition and is removed from the relevant denominator when specified by the component rules.'),
    ]))

    pages.append(section('1. Dataset and Components 1-3', 'Profile-data extraction, field selection, and harmonization', [
        p('Source and scope', 'Heading2'),
        p('The profile-data source used for the first three components is the Zenodo record supplied for this study: “Zenodo record 18225266” [1]. The component workbooks preserve the extracted profile fields and subsequent scoring fields. The record page was rate-limited during this document preparation, so its title, version, creators, and license should be copied directly from the live record metadata into the final manuscript citation before submission; no metadata not visible in the project files is inferred here.'),
        p('Three profile-based components', 'Heading2'),
        tbl(['Component', 'Raw fields selected from the profile dataset', 'Purpose'], [
            ['Community Engagement', '001 service name; 008 provider; 010 website; 013 location; 014 solution category; 019 repository URL; 037 engagement status; 038 activities; 938 other activities; 116 engagement URL; 040 contribution opportunities; 940 other opportunities; 041 contribution guidance; 042-043 Code of Conduct status and URL. Fields 039, 078, and 107 are retained for flags/context, not as scored indicators.', 'Describe engagement routes, contribution opportunities, public guidance, and community norms.'],
            ['Governance Practices', '024 governance structure; 026 governance records; 108 bylaws; 050 board structure; 048 community governance; 029 accessibility; 033 transparent pricing; 035 equity; 044 privacy; 046 values frameworks. Field 100 board level is descriptive and excluded from the active score because its coverage is sparse. Evidence URL columns are retained for traceability.', 'Describe formal governance, participation, and public transparency commitments.'],
            ['Organizational Capacity', '079 technology readiness; 080 maintenance status; 056 launch date; 051 business form; 058 affiliations; 099 SCOSS; 076 SaaS options; 077 registered service providers. Fields 061 funding and 054-055 staffing are retained as contextual/descriptive variables, not current scored metrics.', 'Describe maturity and maintenance, institutional support, and operational/service capacity.'],
        ], [1650, 5000, 2350]),
        p('Extraction and transformation workflow', 'Heading2'),
        p('The profile data were carried into component workbooks and matched to the master infrastructure entries using service identity and provider/context fields. Raw answer values were retained alongside normalized metric values. Multi-select fields were parsed into distinct categories; “other” text fields were used to recover categories not represented by fixed answer options. Evidence URLs were retained so a reviewer can inspect the underlying public claim. Status answers were mapped to normalized scores using the component scoring rules; counts were capped at defined ceilings where an unbounded count could otherwise dominate. The exact mapping for each active metric is itemized in the accompanying component scoring workbooks and the metric summary below.'),
        p('Data cleaning and identity matching', 'Heading2'),
        p('The 112-entry master list is the common analysis frame. Names were normalized for matching across source tables; provider, website, repository, and solution-category fields were used to review ambiguous matches. Duplicate, standard/non-code, not-assessable, or unresolved entries were not silently converted to software-project observations. Source status and coverage were carried into the combined index. The final outputs therefore distinguish a scored zero from a blank, unknown, structurally not-applicable, and excluded entry.'),
        p('Measurement roles', 'Heading2'),
        p('Community Engagement and Governance are based primarily on documented profile claims and URLs; Organizational Capacity combines profile indicators and stated organizational/service facts. These sources capture different forms of evidence and may contain self-reporting or incomplete disclosure. They are interpreted as observable evidence in the dataset, not independently audited facts about internal practice.'),
    ]))

    pages.append(section('2. Development Activity Data', 'Repository resolution, Git history, GitHub APIs, and provenance', [
        p('Repository identification and inclusion', 'Heading2'),
        p('A repository-resolution table (`repos.csv`) links each master-list entry to its canonical repository or product repository set. It records the source URL, decision, host, repository path(s), optional branch, upstream fork, issue-tracker mapping, scoring status, classification, and notes. A product represented by multiple repositories uses semicolon-separated repository paths; duplicate commit hashes across that set are deduplicated. When a repository is a fork and upstream information is available, inherited upstream commits can be excluded. Entries classified as standards, duplicates, not assessable, or requiring a decision are retained in the master set with their disposition rather than forced into the code-activity sample.'),
        p('Git history extraction', 'Heading2'),
        p('Repository histories were collected for the resolved repository or repository set associated with each infrastructure. The selected/default branch is used unless a branch override is recorded. Commit authorship uses author date. Merge commits are excluded from authorship counts and retained separately as diagnostics. A documented bot list, GitHub account type, and name/email patterns exclude automated accounts. Contributor identities are resolved within a repository using repository `.mailmap`, case-insensitive identical email, GitHub no-reply login, GitHub login when available, normalized full-name match (flagged for audit), and optional manual identity map. Identities are not merged across unrelated repositories. These histories support metrics 1-7; GitHub API data support metrics 8-13.'),
        p('GitHub API extraction using R', 'Heading2'),
        p('For Development Activity metrics 8-13, R code queried GitHub’s GraphQL API for issue, pull-request, and fork records and GitHub’s REST starring/history endpoint for star events. The R workflow used the resolved repository and issue-tracker mappings, requested the fields needed to calculate event dates, authors, comments/reviews, state changes, and repository identifiers, and followed API pagination until all eligible records for the study period were retrieved. Raw responses were retained as per-repository JSON data; R then filtered, classified, and summarized those records into metric-level observations and diagnostics. Star gains and forks created were counted within the defined period rather than inferred from current repository totals. GitHub documents cursor-based GraphQL pagination and REST pagination [6, 7].'),
        p('Requests were authenticated with a GitHub token limited to public-repository read access. Rate-limit/API errors, inaccessible repositories, and tracker mismatches were treated as retrieval or coverage issues and recorded for review; they were not coded as zero activity. The repository and issue-tracker mapping was maintained separately so that each API result could be traced back to the corresponding infrastructure and canonical repository.'),
        tbl(['Data family', 'Observed source', 'Use'], [
            ['Commit and tag history', 'Resolved repository histories and derived commit/tag tables', 'Human commits, active weeks, recency, releases/tags, contributors, effective contributors, retention.'],
            ['Issues and comments/events', 'GitHub GraphQL; `raw/<owner>__<repo>/issues.json`', 'Issue response and closure metrics plus diagnostics.'],
            ['Pull requests and reviews/events', 'GitHub GraphQL; `raw/<owner>__<repo>/prs.json`', 'PR response and decision-time metrics plus diagnostics.'],
            ['Fork events', 'GitHub GraphQL; `forks.json`', 'Forks created during the observation window.'],
            ['Stars over time', 'GitHub REST star history; `stars_history.json`', 'Stars gained during the observation window, not current total stars.'],
            ['Non-GitHub tracker evidence', 'Optional `manual_metrics.csv`', 'Manual values or documented missingness for selected Jira/GitLab/other tracker cases.'],
        ], [2200, 3500, 3300]),
    ]))

    pages.append(section('3. Observation Window and Metric Selection', 'How the 12-month period and final indicator set were operationalized', [
        p('Fixed 12-month window', 'Heading2'),
        tbl(['Boundary', 'Definition'], [
            ['Start', '1 October 2025, 00:00:00 UTC, inclusive'],
            ['End', '1 October 2026, 00:00:00 UTC, exclusive'],
            ['Duration', 'One full year; common to all time-windowed Development Activity measures'],
        ], [2200, 6800]),
        p('Rationale and follow-up', 'Heading2'),
        p('A single 12-month window provides a common recent observation period across infrastructures, includes a full annual cycle, and gives enough time for repeated activity and contributor patterns to be observed. It is a design choice, not an empirically proven optimal horizon. The fixed UTC boundaries avoid time-zone ambiguity. For response/closure measures, eligible items must have enough follow-up before the end boundary: 14 days for issue response, 7 days for PR response, 90 days for issue closure, and 30 days for PR decision. This avoids treating recent, right-censored records as failures. The fixed study window is 1 October 2025 through 30 September 2026 UTC.'),
        p('Metric development and finalization', 'Heading2'),
        p('Candidate indicators were screened against four practical questions: (1) construct relevance to continuity, contributor capacity, responsiveness, or adoption; (2) a defensible, explicit operational definition; (3) feasibility of obtaining the same evidence for a meaningful portion of the cohort; and (4) interpretability and non-duplication. Candidate metrics were then tested against the 112-entry repository/tracker map, extraction availability, distribution, minimum sample sizes, and sensitivity behavior. Metrics with insufficient or incompatible evidence were retained as unavailable or diagnostic rather than imputed as poor performance. Several measures were kept as diagnostics to help interpret concentration and extraction quality without affecting the index.'),
        p('Empirical checks for redundancy and feasibility', 'Heading2'),
        p('The saved pipeline notes high association between commits and active weeks (Spearman rho approximately 0.87 among active repositories) and between contributor count and effective contributors (approximately 0.88). These pairs are conceptually related but not identical: activity volume differs from temporal spread, and contributor headcount differs from contribution concentration. Their potential redundancy is reported as a limitation and was considered during metric review. Responsiveness measures require at least 10 eligible issues or pull requests; otherwise that metric is unavailable. The final Development Activity eligibility rule further requires broad metric coverage and at least half of each dimension.'),
        p('Metric families selected', 'Heading2'),
        tbl(['Dimension', 'Metrics', 'Construct represented'], [
            ['Continuity', 'm1 commits; m2 active weeks; m3 recency; m4 releases', 'Recent human development volume, cadence, recency, and release activity.'],
            ['Contributors', 'm5 contributor count; m6 effective contributors; m7 retention', 'Breadth, distribution, and repeated participation of contributors.'],
            ['Responsiveness', 'm8 issue response; m9 PR response; m10 issue closure; m11 PR decision', 'Timeliness and disposition of incoming work, conditional on sufficient volume.'],
            ['Adoption', 'm12 stars gained; m13 forks created', 'Platform-visible attention/reuse proxies; indirect signals, not quality measures.'],
        ], [2100, 3650, 3250]),
    ]))

    pages.append(section('4. Metric Definitions and Scoring', 'Development Activity: raw observation to normalized metric value', [
        p('All Development Activity metric scores are normalized to 0-100. Bots are excluded, and the window and identity rules are applied consistently across the repository cohort. Each of the 13 available, applicable metric scores contributes 1/13 (7.69%) to the component mean.'),
        tbl(['Metric', 'Definition, normalized score, and equal share'], [
            ['m1 Human commits', 'Non-merge, non-bot commits in window; 100*min(1, log10(1+c)/log10(1001)); cap 1,000. Share: 1/13 (7.69%).'],
            ['m2 Active weeks', 'Distinct weeks with at least one qualifying human commit / 52 * 100. Share: 1/13 (7.69%).'],
            ['m3 Recency', 'Days since latest qualifying commit at cutoff; 100 through 30 days, then linearly decreases to 0 at 365 days. Share: 1/13 (7.69%).'],
            ['m4 Release activity', 'Git tags in window: 0=0; 1=50; 2-3=75; 4 or more=100. Share: 1/13 (7.69%).'],
            ['m5 Active contributors', 'Distinct resolved human contributors; 0 when n<=1, otherwise 100*min(1, log2(n)/5); cap 32. Share: 1/13 (7.69%).'],
            ['m6 Effective contributors', 'E=1/sum(p_i^2), where p_i is each contributor share of qualifying commits; 0 when E<=1; 100*min(1, log2(E)/3); cap E=8. Share: 1/13 (7.69%).'],
            ['m7 Contributor retention', 'Contributors active in both half-windows: 0=0; 1=25; 2=50; 3-4=75; 5+=100. Share: 1/13 (7.69%).'],
            ['m8 Issue response', 'Eligible human-authored issues receiving a qualifying non-author, non-bot response within 14 days / eligible issues * 100. Minimum 10. Share: 1/13 (7.69%).'],
            ['m9 PR response', 'Eligible non-bot PRs with qualifying non-author response within 7 days / eligible PRs * 100. Minimum 10. Share: 1/13 (7.69%).'],
            ['m10 Issue close <=90d', 'Eligible human-authored issues closed within 90 days / eligible issues * 100. Minimum 10. Share: 1/13 (7.69%).'],
            ['m11 PR decision <=30d', 'Eligible non-bot PRs merged or closed within 30 days / eligible PRs * 100. Minimum 10. Share: 1/13 (7.69%).'],
            ['m12 Stars gained', 'Stars gained during window; 100*min(1, log10(1+s)/3); cap 1,000. Share: 1/13 (7.69%).'],
            ['m13 Forks created', 'Forks created during window; 100*min(1, log10(1+f)/2); cap 100. Share: 1/13 (7.69%).'],
        ], [1800, 6250, 950]),
        p('Qualifying response events', 'Heading2'),
        p('For issues, qualifying response events include a non-author, non-bot comment, label, assignment, or close action. For pull requests, qualifying events include a non-author, non-bot comment/review, merge by someone other than the author, or close by someone other than the author. The issue-response extractor prefers non-maintainer-origin issues where the sample is adequate and records its basis; if fewer than 10 are available, it may fall back to all eligible issues as documented by the extractor. These definitions are operational proxies for response, not judgments about response quality.'),
        p('Zero versus unavailable', 'Heading2'),
        p('For a Git-complete dormant repository, no qualifying activity can be a measured zero for designated Git metrics (including recency/effective contributor cases handled by the scoring code). In contrast, an API failure, an unsupported tracker, an unresolved repository, or an insufficient 10-item sample is unavailable. The scoring workflow keeps these states separate. Where an issue/PR tracker is outside GitHub, manual observations may be entered through the documented override file; otherwise the associated measures remain unavailable.'),
    ]))

    pages.append(section('5. Profile Metric Sets and Aggregation', 'Community, governance, and organizational indicators', [
        p('Community Engagement: seven active metrics', 'Heading2'),
        tbl(['Metric', 'Profile fields', 'Equal share', 'Operational interpretation'], [
            ['E1 Engagement availability', '037', '1/7 (14.29%)', 'Status/availability of a community engagement mechanism.'],
            ['E2 Engagement activity breadth', '038 + 938', '1/7 (14.29%)', 'Distinct engagement activity categories; score is capped at 8 categories.'],
            ['E3 Engagement evidence strength', '116 + 938', '1/7 (14.29%)', 'Presence/strength of supporting engagement evidence under current URL/description rule.'],
            ['C1 Contribution opportunity breadth', '040 + 940', '1/7 (14.29%)', 'Distinct contribution opportunity types; capped at 6 categories.'],
            ['C2 Contribution guidance', '041', '1/7 (14.29%)', 'Evidence of contribution guidelines or relevant participation guidance.'],
            ['C3 Open contribution infrastructure', '019', '1/7 (14.29%)', 'Public code repository/forge evidence under current URL scoring rule.'],
            ['N1 Community norms / Code of Conduct', '042 + 043', '1/7 (14.29%)', 'Status and URL evidence; implemented without URL receives the existing reduced score.'],
        ], [1950, 1200, 1500, 4300]),
        p('The Community Engagement score is the arithmetic mean of scored, applicable indicators. A score requires at least 5 of 7 metrics (60%) and at least one observed metric in each of Engagement, Contribution, and Norms. Field 039 is not scored and is used for consistency flags. Code-of-conduct evidence remains in this component only.'),
        p('Governance Practices: ten active indicators', 'Heading2'),
        p('The ten scored indicators are 024 governance structure, 026 governance records, 108 bylaws, 050 board structure, 048 community governance, 029 accessibility, 033 transparent pricing, 035 equity commitment, 044 privacy commitment, and 046 values frameworks. Each observed indicator contributes 10% of the component score. Status indicators use the workbook’s common status mapping (implemented/available 1.00; in progress 0.67; considering 0.33; not planning 0.00). Fields 048/050/046 have presence/blank coding as defined in the source workbook; blank-as-zero is an explicit coding choice, not verified absence. Field 100 board-level is descriptive and excluded due to sparse coverage. A score requires at least 7 of 10 indicators and representation of Formal and Transparency.'),
        p('Organizational Capacity: eight candidate indicators', 'Heading2'),
        p('M1 technology readiness (079), M2 maintenance status (080), M3 infrastructure age (056), I1 formal organizational host (051), I2 affiliations/institutional support (058), I3 SCOSS participation (099), O1 hosted/SaaS availability (076), and O2 registered service-provider ecosystem (077). Maturity and Maintenance, Institutional Support, and Operational Capacity form the three dimensions. Each scored, applicable indicator receives 1/k of the component score, where k is the number of applicable indicators with a score; when all eight apply, each share is 12.5%. O1 is structurally excluded when SaaS is Not applicable; O2 applies only to specified third-party delivery arrangements. Unknown/blank values are missing unless the source scoring workbook explicitly maps them. A score requires at least 70% of applicable metrics and at least two dimensions represented.'),
        p('Overall OISI aggregation', 'Heading2'),
        p('Each component score receives an equal 25% share: OISI = 0.25*Community Engagement + 0.25*Governance Practices + 0.25*Organizational Capacity + 0.25*Development Activity. A strict OISI is reported when all four components meet their eligibility requirements. For the secondary all-112 evidence-adjusted view, an unavailable component contributes zero; this convention must not be interpreted as observed zero performance. Component coverage is shown with the score.'),
    ]))

    pages.append(section('6. Equal-Weight Calculation', 'Equal shares within each component and across the four components', [
        p('Purpose', 'Heading2'),
        p('Every metric score is first normalized using its documented anchor. Within a component, the available, applicable scores are then arithmetically averaged. Each component score contributes one quarter of the total OISI. The metrics are not pooled across components: component scoring and coverage checks are completed first, followed by the four-component mean.'),
        p('Formula and shares', 'Heading2'),
        p('For component j: C_j = sum of normalized scores for scored/applicable metrics divided by the number of scored/applicable metrics. In a fully observed component, each metric share is 1/k, where k is the number of active metrics.'),
        tbl(['Component', 'Metrics', 'Share per metric', 'Implied domain shares'], [
            ['Community Engagement', '7', '1/7 = 14.29%', 'Engagement 42.86%; Contribution 42.86%; Norms 14.29%.'],
            ['Governance Practices', '10', '1/10 = 10%', 'Formal 40%; Participatory 10%; Transparency 50%. Field 100 excluded.'],
            ['Organizational Capacity', '8 candidates', '1/8 = 12.5% when all apply', 'Maturity 37.5%; Institutional 37.5%; Operational 25%. N/A removed.'],
            ['Development Activity', '13', '1/13 = 7.69%', 'Continuity 30.77%; Contributors 23.08%; Responsiveness 30.77%; Adoption 15.38%.'],
        ], [1800, 950, 1900, 4300]),
        p('Equal-metric eligibility', 'Heading2'),
        tbl(['Component', 'Minimum coverage rule'], [
            ['Community Engagement', 'At least 5/7 metrics (60%) and at least one observed metric in each of its three groups.'],
            ['Governance Practices', 'At least 7/10 metrics (70%) and at least one Formal and one Transparency indicator.'],
            ['Organizational Capacity', 'At least 70% of applicable metrics and at least two of its three dimensions represented.'],
            ['Development Activity', 'At least 10/13 metrics (75%) and at least half of the metric count in every dimension: 2/4, 2/3, 2/4, and 1/2.'],
        ], [2250, 6750]),
        p('Why this is a useful “basic overview”', 'Heading2'),
        p('The rule is directly explainable: each observed, applicable indicator contributes one equal portion of its component score, and each component contributes one equal quarter of the OISI. This does not mean equal weights are empirically proven to be uniquely correct. Indicators can differ in reliability, redundancy, and conceptual breadth; the metric set, anchors, missingness rules, and eligibility thresholds remain methodological choices. Composite-indicator guidance recommends documenting aggregation assumptions and examining robustness [2].'),
        p('Equal metric shares within components are 14.29% for each of 7 Community Engagement metrics, 10% for each of 10 Governance indicators, 12.5% for each of 8 Organizational Capacity metrics when all apply, and 7.69% for each of 13 Development Activity metrics. Each component is 25% of the overall index.'),
    ]))

    pages.append(section('7. Evaluation, Results, and Limitations', 'Validation checks and interpretation of the current 112-entry analysis', [
        p('Evaluation and verification steps', 'Heading2'),
        tbl(['Step', 'Check performed / reproducible artifact'], [
            ['1. Cohort reconciliation', 'Confirm 112 master entries; inspect repository resolution decisions, duplicate handling, platform, tracker, and classification in repos.csv and the joined workbook.'],
            ['2. Source-to-metric audit', 'Trace each profile metric to its raw field, each repository metric to Git/API/manual inputs, and preserve provenance/cache records.'],
            ['3. Scoring audit', 'Recalculate normalization, status mapping, equal metric means, applicability, eligibility, and component scores from source workbooks and scoring code.'],
            ['4. API completeness', 'Inspect pagination, raw caches, extraction diagnostics, missing tracker states, sample counts, and follow-up cutoffs.'],
            ['5. Distribution/redundancy', 'Review metric ranges and missingness; compare related metrics (e.g., commits vs active weeks; contributor count vs effective contributors).'],
            ['6. Coverage review', 'Report strict complete cases separately from the all-112 evidence-adjusted view; inspect sensitivity to missing-data assumptions and eligibility thresholds.'],
        ], [2050, 6950]),
        p('Current equal-metric results', 'Heading2'),
        tbl(['Result', 'Value in current index workbook'], [
            ['Equal-metric component scores available', 'Community Engagement 88; Governance Practices 108; Organizational Capacity 109; Development Activity 72.'],
            ['All four equal-metric components eligible', '57 entries; mean 74.54; range 46.90-91.52.'],
            ['All 112, unavailable component scores contribute zero', 'Mean 61.46; range 18.91-91.52. Secondary evidence-adjusted view, not observed performance for missing components.'],
        ], [3100, 5900]),
        p('Interpretation', 'Heading2'),
        p('The complete-case result is the primary equal-weight index because all four components meet their scoring rules. The all-112 score is a secondary evidence-adjusted view: it includes zero contributions for unavailable component scores and therefore reflects both measured scores and coverage. Report both with the component coverage, and do not treat a zero contribution for missing evidence as a measured zero outcome.'),
        p('Limitations', 'Heading2'),
        p('Profile evidence may be self-reported or incompletely documented. Repository selection and product-to-repository mapping involve judgment. GitHub API activity does not represent work on every hosting/tracker platform; platform-specific behavior can affect issue and PR metrics. A ten-item threshold leaves low-volume projects without responsiveness scores. Tags do not capture every release practice. Stars and forks are platform-dependent, indirect signals. Identity resolution can merge or split contributors imperfectly. Metrics may be correlated. The 12-month window is a fixed snapshot, not a long-term trend. Missing-as-zero combined scores should never be interpreted as measured zero outcomes. Equal weighting is transparent and interpretable, but it does not establish that every metric has equal reliability or conceptual importance.'),
        p('Reproducibility and reporting package', 'Heading2'),
        p('The core auditable materials include the profile component workbooks, repository-resolution table, R source code used for API retrieval and transformations, raw API JSON responses, derived metric tables, repository-history outputs, and `OISI_Component4_Final_Master_With_Provenance.xlsx`. The combined OISI workbook provides the Codebook, Summary, component scores, scoring rules, equal-metric scores, ranking and all-112 views, and source data. Record the exact dataset version/download date, API retrieval timestamp, R version, package versions, code revision, and token permission scope in the submitted replication package.'),
    ]))

    pages.append(section('References and Author Checks', 'Citation details and items to confirm before submission', [
        p('References', 'Heading2'),
        p('[1] Zenodo. Record 18225266, dataset/profile-data source used for the OISI profile components. https://zenodo.org/records/18225266. Verify and add the exact record title, creators, version, publication year, and DOI from the record metadata before final submission.'),
        p('[2] OECD, European Union, & European Commission Joint Research Centre. (2008). Handbook on Constructing Composite Indicators: Methodology and User Guide. OECD Publishing. https://doi.org/10.1787/9789264043466-en'),
        p('[3] Goggins, S., Lumbard, K., & Germonprez, M. (2021). Open Source Community Health: Analytical Metrics and Their Corresponding Narratives. IEEE/ACM SoHeal. https://doi.org/10.1109/SoHeal52568.2021.00010'),
        p('[4] Yin, L., Chakraborti, M., Yan, Y., Schweik, C., Frey, S., & Filkov, V. (2022). Open Source Software Sustainability: Combining Institutional Analysis and Socio-Technical Networks. Proceedings of the ACM on Human-Computer Interaction, 6(CSCW2), Article 404. https://doi.org/10.1145/3555129'),
        p('[5] CHAOSS Project. CHAOSS Metrics. https://chaoss.community/metrics/'),
        p('[6] GitHub Docs. Using pagination in the GraphQL API. https://docs.github.com/en/graphql/guides/using-pagination-in-the-graphql-api'),
        p('[7] GitHub Docs. REST API endpoints for starring. https://docs.github.com/en/rest/activity/starring'),
        p('Author checks before treating this as submission-final', 'Heading2'),
        p('1. Confirm the Zenodo record metadata and cite its DOI/version and the specific downloaded file(s). The record page was not accessible to the document builder because Zenodo returned a rate limit; the user-provided record URL is included, but bibliographic metadata must be verified.'),
        p('2. Archive the R source code that generated the API observations and derived metric table, together with its R and package versions, retrieval timestamp, API pagination/error logs, and manual overrides.'),
        p('3. Confirm the exact data download date and API collection completion timestamp. Keep these timestamps separate from the fixed observation-window endpoints.'),
        p('4. Confirm whether any extraction or metric-revision decisions after 1 October 2026 changed the reported workbook; if so, version the data freeze and regenerate the result values in this document.'),
        p('Closing note', 'Heading2'),
        p('Alongside the equal-weight study documented here, we are exploring a separate weighted analysis using the same infrastructure set for comparative analysis. That work is distinct from the present methodology and does not change the equal-weight results reported in this document.'),
    ], last=True))

    body = ''.join(pages)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:body>{body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
            '<w:pgMar w:top="620" w:right="650" w:bottom="620" w:left="650" w:header="300" w:footer="300"/>'
            '</w:sectPr></w:body></w:document>')


def main() -> None:
    with ZipFile(OUT, 'w', ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>''')
        z.writestr('_rels/.rels', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>''')
        z.writestr('word/document.xml', build())
        z.writestr('word/styles.xml', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/><w:sz w:val="18"/><w:color w:val="243746"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="70" w:line="215" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:after="80"/></w:pPr><w:rPr><w:rFonts w:ascii="Aptos Display" w:hAnsi="Aptos Display"/><w:b/><w:color w:val="1F4E78"/><w:sz w:val="27"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="60" w:after="35"/></w:pPr><w:rPr><w:b/><w:color w:val="1F4E78"/><w:sz w:val="20"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="120"/></w:pPr><w:rPr><w:color w:val="58717F"/><w:sz w:val="16"/></w:rPr></w:style><w:style w:type="table" w:styleId="GridTable"><w:name w:val="Grid Table"/></w:style></w:styles>''')
    print(OUT.resolve())


if __name__ == '__main__':
    main()
