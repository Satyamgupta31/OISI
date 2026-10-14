from __future__ import annotations

from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
from xml.sax.saxutils import escape


OUT = Path("OISI_Full_Methodology_and_Metrics_Document.docx")


def run(text: str, bold: bool = False, italic: bool = False) -> str:
    props = ""
    if bold or italic:
        props = "<w:rPr>" + ("<w:b/>" if bold else "") + ("<w:i/>" if italic else "") + "</w:rPr>"
    return f'<w:r>{props}<w:t xml:space="preserve">{escape(str(text))}</w:t></w:r>'


def para(text: str = "", style: str = "Normal", bold_prefix: str | None = None) -> str:
    ppr = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>'
    if bold_prefix and text.startswith(bold_prefix):
        content = run(bold_prefix, True) + run(text[len(bold_prefix):])
    else:
        content = run(text)
    return f"<w:p>{ppr}{content}</w:p>"


def table(headers: list[str], rows: list[list[str]], widths: list[int] | None = None) -> str:
    cols = len(headers)
    widths = widths or [9000 // cols] * cols
    out = [
        '<w:tbl><w:tblPr><w:tblStyle w:val="GridTable"/><w:tblW w:w="0" w:type="auto"/>'
        '<w:tblLayout w:type="fixed"/><w:tblBorders>'
        '<w:top w:val="single" w:sz="5" w:color="B7C9D6"/><w:left w:val="single" w:sz="5" w:color="B7C9D6"/>'
        '<w:bottom w:val="single" w:sz="5" w:color="B7C9D6"/><w:right w:val="single" w:sz="5" w:color="B7C9D6"/>'
        '<w:insideH w:val="single" w:sz="4" w:color="D7E0E7"/><w:insideV w:val="single" w:sz="4" w:color="D7E0E7"/>'
        '</w:tblBorders></w:tblPr>'
    ]
    for ridx, row in enumerate([headers] + rows):
        out.append("<w:tr>")
        for cidx, val in enumerate(row):
            fill = '<w:shd w:fill="1F4E78"/>' if ridx == 0 else ""
            color = "FFFFFF" if ridx == 0 else "243746"
            bold = "<w:b/>" if ridx == 0 else ""
            out.append(
                f'<w:tc><w:tcPr><w:tcW w:w="{widths[cidx]}" w:type="dxa"/>{fill}'
                f'<w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:spacing w:after="35" w:before="25"/></w:pPr>'
                f'<w:r><w:rPr>{bold}<w:color w:val="{color}"/><w:sz w:val="17"/></w:rPr>'
                f'<w:t xml:space="preserve">{escape(str(val))}</w:t></w:r></w:p></w:tc>'
            )
        out.append("</w:tr>")
    out.append("</w:tbl>")
    return "".join(out)


def page(title: str, subtitle: str, blocks: list[str], break_after: bool = True) -> str:
    parts = [para(title, "Heading1"), para(subtitle, "Subtitle")]
    parts.extend(blocks)
    if break_after:
        parts.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
    return "".join(parts)


def build_document() -> str:
    pages = []

    pages.append(page("OISI Metrics and Scoring Methodology", "Submission document | Four components | 112 infrastructures | scoring build dated 14 October 2026", [
        para("Purpose and scope", "Heading2"),
        para("The Open Infrastructure Sustainability Index (OISI) is a structured comparison of open infrastructures across four areas: Community Engagement, Governance Practices, Organizational Capacity, and Development Activity. Each component turns observable evidence into a 0–100 score. The combined index gives each component equal influence and retains component-level results so reviewers can inspect what is behind the total."),
        para("The framework assesses visible organizational and community signals associated with continuity and capacity. It is not a direct measurement of future survival, software quality, social impact, or user satisfaction. Open source health research cautions that repository traces alone cannot represent the full social and institutional context of a project [1, 2]. OISI therefore combines repository activity with documented community, governance, and organizational evidence. This broader view is consistent with CHAOSS’s goal of making open source community-health measures transparent and actionable [5]."),
        para("Four-component structure", "Heading2"),
        table(["Component", "Index weight", "What it represents"], [
            ["Community Engagement", "25%", "Engagement channels, activity breadth, contribution opportunities and norms."],
            ["Governance Practices", "25%", "Formal decision structures, participation, transparency and public commitments."],
            ["Organizational Capacity", "25%", "Maintenance, institutional support and operational/service capacity."],
            ["Development Activity", "25%", "Recent continuity, contributor distribution, responsiveness and adoption signals."],
        ], [2500, 1100, 5400]),
        para("Why equal component weights?", "Heading2"),
        para("The four domains represent distinct, complementary conditions rather than interchangeable measures. Assigning 25% to each prevents the most data-rich component from dominating the index and makes the top-level rule transparent. This is a normative design choice for balanced comparison; no cited study establishes that 25% is a uniquely correct empirical weight. Composite-indicator guidance recommends making the framework, weighting choices, aggregation method, and sensitivity limits explicit [3]."),
        para("How to read a score", "Heading2"),
        para("A higher score means stronger measured evidence under these definitions and weights. It does not prove that one infrastructure is universally better. Scores should be read with the component profile, data coverage, and source notes. Small differences should be treated cautiously, especially where evidence completeness differs."),
        para("Weighting convention", "Heading2"),
        para("Within components, metrics are weighted according to their role and intended importance. If a component permits a score with some metrics unavailable, the observed metric weights are renormalized for that component. The combined index then assigns the resulting component score its fixed 25% share. Thus component-level missing-metric handling and combined-index missing-component handling are separate steps."),
    ]))

    pages.append(page("1. Community Engagement", "Component weight in OISI: 25% | Seven metrics | Three sub-components", [
        para("This component measures whether people can find ways to engage, contribute, and understand expected community conduct. Its metric weights total 100% within the component. Engagement & Participation receives 40%, Contribution Mechanisms 45%, and Community Norms 15%. The larger contribution share reflects the component’s focus on actionable participation routes, while the engagement share captures recurring ways to connect. Norms are included as an enabling condition."),
        table(["Code and metric", "Weight", "Operational meaning and scoring"], [
            ["E1 Engagement availability", "10%", "Source field 037. Status ladder: implemented/available 1.00; in progress 0.67; considering 0.33; not planning 0.00. Confirms whether engagement is an explicit, planned capability."],
            ["E2 Engagement activity breadth", "25%", "Fields 038 plus coded 938. Distinct engagement categories are combined; score = min(category count, 8) / 8. Capped at eight categories to reward variety without allowing long lists to dominate."],
            ["E3 Engagement evidence strength", "5%", "Fields 116 and 938. A direct engagement URL scores 1.00; a substantive 938 description without a URL scores 0.50; otherwise unavailable. A small weight recognizes corroborating evidence without double-counting activity breadth."],
            ["C1 Contribution opportunity breadth", "25%", "Fields 040 plus 940. Distinct contribution opportunity types; score = min(count, 6) / 6. Rewards multiple routes into work, capped at six."],
            ["C2 Contribution guidance", "15%", "Field 041. A recognizable contribution guide scores 1.00; another relevant participation page scores 0.50; blank or suspect URL is unavailable pending review."],
            ["C3 Open contribution infrastructure", "5%", "Field 019. Public forge URL scores 1.00; another repository URL scores 0.50; blank is unavailable. Kept small because a repository alone does not establish an inclusive contribution process."],
            ["N1 Community norms / Code of Conduct", "15%", "Fields 042 and 043. Uses the status ladder above; when status is implemented but no evidence URL is supplied, score is reduced to 0.75. Measures stated behavioral expectations and evidence."],
        ], [1950, 700, 6350]),
        para("Why these percentages?", "Heading2"),
        para("E2 and C1 receive the largest shares (25% each) because breadth of engagement and breadth of concrete contribution routes are central constructs, and their capped counts permit consistent comparison. C2 receives 15% because guidance can reduce entry friction. N1 receives 15% as a basic community safety and expectation signal. E1 (10%) identifies whether engagement is planned; E3 (5%) and C3 (5%) are supporting evidence/access signals and receive lower weights to limit overlap with the core breadth measures."),
        para("Calculation and eligibility", "Heading2"),
        para("For observed metric scores x and weights w, the component score is 100 × Σ(wx) / Σ(w observed). A score is eligible when at least 60% of metric weight is observed and at least one metric is scored in each of the three sub-components. Completeness of at least 85% is labelled Reportable; eligible scores below 85% are Provisional. Unknown, blank, and Not applicable values are excluded from the component denominator. Field 039 is not scored; it is used only for consistency flags. Community governance field 048 is not part of this component. N1 is the only Code of Conduct measure: implemented without a URL scores 0.75; Governance Practices does not score that field, so its evidence-confidence flag is not a directly comparable Code of Conduct rule."),
    ]))

    pages.append(page("2. Governance Practices", "Component weight in OISI: 25% | Observed-weight scoring | Current governance workbook", [
        para("Governance measures how publicly visible decision structures, participation arrangements, and commitments are. Status fields use an ordinal ladder: Implemented/available = 1.00; In progress = 0.67; Considering = 0.33; Not planning = 0.00. Unknown and Not applicable are normally missing. The sub-component weights are Formal 40%, Participatory 25%, and Transparency 35%."),
        table(["Metric", "Weight within component", "Role in the score"], [
            ["024 Governance structure", "12% Formal", "Evidence that a governance structure is defined; core institutional decision signal."],
            ["026 Governance records", "12% Formal", "Public records of governance or decisions support accountability and traceability."],
            ["108 Bylaws", "8% Formal", "Written rules indicate formalization; Unknown is treated as missing."],
            ["050 Board structure", "8% Formal", "Presence of a board or steering body. The type of body is descriptive, not ranked."],
            ["048 Community governance", "20 nominal points; effective role = 25% Participatory", "Formal = 1.00; Ad hoc = 0.50; blank is scored 0 under the workbook’s current setting. It is the only scored indicator in Participatory, so its sub-component score receives the fixed 25% Participatory weight."],
            ["100 Board level", "5 nominal points; 0% scored", "Descriptive only because coverage is 15/112. It is excluded from the scored metric denominator. The workbook notes that its nominal points go to 048; mechanically, the score formula gives the Participatory sub-component 25% and 048 supplies that sub-component score."],
            ["029 Web accessibility", "7% Transparency", "Public accessibility commitment or status."],
            ["033 Transparent pricing", "7% Transparency", "Pricing transparency, including the workbook’s rule for No direct costs."],
            ["035 Equity commitment", "7% Transparency", "Stated commitment to equity."],
            ["044 Privacy commitment", "7% Transparency", "Stated commitment to privacy."],
            ["046 Values frameworks", "7% Transparency", "Presence of a named framework. Blank is scored 0 under the current workbook setting."],
        ], [2700, 1300, 5000]),
        para("Why these percentages?", "Heading2"),
        para("Formal governance receives 40% because structures, records, rules, and a governing body document how authority is organized. Transparency receives 35% because public commitments and accessible information allow external scrutiny. Participatory governance receives 25%: it is essential to community influence, but the current measured evidence is concentrated in one field (048), while 100_board_level is too sparsely populated to score reliably. Within Formal, structure and records have the largest shares (12% each); bylaws and board presence are supporting signals (8% each). Transparency indicators are evenly weighted at 7% each because they cover distinct public-accountability topics."),
        para("Calculation, eligibility, and reviewer note", "Heading2"),
        para("Status values become 0–1 metric scores. Presence variables are scored 1 when a value is listed; the current Parameters sheet sets blank 050, 046, and 048 to 0. Metric weights are renormalized within each observed sub-component; the component score is 100 times the weighted mean of observed Formal (40%), Participatory (25%), and Transparency (35%) sub-component scores. A score is published at 70% completeness only if Formal and Transparency are both represented. Evidence confidence is a separate reliability flag: the share of Implemented claims with an evidence URL; it does not change the score."),
        para("Current workbook caveats: 048 is present in Governance Practices, even though it is excluded from Community Engagement. Its 25 blank values are treated as zero by the existing rule and should be interpreted as an operational coding assumption, not a confirmed ‘None’ response. The source workbook also flags self-reported evidence, the low coverage of 100, and limited discrimination from 033 pricing responses. A sensitivity comparison that treats blank 048, 050, and 046 values as missing is included in the accompanying index workbook; under that alternative, 89 rather than 108 Governance scores remain eligible."),
    ]))

    pages.append(page("3. Organizational Capacity", "Component weight in OISI: 25% | Three sub-components | Applicability-aware denominator", [
        para("Organizational Capacity assesses whether the infrastructure has a maintained technical base, institutional support, and an operational service path. The internal weights are Maturity & Maintenance 35%, Institutional Support 35%, and Operational Capacity 30%. Technology and maintenance are deliberately limited to 35% together because these source fields are relatively uniform and could otherwise inflate the component score."),
        table(["Code and metric", "Weight", "Scoring rule and interpretation"], [
            ["M1 Technology readiness", "10%", "Field 079. Operational system = 1.00; complete/qualified system = 8/9; validated relevant environment = 5/9; otherwise NA. Distinguishes demonstrated readiness levels."],
            ["M2 Maintenance status", "20%", "Field 080. Actively maintained = 1.00; minimally maintained = 0.50; inactive/not maintained = 0.00; unknown/blank = NA. Highest maturity weight because upkeep directly signals continued stewardship."],
            ["M3 Infrastructure age", "5%", "Field 056. Score = min(age in years, 15) / 15, using 2026 as the as-of year. A modest tenure signal, not a claim that older is always better."],
            ["I1 Formal organizational host", "15%", "Field 051. Formal organization (including academic/nonprofit/commercial host) = 1.00; volunteer community = 0.50; other/blank = NA. Indicates an identifiable institutional anchor."],
            ["I2 Affiliations / institutional support", "15%", "Field 058. Named affiliation/support = 1.00; explicit no affiliations = 0.00; blank = NA. Measures stated external support relationships."],
            ["I3 SCOSS participation", "5%", "Field 099. TRUE = 1.00; FALSE = 0.00; unknown/blank = NA. Kept small as a specific support signal rather than a universal requirement."],
            ["O1 Hosted/SaaS availability", "20%", "Field 076. Provider or third-party hosted service available = 1.00; no services available = 0.00; Not applicable/unknown = NA. Captures an operational access route."],
            ["O2 Registered service-provider ecosystem", "10%", "Field 077. If third-party delivery applies: provider list = 1.00, no list = 0.00. Otherwise NA. Conditional to avoid penalizing models where it does not apply."],
        ], [2350, 750, 5900]),
        para("Why these percentages?", "Heading2"),
        para("Maintenance (20%) and Hosted/SaaS service availability (20%) are tied for the largest single metric weights. Upkeep is a direct stewardship signal; service availability captures an operational access route. Host and affiliation (15% each) represent complementary forms of institutional support; the ecosystem measure (10%) adds evidence about third-party provision. Readiness (10%) and age (5%) are supporting technical context, while SCOSS participation (5%) is intentionally not treated as necessary for every viable infrastructure."),
        para("Calculation and eligibility", "Heading2"),
        para("The score is 100 × Σ(metric score × metric weight) / Σ(scored applicable metric weights). Unknown, blank, and not applicable are excluded from the denominator. O1 is removed from the applicable denominator when 076 is Not applicable; O2 is considered only when third-party delivery applies. At least 70% applicable completeness and at least two of the three sub-components are required. Eligible scores at 85% completeness are Reportable; 70% to below 85% are Provisional. Lower and upper bounds in the workbook show outcomes if missing applicable metrics were all zero or all one."),
    ]))

    pages.append(page("4. Development Activity", "Component weight in OISI: 25% | 13 repository metrics | Fixed 12-month observation window", [
        para("Development Activity uses repository history and issue/PR data from 1 October 2025 inclusive through 1 October 2026 exclusive (UTC). Metrics are normalized to 0–100. Bots are excluded; merge commits are excluded from authorship metrics; human identities are resolved using repository evidence and documented rules. Repositories that are dormant can receive zero on activity metrics when the underlying Git history is complete; this differs from missing data. The selection of trace-based indicators follows the broader community-health measurement practice described by CHAOSS, while retaining the limitation that trace data do not fully explain project context [4]."),
        table(["Dimension / metrics", "Weights", "Measures and score anchors"], [
            ["Continuity: m1 Human commits; m2 Active weeks; m3 Recency; m4 Release activity", "10%, 15%, 5%, 5% (35%)", "m1: log scale, capped at 1,000 commits. m2: active weeks / 52. m3: 100 through 30 days since last commit, then linear decline to 0 at 365 days. m4 tags: 0=0, 1=50, 2–3=75, 4+=100."],
            ["Contributors: m5 Active contributors; m6 Effective contributors; m7 Retention", "10%, 10%, 5% (25%)", "m5: log2 scale capped at 32 people (one contributor scores 0). m6: inverse concentration E=1/Σ(share²), log2 scale capped at E=8. m7 contributors active in both half-windows: 0=0, 1=25, 2=50, 3–4=75, 5+=100."],
            ["Responsiveness: m8 Issue response; m9 PR response; m10 Issue close ≤90d; m11 PR decision ≤30d", "10%, 10%, 7.5%, 7.5% (35%)", "Each is the percentage meeting its time rule. Response windows: issue 14 days, PR 7 days. Decision windows: issue closure 90 days, PR merge/close 30 days. Each metric requires at least 10 eligible items."],
            ["Adoption: m12 Stars gained; m13 Forks created", "2.5%, 2.5% (5%)", "Log-scaled within-window counts, capped at 1,000 stars and 100 forks. Equal small weights treat these as indirect signals rather than direct measures of quality or impact."],
        ], [2200, 1500, 5300]),
        para("Why these percentages?", "Heading2"),
        para("Continuity and Responsiveness receive 35% each: sustained work and timely handling of user/contributor work are core operational signals, and neither alone is sufficient. Contributors receive 25% to capture capacity and distribution beyond raw output. Adoption receives 5% because stars and forks are indirect, platform-dependent attention signals; giving them a small equal share limits their influence. Within Continuity, active weeks (15%) emphasizes sustained cadence; commits (10%) reflect volume; recency and releases (5% each) are supporting signals. Contributor count and distribution each receive 10%; retention receives 5% as a narrower longitudinal signal. Responsiveness divides 20% equally across response rates and 15% across outcome-time metrics."),
        para("Why equal star and fork weights?", "Heading2"),
        para("Stars and forks represent different forms of attention or reuse and neither is treated as a direct quality measure. Their 2.5% each avoids allowing one platform action to dominate adoption. The component workbook’s sensitivity analysis reports negligible ranking change under asymmetric reweighting; this supports keeping the transparent equal split, but it does not prove the two signals are conceptually identical."),
        para("Eligibility", "Heading2"),
        para("The weighted mean is calculated over available metric weights. A component score is issued only when at least 75% of total metric weight is available and at least half of each dimension’s weight is available. This prevents scores based only on Git activity or only on GitHub responsiveness/adoption. For eligible repositories, the score is the normalized weighted mean over available metrics. Low issue/PR volume and non-GitHub trackers can make responsiveness unavailable; this is not evidence that the project failed to respond."),
    ]))

    pages.append(page("5. Equal Metric Weight Perspective", "Alternate calculation | Equal weight for each scored metric within its component", [
        para("This perspective asks what the results look like when no metric is given a larger share within its component. It uses the same metric set, source values, score anchors, and applicability rules as the current model. It replaces within-component metric weights with a simple mean over scored/applicable metrics. The four components remain equally weighted at 25% each in the combined OISI. This is a comparison scenario, not a replacement of the current weighted scores."),
        table(["Component", "Equal metric share", "Implied sub-component / dimension shares"], [
            ["Community Engagement (7 metrics)", "Each metric = 1/7 (14.29%)", "Engagement 3/7 (42.86%); Contribution 3/7 (42.86%); Norms 1/7 (14.29%)."],
            ["Governance Practices (10 scored metrics)", "Each metric = 1/10 (10%)", "Formal 4/10 (40%); Participatory 1/10 (10%); Transparency 5/10 (50%). Field 100 remains descriptive and excluded."],
            ["Organizational Capacity (8 metrics)", "Each applicable metric = 1/8 (12.5% when all apply)", "Maturity 3/8 (37.5%); Institutional 3/8 (37.5%); Operational 2/8 (25%). Not-applicable O1/O2 fields leave the denominator."],
            ["Development Activity (13 metrics)", "Each metric = 1/13 (7.69%)", "Continuity 4/13 (30.77%); Contributors 3/13 (23.08%); Responsiveness 4/13 (30.77%); Adoption 2/13 (15.38%)."],
        ], [2450, 2000, 4500]),
        para("Equal metric scoring rule", "Heading2"),
        para("For each component, add the available normalized metric scores and divide by the number of scored/applicable metrics. Community Engagement requires at least 60% of its seven metrics and coverage in all three sub-components. Governance requires at least 70% of ten active indicators and both Formal and Transparency represented; blank 048/050/046 continue to score zero under the current input coding. Organizational Capacity requires at least 70% of applicable metrics and two sub-components. Development Activity requires at least 75% of 13 metrics and at least half of each dimension’s metric count. These are count-based equivalents of the existing completeness safeguards under equal metric weights."),
        para("Results and comparison", "Heading2"),
        para("Under equal metric weights, 88 Community Engagement, 108 Governance Practices, 109 Organizational Capacity, and 72 Development Activity scores are eligible. Fifty-seven infrastructures have equal-metric scores in all four components; their mean is 74.54 and range is 46.90–91.52. The all-112 equal-metric score with unavailable components set to zero has a mean of 61.46 and range of 18.91–91.52."),
        para("Against the existing weighted method, the equal-metric combined score has Spearman rank correlation 0.966 and 10/10 top-ten overlap on the 53 records complete under both methods. Across all 112 zero-filled results, correlation is 0.937 and top-ten overlap is 7/10. By component, correlations range from 0.764 for Organizational Capacity to 0.976 for Development Activity. The largest conceptual shift is in Governance: the single Participatory metric 048 moves from a 25% sub-component share in the current model to 10% under equal metric weighting, while Transparency rises from 35% to 50%."),
        para("The equal-metric alternative is easier to explain as ‘one scored indicator, one share,’ but it does not remove judgment: the metric list, score anchors, eligibility thresholds, and treatment of missing values still shape the outcome. Both views should therefore be shown with component scores and coverage, and neither should be presented as a uniquely correct weighting scheme."),
    ]))

    pages.append(page("6. Combined Index, Missing Data, and Verification", "Interpretation, reproducibility, and reviewer-facing limitations", [
        para("Combined OISI calculation", "Heading2"),
        para("For complete records: OISI = 0.25 × Community Engagement + 0.25 × Governance Practices + 0.25 × Organizational Capacity + 0.25 × Development Activity. The current all-112 presentation also reports an evidence-adjusted score: (CE + Governance + Organizational Capacity + Development Activity) / 4, with each unavailable component score contributing 0. This makes a 0–100 value available for all entries but is a conservative evidence convention, not an observed zero-performance score."),
        para("Worked example: 2i2c has Community Engagement 68.75, Governance Practices 48.72, Organizational Capacity 85.45, and Development Activity 74.00. Its evidence-adjusted OISI is (68.75 + 48.72 + 85.45 + 74.00) / 4 = 69.23. All four component scores are available, so its evidence-adjusted and strict four-component scores are identical."),
        para("Coverage and current cohort", "Heading2"),
        table(["Measure", "Current result"], [
            ["Master set", "112 infrastructures"],
            ["Component scores available", "Community Engagement 83; Governance Practices 108; Organizational Capacity 108; Development Activity 72"],
            ["Four-component coverage", "53 complete; 58 partial (42 with 3 components, 16 with 2); 1 with 1 component"],
            ["Development Activity completeness", "45 of the 72 scored entries have all 13 metric scores available"],
            ["Development Activity not scored", "40 entries: 10 excluded/not attempted and 30 insufficient data. Of the 30, 24 have low issue/PR volume and 6 use non-GitHub tracker/platform cases."],
            ["Evidence-adjusted all-112 results", "Mean 61.76; range 23.12–94.19"],
        ], [3000, 6000]),
        para("Primary comparison and all-112 view", "Heading2"),
        para("Use the 53 infrastructures with all four component scores for the primary like-for-like ranking. Keep the zero-filled score for all 112 as a secondary evidence-coverage view. It assigns zero points to an unavailable component score while retaining missing component names, source status, and coverage. It is not an observed zero-performance result. The available-component average remains a diagnostic only because low coverage can make it appear high."),
        para("Sensitivity analysis", "Heading2"),
        para("The accompanying workbook compares: (a) zero-filled versus available-component averages across all 112; (b) the Governance result when blank 048, 050, and 046 values are treated as missing rather than zero; and (c) four one-at-a-time weight perturbations that raise one component from 25% to 30% and reduce each other component proportionally. Weight scenarios are compared on the same 53 complete cases. The workbook reports Spearman rank correlation and top-ten overlap against the equal-weight complete-case baseline. These are stress tests, not proposed replacement weights. Results: the available-component average has rank correlation 0.534 and 4/10 top-ten overlap versus the all-112 zero-filled ranking; the Governance blank-as-missing alternative has 0.962 rank correlation and 9/10 top-ten overlap versus the zero-filled ranking. The four modest weight perturbations have correlations from 0.992 to 0.995 and top-ten overlap from 9/10 to 10/10 on complete cases."),
        para("Complete-case score spread and correlation", "Heading2"),
        para("Among the 53 complete cases, component standard deviations are: Community Engagement 14.98, Governance Practices 22.35, Organizational Capacity 9.75, and Development Activity 15.00. Pairwise Pearson correlations range from -0.061 to 0.349. Equal numeric weights therefore do not imply equal influence on rank order; the accompanying sensitivity sheet makes rank stability visible. The low correlations also support retaining component profiles alongside the composite."),
        para("Verification workflow", "Heading2"),
        para("1) Confirm the same 112 infrastructures are represented in each source and resolve duplicate or ambiguous name matches. 2) Recalculate each component score from its own metric values, weights, and eligibility rules. 3) Check component statuses, completeness, and provenance against the source sheets. 4) Confirm the combined formula, four fixed 25% weights, zero-imputation flag, coverage count, and rank. 5) Review sensitivity to alternative assumptions and report the score alongside component scores and coverage. The workbook’s Codebook, Summary, Scoring Rules, Component Scores, Verification/Diagnostics, and Data sheets support these checks."),
        para("Limitations to disclose", "Heading2"),
        para("The component data combine self-reported profile fields and repository traces; evidence quality and platform coverage vary. GitHub-centered response metrics do not fully represent activity on Jira, GitLab, or other systems. Stars and forks are indirect adoption measures. Status scales and capped category counts simplify varied evidence into common scores. Equal top-level weights and within-component percentages are explicit design judgments, not uniquely determined by research. Rankings are descriptive under the chosen rules and should not be interpreted as causal claims or definitive predictions of sustainability."),
        para("References", "Heading2"),
        para("[1] Goggins, S., Lumbard, K., & Germonprez, M. (2021). Open Source Community Health: Analytical Metrics and Their Corresponding Narratives. IEEE/ACM SoHeal. https://doi.org/10.1109/SoHeal52568.2021.00010"),
        para("[2] Yin, L., Chakraborti, M., Yan, Y., Schweik, C., Frey, S., & Filkov, V. (2022). Open Source Software Sustainability: Combining Institutional Analysis and Socio-Technical Networks. Proceedings of the ACM on Human-Computer Interaction, 6(CSCW2), Article 404. https://doi.org/10.1145/3555129"),
        para("[3] OECD, European Union & European Commission JRC. (2008). Handbook on Constructing Composite Indicators: Methodology and User Guide. OECD Publishing. https://doi.org/10.1787/9789264043466-en"),
        para("[4] CHAOSS Project. CHAOSS Metrics: implementation-agnostic measures for understanding open source community health. https://chaoss.community/metrics/"),
        para("[5] Linux Foundation CHAOSS. About CHAOSS and its open source community-health measurement mission. https://www.chaoss.community/about-chaoss/"),
    ], break_after=False))

    body = "".join(pages)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:body>{body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="650" w:right="720" w:bottom="650" w:left="720" w:header="300" w:footer="300"/>'
        '<w:cols w:space="720"/></w:sectPr></w:body></w:document>'
    )


def main() -> None:
    document = build_document()
    with ZipFile(OUT, "w", ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>''')
        z.writestr("_rels/.rels", '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>''')
        z.writestr("word/document.xml", document)
        z.writestr("word/styles.xml", '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/><w:sz w:val="18"/><w:color w:val="243746"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="85" w:line="220" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="0" w:after="90"/></w:pPr><w:rPr><w:rFonts w:ascii="Aptos Display" w:hAnsi="Aptos Display"/><w:b/><w:color w:val="1F4E78"/><w:sz w:val="27"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="75" w:after="45"/></w:pPr><w:rPr><w:b/><w:color w:val="1F4E78"/><w:sz w:val="20"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="150"/></w:pPr><w:rPr><w:color w:val="58717F"/><w:sz w:val="17"/></w:rPr></w:style>
<w:style w:type="table" w:styleId="GridTable"><w:name w:val="Grid Table"/></w:style>
</w:styles>''')
    print(OUT)


if __name__ == "__main__":
    main()
