from __future__ import annotations

from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
from xml.sax.saxutils import escape


OUT = Path("OISI_Equal_Metric_Weight_Methodology.docx")


def run(text: str, bold: bool = False) -> str:
    props = '<w:rPr><w:b/></w:rPr>' if bold else ''
    return f'<w:r>{props}<w:t xml:space="preserve">{escape(str(text))}</w:t></w:r>'


def para(text: str = "", style: str = "Normal") -> str:
    return f'<w:p><w:pPr><w:pStyle w:val="{style}"/></w:pPr>{run(text)}</w:p>'


def table(headers: list[str], rows: list[list[str]], widths: list[int] | None = None) -> str:
    widths = widths or [9000 // len(headers)] * len(headers)
    chunks = ['<w:tbl><w:tblPr><w:tblStyle w:val="GridTable"/><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/><w:tblBorders><w:top w:val="single" w:sz="5" w:color="B7C9D6"/><w:left w:val="single" w:sz="5" w:color="B7C9D6"/><w:bottom w:val="single" w:sz="5" w:color="B7C9D6"/><w:right w:val="single" w:sz="5" w:color="B7C9D6"/><w:insideH w:val="single" w:sz="4" w:color="D7E0E7"/><w:insideV w:val="single" w:sz="4" w:color="D7E0E7"/></w:tblBorders></w:tblPr>']
    for ri, row in enumerate([headers] + rows):
        chunks.append('<w:tr>')
        for ci, val in enumerate(row):
            head = ri == 0
            fill = '<w:shd w:fill="1F4E78"/>' if head else ''
            color = 'FFFFFF' if head else '243746'
            bold = '<w:b/>' if head else ''
            chunks.append(f'<w:tc><w:tcPr><w:tcW w:w="{widths[ci]}" w:type="dxa"/>{fill}<w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:spacing w:after="45" w:before="35"/></w:pPr><w:r><w:rPr>{bold}<w:color w:val="{color}"/><w:sz w:val="16"/></w:rPr><w:t xml:space="preserve">{escape(str(val))}</w:t></w:r></w:p></w:tc>')
        chunks.append('</w:tr>')
    chunks.append('</w:tbl>')
    return ''.join(chunks)


def page(title: str, subtitle: str, blocks: list[str], last: bool = False) -> str:
    content = para(title, 'Heading1') + para(subtitle, 'Subtitle') + ''.join(blocks)
    if not last:
        content += '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
    return content


def build() -> str:
    pages = []
    pages.append(page('OISI Equal-Metric Weighting Methodology', 'Standalone method note | Four components | 112 infrastructures | 14 October 2026', [
        para('Purpose and scope', 'Heading2'),
        para('This document specifies an alternative OISI calculation in which every scored, applicable metric receives the same share within its own component. It is intended to be read alongside the existing component score sheets and the OISI full-index workbook. It documents the equal-metric perspective; it does not replace the original weighted results.'),
        para('The distinction that matters', 'Heading2'),
        table(['Level', 'Rule in this analysis'], [
            ['Within a component', 'Equal arithmetic mean across the component\'s scored, applicable metrics.'],
            ['Across the four components', 'Each component remains 25% of the combined OISI.'],
            ['Metric definitions and anchors', 'Retained from the current component scoring sheets; this alternative changes metric influence, not the evidence or scoring anchors.'],
            ['Unavailable metrics', 'Not silently turned into zero. They are omitted from that component mean, subject to the minimum coverage safeguards below.'],
        ], [2400, 6600]),
        para('Formula', 'Heading2'),
        para('For component j, let S_j be the set of scored and applicable metrics and x_ij their normalized 0-100 values. The equal-metric component score is C_j = sum(x_ij for i in S_j) / |S_j|. The strict combined index, when all four components qualify, is OISI = 0.25*CE + 0.25*GP + 0.25*OC + 0.25*DA.'),
        para('Why examine equal metric weights?', 'Heading2'),
        para('Equal shares provide an interpretable benchmark: no indicator receives extra influence simply because a prior design assigned it a larger percentage. This makes it easier to see how much conclusions depend on the original within-component weighting choices. It is a transparent normative scenario, not a claim that every indicator has identical construct validity, reliability, or causal importance. Composite-indicator guidance recommends making weighting and aggregation assumptions explicit and testing their influence [1]. Open-source health research also cautions against relying on a single repository signal to represent a multidimensional community [2, 3].'),
        para('The four component weights are still 25% each. Equal metric weighting does not mean pooling all 38 indicators into one mean: each component is calculated first, then the four component scores are combined.'),
    ]))
    pages.append(page('1. Community Engagement', '7 active metrics | 14.29% each when all seven are scored', [
        table(['Metric', 'Metric share', 'Existing score meaning retained'], [
            ['E1 Engagement availability', '1/7 = 14.29%', 'Availability/status ladder from the current CE scoring sheet.'],
            ['E2 Engagement activity breadth', '1/7 = 14.29%', 'Breadth count from fields 038 and 938, normalized and capped using the current anchor.'],
            ['E3 Engagement evidence strength', '1/7 = 14.29%', 'Evidence strength from fields 116 and 938 using the current URL/description rule.'],
            ['C1 Contribution opportunity breadth', '1/7 = 14.29%', 'Breadth count from fields 040 and 940, normalized and capped using the current anchor.'],
            ['C2 Contribution guidance', '1/7 = 14.29%', 'Contribution guidance score from field 041.'],
            ['C3 Open contribution infrastructure', '1/7 = 14.29%', 'Public forge/repository accessibility score from field 019.'],
            ['N1 Community norms / Code of Conduct', '1/7 = 14.29%', 'Status and evidence rule from fields 042-043; existing implemented-without-URL adjustment is retained.'],
        ], [2850, 1350, 4800]),
        para('Implied group influence', 'Heading2'),
        table(['Group', 'Metrics', 'Equal-metric share'], [
            ['Engagement', 'E1-E3', '3/7 = 42.86%'],
            ['Contribution', 'C1-C3', '3/7 = 42.86%'],
            ['Norms', 'N1', '1/7 = 14.29%'],
        ], [2300, 1800, 4900]),
        para('Eligibility and missingness', 'Heading2'),
        para('A component score requires at least 5 of 7 metrics (the threshold is at least 60%) and at least one observed metric in each of Engagement, Contribution, and Norms. At least 6 of 7 metrics (85%) is labelled Reportable; otherwise an eligible score is Provisional. Blank, unknown, or not-applicable values are excluded rather than assumed to be zero. Field 039 is descriptive/flag-only and is not one of the seven scored metrics. Governance field 048 is not part of this component.'),
        para('Interpretation', 'Heading2'),
        para('Under equal weighting, the two three-metric groups each account for 42.86% of a fully observed CE score, while the single norms metric accounts for 14.29%. This is a consequence of equal metric shares, not an additional decision to weight groups equally.'),
    ]))
    pages.append(page('2. Governance Practices', '10 active score indicators | 10% each | board-level field excluded', [
        table(['Active indicator', 'Metric share', 'Scoring content retained'], [
            ['s024 Governance structure', '10%', 'Current status score for a documented governance structure.'],
            ['s026 Governance records', '10%', 'Current status score for public governance/decision records.'],
            ['s108 Bylaws', '10%', 'Current score for bylaws; unknown remains missing.'],
            ['s050 Board structure', '10%', 'Presence score as coded in the source workbook.'],
            ['s048 Community governance', '10%', 'Formal = 1; Ad hoc = 0.5; None = 0. Existing blank-as-zero coding retained.'],
            ['s029 Web accessibility', '10%', 'Current status score for accessibility.'],
            ['s033 Transparent pricing', '10%', 'Current transparent-pricing rule, including the existing no-direct-costs mapping.'],
            ['s035 Equity commitment', '10%', 'Current status score for an equity commitment.'],
            ['s044 Privacy commitment', '10%', 'Current status score for a privacy commitment.'],
            ['s046 Values frameworks', '10%', 'Presence score as coded in the source workbook; blank-as-zero retained.'],
        ], [2600, 1100, 5250]),
        para('Implied group influence', 'Heading2'),
        table(['Group', 'Indicators', 'Equal-metric share'], [
            ['Formal', 's024, s026, s108, s050', '4/10 = 40%'],
            ['Participatory', 's048', '1/10 = 10%'],
            ['Transparency', 's029, s033, s035, s044, s046', '5/10 = 50%'],
        ], [2300, 3500, 3200]),
        para('Eligibility, coding caveat, and exclusions', 'Heading2'),
        para('At least 7 of 10 active indicators (70%) must be scored, with at least one Formal and one Transparency indicator represented. Unknown status values remain missing. Existing source coding treats blank s048, s050, and s046 values as zero; this equal-metric comparison preserves that choice so the method comparison changes weighting rather than recoding evidence. A blank coded zero is an assumption and should not be described as a verified absence.'),
        para('The 100_board_level field is descriptive and excluded from the ten scored indicators because its coverage is too sparse for stable comparison. No Governance parameter is added or removed by the equal-metric method; 048 remains a Governance metric, not a Community Engagement metric.'),
    ]))
    pages.append(page('3. Organizational Capacity', '8 candidate metrics | 12.5% each when all eight apply', [
        table(['Metric', 'Metric share', 'Scoring meaning retained'], [
            ['M1 Technology readiness', '1/8 = 12.5%', 'Readiness levels from field 079.'],
            ['M2 Maintenance status', '1/8 = 12.5%', 'Active/minimal/inactive maintenance mapping from field 080.'],
            ['M3 Infrastructure age', '1/8 = 12.5%', 'Capped age score from field 056, as in the current workbook.'],
            ['I1 Formal organizational host', '1/8 = 12.5%', 'Host type score from field 051.'],
            ['I2 Affiliations / institutional support', '1/8 = 12.5%', 'Affiliation/support score from field 058.'],
            ['I3 SCOSS participation', '1/8 = 12.5%', 'TRUE/FALSE mapping from field 099.'],
            ['O1 Hosted/SaaS service availability', '1/8 = 12.5%', 'Service availability score from field 076, where applicable.'],
            ['O2 Registered service-provider ecosystem', '1/8 = 12.5%', 'Provider ecosystem score from field 077, where applicable.'],
        ], [3000, 1350, 4600]),
        para('Implied dimension shares when all metrics apply', 'Heading2'),
        table(['Dimension', 'Metrics', 'Equal-metric share'], [
            ['Maturity and maintenance', 'M1-M3', '3/8 = 37.5%'],
            ['Institutional support', 'I1-I3', '3/8 = 37.5%'],
            ['Operational capacity', 'O1-O2', '2/8 = 25%'],
        ], [2600, 1900, 4500]),
        para('Applicability and eligibility', 'Heading2'),
        para('The denominator is the number of applicable metrics. O1 is excluded when field 076 is explicitly Not applicable. O2 applies only when the SaaS/provider arrangement indicates third-party delivery; otherwise it is removed from the applicable denominator. Other unknown/blank metric values are missing, not zero. A score requires at least 70% of applicable metrics and at least two of the three dimensions represented. Eligible scores with at least 85% applicable coverage are Reportable; scores from 70% to below 85% are Provisional.'),
        para('Why the dimension shares vary', 'Heading2'),
        para('Equal shares attach to metrics, not dimensions. Therefore, dimensions with three metrics each account for 37.5% when all apply, and the two-metric operational dimension accounts for 25%. If a metric is structurally not applicable, its share is redistributed across the remaining scored/applicable metrics by the arithmetic mean; it is not treated as a performance failure.'),
    ]))
    pages.append(page('4. Development Activity', '13 repository metrics | 7.69% each | observation window: 1 Oct 2025 to 1 Oct 2026 UTC', [
        para('The normalized metric values and their operational definitions are retained from the Development Activity scoring workbook. For the equal-metric view, each available metric contributes 1/13 = 7.69% of the component mean; count-based score anchors below are unchanged.'),
        table(['Metric', 'Definition / scoring anchor'], [
            ['m1 Human commits', 'Log scale: 100*min(1, log10(1+c)/log10(1001)); c capped at 1,000.'],
            ['m2 Active weeks', 'Distinct active weeks / 52 * 100.'],
            ['m3 Recency', '100 through 30 days since last commit; then linear decline to 0 at 365 days.'],
            ['m4 Release activity', 'Tags: 0=0; 1=50; 2-3=75; 4+=100.'],
            ['m5 Active contributors', '0 when n<=1; otherwise 100*min(1, log2(n)/5), capped at 32 contributors.'],
            ['m6 Effective contributors', 'E=1/sum(commit_share^2); 0 when E<=1; log2-normalized and capped at E=8.'],
            ['m7 Contributor retention', 'Active in both half-windows: 0=0; 1=25; 2=50; 3-4=75; 5+=100.'],
            ['m8 Issue response', 'Share of eligible issues receiving a response within 14 days; minimum 10 eligible items.'],
            ['m9 PR response', 'Share of eligible pull requests receiving a response within 7 days; minimum 10.'],
            ['m10 Issue close within 90d', 'Share of eligible issues closed within 90 days; minimum 10.'],
            ['m11 PR decision within 30d', 'Share of eligible pull requests decided within 30 days; minimum 10.'],
            ['m12 Stars gained', '100*min(1, log10(1+s)/3); s capped at 1,000.'],
            ['m13 Forks created', '100*min(1, log10(1+f)/2); f capped at 100.'],
        ], [2450, 6550]),
        para('Implied dimension shares', 'Heading2'),
        table(['Dimension', 'Metrics', 'Equal-metric share'], [
            ['Continuity', 'm1-m4', '4/13 = 30.77%'],
            ['Contributors', 'm5-m7', '3/13 = 23.08%'],
            ['Responsiveness', 'm8-m11', '4/13 = 30.77%'],
            ['Adoption', 'm12-m13', '2/13 = 15.38%'],
        ], [2450, 1750, 4800]),
        para('Coverage and interpretation', 'Heading2'),
        para('A Development Activity score requires at least 10 of 13 metrics (75%) and at least half of each dimension: 2/4 Continuity, 2/3 Contributors, 2/4 Responsiveness, and 1/2 Adoption. Bots are excluded and authorship uses human, non-merge commit rules. A complete-data zero is a measured zero under the defined anchor; unavailable values due to low activity volume or unsupported trackers are missing. Equal metric weights increase Adoption to 15.38% of the component from the prior weighted design\'s 5%, because stars and forks are two of the 13 indicators. Treat those as indirect attention/reuse signals, not measures of quality.'),
    ]))
    pages.append(page('5. Combined Results, Comparison, and Use', 'What the equal-metric perspective changes, and how to report it', [
        para('Combined-score reporting', 'Heading2'),
        para('Primary result: calculate the strict combined OISI only when all four equal-metric component scores qualify: (CE + Governance + Organizational Capacity + Development Activity)/4. Secondary all-112 result: the workbook also reports an evidence-adjusted score with unavailable component scores contributing zero. That zero is a conservative reporting convention, not observed zero performance. Always show the count of available components and retain component-level status.'),
        table(['Component / cohort', 'Equal-metric results'], [
            ['Eligible component scores', 'CE 88; Governance 108; Organizational Capacity 109; Development Activity 72.'],
            ['All four components eligible', '57 infrastructures; mean 74.54; range 46.90-91.52.'],
            ['All 112, missing components contribute zero', 'Mean 61.46; range 18.91-91.52.'],
            ['Weighted-vs-equal, strict cohort', '53 entries scoreable under both; Spearman rho .966; top-10 overlap 10/10; weighted mean 75.18, equal mean 73.88.'],
            ['Weighted-vs-equal, all 112', 'Spearman rho .9367; top-10 overlap 7/10; weighted mean 61.76, equal mean 61.46.'],
        ], [3200, 5800]),
        para('Component-level rank comparison on entries scoreable under both methods', 'Heading2'),
        table(['Component', 'n', 'Spearman rho', 'Top-10 overlap'], [
            ['Community Engagement', '83', '.9376', '8/10'],
            ['Governance Practices', '108', '.9741', '10/10'],
            ['Organizational Capacity', '108', '.9381', '9/10'],
            ['Development Activity', '72', '.9763', '8/10'],
        ], [3000, 1000, 2300, 2700]),
        para('Reading these results', 'Heading2'),
        para('The high strict-cohort rank correlation and 10/10 top-ten overlap indicate substantial agreement at the top of the combined ranking, while the score mean shifts from 75.18 to 73.88. Component results still move: equal shares particularly alter Organizational Capacity and Community Engagement scores. Correlation is rank agreement, not proof that either weighting is objectively correct. Report score differences, rank movement, and coverage rather than using correlation alone.'),
        para('Reviewer-ready justification', 'Heading2'),
        para('“We report equal metric weighting as a transparent alternative specification. Within each component, each scored and applicable indicator receives the same share; the four components remain equally weighted at 25%. This isolates the effect of within-component weighting while retaining the existing evidence definitions and coverage safeguards. We compare it with the prior weighted specification on common eligible records and publish component scores, missingness, and rank-stability statistics. Equal weighting is a design choice, not an empirically proven universal optimum.”'),
        para('Limitations and reproducibility', 'Heading2'),
        para('Equal weights do not eliminate judgment: metric selection, scale anchors, applicability, minimum coverage, and zero-coded blanks still matter. The measures have different reliability and construct scope; a metric-count average gives a short, indirect indicator the same nominal share as a broader measure. Repository activity omits work on other platforms and non-code contributions. For reproduction, use the component source sheets named in the workbook, the Equal Metric Scores sheet for per-metric values and component recalculations, Equal Metric Comparison for rank statistics, and Equal Metric Ranking / Equal Metric All 112 for results. Do not mix the strict complete-case ranking with the all-112 zero-filled ranking.'),
        para('References', 'Heading2'),
        para('[1] OECD, European Union, & European Commission Joint Research Centre. (2008). Handbook on Constructing Composite Indicators: Methodology and User Guide. OECD Publishing. https://doi.org/10.1787/9789264043466-en'),
        para('[2] Goggins, S., Lumbard, K., & Germonprez, M. (2021). Open Source Community Health: Analytical Metrics and Their Corresponding Narratives. IEEE/ACM SoHeal. https://doi.org/10.1109/SoHeal52568.2021.00010'),
        para('[3] Yin, L., Chakraborti, M., Yan, Y., Schweik, C., Frey, S., & Filkov, V. (2022). Open Source Software Sustainability: Combining Institutional Analysis and Socio-Technical Networks. Proceedings of the ACM on Human-Computer Interaction, 6(CSCW2), Article 404. https://doi.org/10.1145/3555129'),
        para('[4] CHAOSS Project. CHAOSS Metrics. https://chaoss.community/metrics/'),
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
        z.writestr('word/styles.xml', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/><w:sz w:val="18"/><w:color w:val="243746"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="75" w:line="215" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:after="80"/></w:pPr><w:rPr><w:rFonts w:ascii="Aptos Display" w:hAnsi="Aptos Display"/><w:b/><w:color w:val="1F4E78"/><w:sz w:val="27"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="65" w:after="40"/></w:pPr><w:rPr><w:b/><w:color w:val="1F4E78"/><w:sz w:val="20"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="130"/></w:pPr><w:rPr><w:color w:val="58717F"/><w:sz w:val="16"/></w:rPr></w:style><w:style w:type="table" w:styleId="GridTable"><w:name w:val="Grid Table"/></w:style></w:styles>''')
    print(OUT.resolve())


if __name__ == '__main__':
    main()
