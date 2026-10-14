from __future__ import annotations

from pathlib import Path
import re
from zipfile import ZIP_DEFLATED, ZipFile
from xml.sax.saxutils import escape

OUT = Path('OISI_Research_Methodology_EqualWeights_with_Formulas.docx')

SECTIONS = [
    ('Study design', 'The Open Infrastructure Sustainability Index (OISI) evaluates open infrastructures across Community Engagement, Governance Practices, Organizational Capacity, and Development Activity. It summarizes observable evidence; it does not directly predict survival, software quality, or social impact. The first three components use profile data associated with Zenodo record 18225266 [1]. The fourth uses repository and GitHub activity.'),
    ('Profile-data components', 'Community Engagement uses fields 037, 038, 938, 116, 040, 940, 041, 019, 042, and 043 for engagement, contribution, repository access, and community norms. Governance uses 024, 026, 108, 050, 048, 029, 033, 035, 044, and 046; field 100 is excluded from scoring due to sparse coverage. Organizational Capacity uses 079, 080, 056, 051, 058, 099, 076, and 077 for readiness, maintenance, institutional support, and service capacity. Funding and staffing are contextual. Raw answers and evidence links are retained; status responses use defined anchors and counts are capped.'),
    ('Development Activity', 'Repository resolution records each infrastructure’s repository, platform, branch, upstream, and issue tracker. Git history supplies commits, active weeks, recency, releases, contributor counts/distribution, and retention. R code queried GitHub GraphQL for issues, pull requests, and forks, and the REST star-history endpoint for gained stars. Paginated responses were filtered to the study period, retained as raw records, and summarized. Bots were excluded; identities were resolved within repositories. Available non-GitHub tracker evidence could be entered manually. Retrieval failures and insufficient samples remain missing, not zero.'),
    ('Window and metric selection', 'The observation window is 1 October 2025 inclusive to 1 October 2026 exclusive (UTC), a common recent annual cycle. Follow-up cutoffs are 14 days for issue response, 7 days for PR response, 90 days for issue closure, and 30 days for PR decisions. Metrics were selected for relevance, clear definitions, cohort-wide data feasibility, interpretability, and manageable redundancy. The 13 indicators measure continuity (commits, active weeks, recency, releases), contributors (count, distribution, retention), responsiveness (issue/PR response and disposition), and adoption (stars, forks). Response metrics require at least ten eligible items.'),
    ('Equal-weight scoring', 'Indicators are normalized to 0-100 and averaged equally: every CE metric contributes 14.29%; every Governance indicator, 10%; each applicable Organizational Capacity metric, 12.5% when all eight apply; and every Development Activity metric, 7.69%. Each component contributes 25% to OISI. Eligibility requires at least five of seven CE metrics across all groups; seven of ten Governance indicators including Formal and Transparency; 70% of applicable Organizational Capacity metrics across two dimensions; and ten of thirteen Development metrics plus at least half of each dimension. Not-applicable fields leave the denominator. Exact equation forms are shown in the formula tables below.'),
    ('Evaluation and interpretation', 'Checks reconcile the cohort, trace sources, recalculate scores, and inspect missingness, distributions, API diagnostics, and redundancy. Report complete cases separately from the all-entry view, where missing component scores contribute zero, not measured zero performance. Currently, 57 entries have all four scores (mean 74.54; range 46.90-91.52). Limitations include self-reporting, platform gaps, proxy measures, identity errors, and a one-year snapshot. Equal shares do not imply equal reliability [2-4].'),
    ('Continuing analysis', 'Alongside this equal-weight study, we are exploring a separate weighted analysis of the same infrastructure set for comparative analysis. It does not alter the present results.'),
    ('References', '[1] Zenodo record 18225266. https://zenodo.org/records/18225266. [2] OECD, EU & EC-JRC (2008), Handbook on Constructing Composite Indicators. https://doi.org/10.1787/9789264043466-en. [3] Goggins, Lumbard & Germonprez (2021), Open Source Community Health. https://doi.org/10.1109/SoHeal52568.2021.00010. [4] Yin et al. (2022), Open Source Software Sustainability, PACM HCI 6(CSCW2), Article 404. https://doi.org/10.1145/3555129.'),
]

METRIC_TABLES = {
    'Profile-data components': [
        ['Community Engagement', 'Fields', 'Measure / scoring interpretation', 'Equal share'],
        ['E1 Engagement availability', '037', 'Availability status; use the component status scale.', '14.29%'],
        ['E2 Engagement activity breadth', '038, 938', 'Distinct engagement activity categories; capped at 8.', '14.29%'],
        ['E3 Engagement evidence strength', '116, 938', 'Evidence URL or substantive description under current evidence rule.', '14.29%'],
        ['C1 Contribution opportunity breadth', '040, 940', 'Distinct contribution opportunity types; capped at 6.', '14.29%'],
        ['C2 Contribution guidance', '041', 'Contribution guide or relevant participation guidance.', '14.29%'],
        ['C3 Open contribution infrastructure', '019', 'Public repository/forge evidence.', '14.29%'],
        ['N1 Community norms / Code of Conduct', '042, 043', 'Status and URL evidence; current implemented-without-URL adjustment retained.', '14.29%'],
        ['Governance Practices', 'Fields', 'Measure / scoring interpretation', 'Equal share'],
        ['Governance structure', '024', 'Documented governance structure; status score.', '10%'],
        ['Governance records', '026', 'Public governance or decision records; status score.', '10%'],
        ['Bylaws', '108', 'Bylaws evidence; status score.', '10%'],
        ['Board structure', '050', 'Board/steering body presence as coded in source workbook.', '10%'],
        ['Community governance', '048', 'Formal=1; Ad hoc=0.5; None=0; blank follows source coding.', '10%'],
        ['Web accessibility', '029', 'Accessibility commitment/status.', '10%'],
        ['Transparent pricing', '033', 'Pricing transparency; current no-direct-costs mapping retained.', '10%'],
        ['Equity commitment', '035', 'Public equity commitment/status.', '10%'],
        ['Privacy commitment', '044', 'Public privacy commitment/status.', '10%'],
        ['Values frameworks', '046', 'Named framework presence as coded in source workbook.', '10%'],
        ['Organizational Capacity', 'Fields', 'Measure / scoring interpretation', 'Equal share'],
        ['M1 Technology readiness', '079', 'Readiness level under component scoring anchors.', '1/k; 12.5% if k=8'],
        ['M2 Maintenance status', '080', 'Maintenance status under component scoring anchors.', '1/k; 12.5% if k=8'],
        ['M3 Infrastructure age', '056', 'Capped age/tenure score under component rule.', '1/k; 12.5% if k=8'],
        ['I1 Formal organizational host', '051', 'Organizational host type under component rule.', '1/k; 12.5% if k=8'],
        ['I2 Affiliations / institutional support', '058', 'Affiliation/support evidence.', '1/k; 12.5% if k=8'],
        ['I3 SCOSS participation', '099', 'Participation status under component rule.', '1/k; 12.5% if k=8'],
        ['O1 Hosted/SaaS availability', '076', 'Service availability; exclude if structurally Not applicable.', '1/k; 12.5% if k=8'],
        ['O2 Service-provider ecosystem', '077', 'Registered provider evidence; applicable to specified third-party delivery.', '1/k; 12.5% if k=8'],
    ],
    'Development Activity': [
        ['Metric', 'Raw measure', 'Normalized scoring rule', 'Equal share'],
        ['m1 Human commits', 'Non-merge human commits c', '{{MATH:dev_m1}}; cap 1,000.', '7.69%'],
        ['m2 Active weeks', 'Distinct active weeks', '{{MATH:dev_m2}}.', '7.69%'],
        ['m3 Recency', 'Days since last qualifying commit d', '{{MATH:dev_m3}}.', '7.69%'],
        ['m4 Release activity', 'Git tags in window t', '{{MATH:dev_m4}}.', '7.69%'],
        ['m5 Active contributors', 'Distinct resolved human contributors n', '0 if n<=1; otherwise {{MATH:dev_m5}}; cap 32.', '7.69%'],
        ['m6 Effective contributors', '{{MATH:dev_eff_n}}', 'If E<=1, score=0; otherwise {{MATH:dev_m6}}; cap E=8.', '7.69%'],
        ['m7 Contributor retention', 'Contributors active in both half-windows r', '{{MATH:dev_m7}}.', '7.69%'],
        ['m8 Issue response', 'Issues responded to within 14 days', '{{MATH:dev_m8}}; minimum 10 eligible issues.', '7.69%'],
        ['m9 PR response', 'PRs responded to within 7 days', '{{MATH:dev_m9}}; minimum 10 eligible PRs.', '7.69%'],
        ['m10 Issue closure', 'Issues closed within 90 days', '{{MATH:dev_m10}}; minimum 10 eligible issues.', '7.69%'],
        ['m11 PR decision', 'PRs merged/closed within 30 days', '{{MATH:dev_m11}}; minimum 10 eligible PRs.', '7.69%'],
        ['m12 Stars gained', 'Stars added during window s', '{{MATH:dev_m12}}; cap 1,000.', '7.69%'],
        ['m13 Forks created', 'Forks created during window f', '{{MATH:dev_m13}}; cap 100.', '7.69%'],
    ],
}

FORMULA_TABLE = [
    ['Component', 'Metric(s)', 'Normalization formula / rule', 'Missing-data note'],
    ['Community Engagement', 'E1', '{{MATH:status_scale}}', 'Unknown/blank = NA.'],
    ['Community Engagement', 'E2', '{{MATH:ce_e2}}.', 'No recognized categories = NA.'],
    ['Community Engagement', 'E3', '{{MATH:ce_e3}}.', 'Evidence score.'],
    ['Community Engagement', 'C1', '{{MATH:ce_c1}}.', 'No opportunities = NA.'],
    ['Community Engagement', 'C2', '{{MATH:guide_rule}}', 'Blank/suspect URL = NA.'],
    ['Community Engagement', 'C3', '{{MATH:repo_rule}}', 'Blank URL = NA.'],
    ['Community Engagement', 'N1', '{{MATH:coc_rule}}', 'Unknown/blank = NA.'],
    ['Governance Practices', '024, 026, 108, 029, 035, 044', '{{MATH:status_scale}}', 'Unknown/NA excluded.'],
    ['Governance Practices', '033', '{{MATH:pricing_rule}}', 'Unknown/NA excluded.'],
    ['Governance Practices', '050, 046', '{{MATH:presence_rule}}.', 'Blank is not verified absence.'],
    ['Governance Practices', '048', '{{MATH:community_gov_rule}}.', 'Blank-as-zero is a coding assumption.'],
    ['Organizational Capacity', 'M1 Readiness', '{{MATH:org_m1}}.', 'Other/unknown = NA.'],
    ['Organizational Capacity', 'M2 Maintenance', '{{MATH:org_m2}}.', 'Unknown/blank = NA.'],
    ['Organizational Capacity', 'M3 Age', '{{MATH:org_m3}}.', 'Missing/unparseable = NA.'],
    ['Organizational Capacity', 'I1 Host', '{{MATH:org_i1}}.', 'Other/blank = NA.'],
    ['Organizational Capacity', 'I2 Affiliations', 'Named support=1; explicit no affiliation=0.', 'Blank = NA.'],
    ['Organizational Capacity', 'I3 SCOSS', '{{MATH:boolean_score}}.', 'Unknown/blank = NA.'],
    ['Organizational Capacity', 'O1 SaaS', '{{MATH:org_o1}}.', 'Not applicable excluded; unknown=NA.'],
    ['Organizational Capacity', 'O2 Providers', '{{MATH:org_o2}}.', 'Otherwise not applicable.'],
]

AGGREGATION_TABLE = [
    ['Calculation', 'Formula', 'Rule'],
    ['Common metric scale', '{{MATH:scale_conversion}}', 'Profile anchors are 0-1; Development Activity anchors are already 0-100.'],
    ['Equal component score', '{{MATH:component_mean}}', 'Publish only when that component’s coverage and dimension rules pass.'],
    ['Strict OISI', '{{MATH:oisi_strict}}', 'All four component scores qualify; each contributes 25%.'],
    ['All-entry evidence view', '{{MATH:oisi_missing_zero}}', 'Secondary coverage view, not measured zero performance.'],
    ['Illustrative example', '{{MATH:example_oisi}}', 'Example values only; not a repository result.'],
]


def run(text: str, bold: bool = False) -> str:
    props = '<w:rPr><w:b/></w:rPr>' if bold else ''
    return f'<w:r>{props}<w:t xml:space="preserve">{escape(text)}</w:t></w:r>'


def mr(text: str) -> str:
    return f'<m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t xml:space="preserve">{escape(text)}</m:t></m:r>'


def mfrac(numerator: str, denominator: str) -> str:
    return f'<m:f><m:fPr><m:type m:val="bar"/><m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr></m:ctrlPr></m:fPr><m:num>{numerator}</m:num><m:den>{denominator}</m:den></m:f>'


def msub(base: str, sub: str) -> str:
    return f'<m:sSub><m:e>{mr(base)}</m:e><m:sub>{mr(sub)}</m:sub></m:sSub>'


def msup(base: str, sup: str) -> str:
    return f'<m:sSup><m:e>{mr(base)}</m:e><m:sup>{mr(sup)}</m:sup></m:sSup>'


def msum(index: str, expression: str) -> str:
    return f'<m:nary><m:naryPr><m:chr m:val="∑"/><m:limLoc m:val="undOvr"/></m:naryPr><m:sub>{mr(index)}</m:sub><m:sup>{mr("")}</m:sup><m:e>{expression}</m:e></m:nary>'


def omath(*parts: str) -> str:
    return '<m:oMath>' + ''.join(parts) + '</m:oMath>'


MATH = {
    'status_scale': omath(mr('score = 1 if implemented; 0.67 if in progress; 0.33 if considering; 0 if not planning')),
    'ce_e2': omath(mr('E2 = '), mfrac(mr('min(distinct recognized categories, 8)'), mr('8'))),
    'ce_e3': omath(mr('E3 = 1 if URL; 0.5 if substantive 938 text; NA otherwise')),
    'ce_c1': omath(mr('C1 = '), mfrac(mr('min(distinct opportunities, 6)'), mr('6'))),
    'guide_rule': omath(mr('C2 = 1 for contribution guide; 0.5 for other relevant page')),
    'repo_rule': omath(mr('C3 = 1 for public forge; 0.5 for other repository URL')),
    'coc_rule': omath(mr('N1 = status score; 0.75 if implemented without URL')),
    'pricing_rule': omath(mr('s033 = status score; 1 for “No direct costs”')),
    'presence_rule': omath(mr('s050, s046 = 1 if present; 0 if blank under current coding')),
    'community_gov_rule': omath(mr('s048 = 1 (Formal); 0.5 (Ad hoc); 0 (None or blank under current coding)')),
    'org_m1': omath(msub('M', '1'), mr(' = 1 (operational proven); 8/9 (complete); 5/9 (relevant validation)')),
    'org_m2': omath(msub('M', '2'), mr(' = 1 (active); 0.5 (minimal); 0 (inactive/not maintained)')),
    'org_m3': omath(msub('M', '3'), mr(' = '), mfrac(mr('min(max(0, 2026 - launch year), 15)'), mr('15'))),
    'org_i1': omath(msub('I', '1'), mr(' = 1 (formal host); 0.5 (volunteer community)')),
    'org_i2': omath(msub('I', '2'), mr(' = 1 (named support); 0 (explicit no affiliation)')),
    'boolean_score': omath(mr('I3 = 1 if TRUE; 0 if FALSE')),
    'org_o1': omath(mr('O1 = 1 if provider service; 0 if no service; NA if not applicable/unknown')),
    'org_o2': omath(mr('O2 = 1 if provider listed; 0 if none, when third-party delivery applies')),
    'dev_m1': omath(msub('s', '1'), mr(' = 100 × min(1, '), mfrac(mr('log₁₀(1+c)'), mr('log₁₀(1001)')), mr(')')),
    'dev_m2': omath(msub('s', '2'), mr(' = 100 × '), mfrac(mr('active weeks'), mr('52'))),
    'dev_m5': omath(msub('s', '5'), mr(' = 100 × min(1, '), mfrac(mr('log₂(n)'), mr('5')), mr(')')),
    'dev_eff_n': omath(mr('E = '), mfrac(mr('1'), msum('i', msup(msub('p', 'i'), '2')))),
    'dev_m6': omath(msub('s', '6'), mr(' = 100 × min(1, '), mfrac(mr('log₂(E)'), mr('3')), mr(')')),
    'dev_m3': omath(msub('s', '3'), mr(' = 100 if d ≤ 30; otherwise max(0, '), mfrac(mr('365 - d'), mr('335')), mr(' × 100)')),
    'dev_m4': omath(msub('s', '4'), mr(' = 0 if t=0; 50 if t=1; 75 if 2≤t≤3; 100 if t≥4')),
    'dev_m7': omath(msub('s', '7'), mr(' = 0 if r=0; 25 if r=1; 50 if r=2; 75 if 3≤r≤4; 100 if r≥5')),
    'dev_m8': omath(msub('s', '8'), mr(' = '), mfrac(mr('issues responded within 14 days'), mr('eligible issues')), mr(' × 100')),
    'dev_m9': omath(msub('s', '9'), mr(' = '), mfrac(mr('PRs responded within 7 days'), mr('eligible PRs')), mr(' × 100')),
    'dev_m10': omath(msub('s', '10'), mr(' = '), mfrac(mr('issues closed within 90 days'), mr('eligible issues')), mr(' × 100')),
    'dev_m11': omath(msub('s', '11'), mr(' = '), mfrac(mr('PRs decided within 30 days'), mr('eligible PRs')), mr(' × 100')),
    'dev_m12': omath(msub('s', '12'), mr(' = 100 × min(1, '), mfrac(mr('log₁₀(1+s)'), mr('3')), mr(')')),
    'dev_m13': omath(msub('s', '13'), mr(' = 100 × min(1, '), mfrac(mr('log₁₀(1+f)'), mr('2')), mr(')')),
    'scale_conversion': omath(mr('metric points = 100 × rᵢ; rᵢ ∈ [0,1]')),
    'component_mean': omath(msub('C', 'j'), mr(' = '), mfrac(msum('i', msub('x', 'i')), mr('nⱼ'))),
    'oisi_strict': omath(mr('OISI = '), mfrac(mr('CE + Gov + Org + Dev'), mr('4'))),
    'oisi_missing_zero': omath(mr('OISI₄ = '), mfrac(mr('CE + Gov + Org + Dev; missing component = 0'), mr('4'))),
    'example_oisi': omath(mr('OISI = '), mfrac(mr('80 + 70 + 60 + 90'), mr('4')), mr(' = 75.00')),
    'share_7': omath(mfrac(mr('1'), mr('7'))),
    'share_10': omath(mfrac(mr('1'), mr('10'))),
    'share_k': omath(mfrac(mr('1'), mr('k'))),
    'share_13': omath(mfrac(mr('1'), mr('13'))),
    'elig_ce': omath(mfrac(mr('5'), mr('7'))),
    'elig_gov': omath(mfrac(mr('7'), mr('10'))),
    'elig_dev': omath(mfrac(mr('10'), mr('13'))),
}


def table_content(value: str) -> str:
    parts = re.split(r'(\{\{MATH:[a-z0-9_]+\}\})', str(value))
    rendered = []
    for part in parts:
        match = re.fullmatch(r'\{\{MATH:([a-z0-9_]+)\}\}', part)
        if match:
            rendered.append(MATH[match.group(1)])
        elif part:
            rendered.append(run(part))
    return ''.join(rendered)


def make_table(rows: list[list[str]]) -> str:
    widths = [2150, 1450, 3550, 1850] if len(rows[0]) == 4 else [2400, 3900, 2700]
    out = ['<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="9000" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="65" w:type="dxa"/><w:left w:w="90" w:type="dxa"/><w:bottom w:w="65" w:type="dxa"/><w:right w:w="90" w:type="dxa"/></w:tblCellMar><w:tblBorders><w:top w:val="single" w:sz="5" w:color="B7C9D6"/><w:left w:val="single" w:sz="5" w:color="B7C9D6"/><w:bottom w:val="single" w:sz="5" w:color="B7C9D6"/><w:right w:val="single" w:sz="5" w:color="B7C9D6"/><w:insideH w:val="single" w:sz="4" w:color="D7E0E7"/><w:insideV w:val="single" w:sz="4" w:color="D7E0E7"/></w:tblBorders></w:tblPr>']
    for ri, row in enumerate(rows):
        header = ri == 0 or row[1] == 'Fields'
        fill = '<w:shd w:fill="1F4E78"/>' if ri == 0 else ('<w:shd w:fill="DCE6F1"/>' if header else '')
        for_cell = []
        for ci, value in enumerate(row):
            color = 'FFFFFF' if ri == 0 else '243746'
            bold = '<w:b/>' if header else ''
            for_cell.append(f'<w:tc><w:tcPr><w:tcW w:w="{widths[ci]}" w:type="dxa"/>{fill}<w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:spacing w:after="45" w:before="35" w:line="220" w:lineRule="auto"/></w:pPr>{table_content(value)}</w:p></w:tc>')
        row_props = '<w:trPr><w:cantSplit/></w:trPr>' + ('<w:tblHeader w:val="true"/>' if ri == 0 else '')
        out.extend([f'<w:tr>{row_props}', *for_cell, '</w:tr>'])
    out.append('</w:tbl>')
    return ''.join(out)


def make_xml() -> str:
    parts = ['<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr>', run('OISI Research Methodology | Equal Weighting and Formulas'), '</w:p>', '<w:p><w:pPr><w:pStyle w:val="Subtitle"/></w:pPr>', run('Metric definitions, normalization, and index calculations for the study cohort'), '</w:p>']
    for heading, body in SECTIONS:
        parts.extend([f'<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>{run(heading)}</w:p>', f'<w:p><w:pPr><w:spacing w:after="100" w:line="270" w:lineRule="auto"/></w:pPr>{table_content(body)}</w:p>'])
        if heading in METRIC_TABLES:
            parts.append(make_table(METRIC_TABLES[heading]))
            if heading == 'Profile-data components':
                parts.extend(['<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr>', run('Metric normalization formulas'), '</w:p>', make_table(FORMULA_TABLE)])
        if heading == 'Equal-weight scoring':
            parts.extend(['<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr>', run('Final component and index calculations'), '</w:p>', make_table(AGGREGATION_TABLE)])
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><w:body>'
            + ''.join(parts) + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="850" w:right="900" w:bottom="850" w:left="900"/></w:sectPr></w:body></w:document>')


def main() -> None:
    with ZipFile(OUT, 'w', ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>''')
        z.writestr('_rels/.rels', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>''')
        z.writestr('word/document.xml', make_xml())
        z.writestr('word/styles.xml', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/><w:sz w:val="20"/><w:color w:val="243746"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="100" w:line="270" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:rPr><w:rFonts w:ascii="Aptos Display" w:hAnsi="Aptos Display"/><w:b/><w:color w:val="1F4E78"/><w:sz w:val="32"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:rPr><w:color w:val="58717F"/><w:sz w:val="18"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="180" w:after="90"/></w:pPr><w:rPr><w:b/><w:color w:val="1F4E78"/><w:sz w:val="23"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="150" w:after="70"/></w:pPr><w:rPr><w:b/><w:color w:val="1F4E78"/><w:sz w:val="20"/></w:rPr></w:style><w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/></w:style></w:styles>''')
    text = ' '.join(body for _, body in SECTIONS)
    print(OUT.resolve())
    print('Body word count:', len(text.split()))


if __name__ == '__main__':
    main()
