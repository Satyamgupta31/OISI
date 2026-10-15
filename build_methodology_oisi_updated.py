from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
import xml.etree.ElementTree as ET


ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "Methodology_OISI_Updated_Final.docx"

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ET.register_namespace("w", W)
ET.register_namespace("m", M)
ET.register_namespace("r", R)


def el(ns, name, attrs=None):
    return ET.Element(f"{{{ns}}}{name}", attrs or {})


def sub(parent, ns, name, attrs=None):
    return ET.SubElement(parent, f"{{{ns}}}{name}", attrs or {})


def run(text, bold=False, italic=False, size=None):
    r = el(W, "r")
    props = sub(r, W, "rPr")
    if bold:
        sub(props, W, "b")
    if italic:
        sub(props, W, "i")
    if size:
        sub(props, W, "sz", {f"{{{W}}}val": str(size)})
    t = sub(r, W, "t")
    if text[:1].isspace() or text[-1:].isspace():
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    return r


def paragraph(text="", style=None, bold=False, italic=False, align=None, after=100):
    p = el(W, "p")
    ppr = sub(p, W, "pPr")
    if style:
        sub(ppr, W, "pStyle", {f"{{{W}}}val": style})
    if align:
        sub(ppr, W, "jc", {f"{{{W}}}val": align})
    spacing = sub(ppr, W, "spacing", {f"{{{W}}}after": str(after), f"{{{W}}}line": "276", f"{{{W}}}lineRule": "auto"})
    if not text:
        return p
    p.append(run(text, bold=bold, italic=italic))
    return p


def mr(text):
    r = el(M, "r")
    t = sub(r, M, "t")
    t.text = text
    return r


def mfrac(numerator, denominator):
    f = el(M, "f")
    sub(f, M, "fPr")
    num = sub(f, M, "num")
    den = sub(f, M, "den")
    for part in numerator if isinstance(numerator, list) else [numerator]:
        num.append(part if isinstance(part, ET.Element) else mr(part))
    for part in denominator if isinstance(denominator, list) else [denominator]:
        den.append(part if isinstance(part, ET.Element) else mr(part))
    return f


def msub(base, index):
    s = el(M, "sSub")
    sub(s, M, "sSubPr")
    sub(s, M, "e").append(mr(base))
    sub(s, M, "sub").append(mr(index))
    return s


def msup(base, exponent):
    s = el(M, "sSup")
    sub(s, M, "sSupPr")
    sub(s, M, "e").append(mr(base))
    sub(s, M, "sup").append(mr(exponent))
    return s


def equation(parts):
    p = el(W, "p")
    ppr = sub(p, W, "pPr")
    sub(ppr, W, "jc", {f"{{{W}}}val": "center"})
    omathpara = sub(p, M, "oMathPara")
    omath = sub(omathpara, M, "oMath")
    for part in parts:
        omath.append(part if isinstance(part, ET.Element) else mr(part))
    return p


def cell(text, header=False, width=3000):
    c = el(W, "tc")
    tcpr = sub(c, W, "tcPr")
    sub(tcpr, W, "tcW", {f"{{{W}}}w": str(width), f"{{{W}}}type": "dxa"})
    p = sub(c, W, "p")
    ppr = sub(p, W, "pPr")
    sub(ppr, W, "spacing", {f"{{{W}}}after": "40", f"{{{W}}}line": "230", f"{{{W}}}lineRule": "auto"})
    if header:
        rpr = sub(p, W, "rPr")
        sub(rpr, W, "b")
        sub(rpr, W, "color", {f"{{{W}}}val": "FFFFFF"})
    r = sub(p, W, "r")
    if header:
        rpr = sub(r, W, "rPr")
        sub(rpr, W, "b")
        sub(rpr, W, "color", {f"{{{W}}}val": "FFFFFF"})
    t = sub(r, W, "t")
    t.text = str(text)
    return c


def table(headers, rows, widths=None):
    n = len(headers)
    widths = widths or [9360 // n] * n
    tbl = el(W, "tbl")
    pr = sub(tbl, W, "tblPr")
    sub(pr, W, "tblW", {f"{{{W}}}w": "9360", f"{{{W}}}type": "dxa"})
    borders = sub(pr, W, "tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        sub(borders, W, edge, {f"{{{W}}}val": "single", f"{{{W}}}sz": "4", f"{{{W}}}color": "B7C9C2"})
    grid = sub(tbl, W, "tblGrid")
    for width in widths:
        sub(grid, W, "gridCol", {f"{{{W}}}w": str(width)})
    for i, values in enumerate([headers, *rows]):
        tr = sub(tbl, W, "tr")
        if i == 0:
            trpr = sub(tr, W, "trPr")
            sub(trpr, W, "tblHeader")
        for j, value in enumerate(values):
            c = cell(value, header=(i == 0), width=widths[j])
            if i == 0:
                shd = sub(c.find(f"{{{W}}}tcPr"), W, "shd", {f"{{{W}}}fill": "315E50", f"{{{W}}}val": "clear"})
            tr.append(c)
    return tbl


body = []


def add(text="", style=None, **kwargs):
    body.append(paragraph(text, style, **kwargs))


def heading(text, level=1):
    add(text, f"Heading{level}", after=140)


def eq(label, parts):
    add(label, style="Caption", after=30)
    body.append(equation(parts))


def tab(headers, rows, widths=None):
    body.append(table(headers, rows, widths))
    body.append(paragraph("", after=80))


add("Open Scholarly Infrastructure Index", "Title", align="center", after=40)
add("Data sources, metric construction, scoring and aggregation", "Subtitle", align="center", after=200)
add("Methodology document | Equal-weight specification", "Subtitle", align="center", after=260)
add("Abstract", "Heading1")
add("This methodology describes how the Open Scholarly Infrastructure Index (OISI) was assembled from the Infra Finder dataset and repository activity data; how four components and their indicators were selected; and how raw fields were normalized, scored and combined. Community Engagement, Governance Practices, Organizational Capacity and Development Activity each contribute 25% to the overall index. Metrics are equally weighted within each component. Missing values are zero-filled in the all-entry evidence-adjusted score, while observed-data indicators and complementary diagnostic views are retained. The cited literature motivates broad constructs and transparent composite-index design; the specific OISI formulas and thresholds are study operationalizations.")

heading("1. Study Design and Data Sources")
add("The unit of analysis is an infrastructure/service represented in the Infra Finder master set. The study population contains 112 entries. The first three components use the Infra Finder version 3.0 data package deposited by Invest in Open Infrastructure on Zenodo (record 18225266). The package includes structured data files and field definitions that support consistent interpretation of the source variables.")
add("We reviewed the dataset structure, field names, controlled response options and available records, then selected fields relevant to four complementary constructs: community participation, governance, organizational capacity and development activity. Source fields were mapped to metric codes; categorical responses were harmonized to the scoring categories, and narrative/free-text fields were coded using the component-specific rules. The field-to-metric mapping and scoring definitions are documented in the component workbooks.")
add("Development Activity was assembled separately from repository-level data. The repository list was resolved to public GitHub repositories where a match was available. A separate R-language script was used to collect GitHub API records for the defined observation window. The returned records were mapped to the repository register and processed under the stated metric definitions to produce the API-derived activity inputs. The R script and its extraction outputs should be retained with the study's reproducibility materials; the API-derived measures are limited to activity visible through GitHub.")
add("The observation window for time-dependent activity measures is 1 October 2025 00:00 UTC inclusive through 1 October 2026 00:00 UTC exclusive. This fixed 12-month window gives each repository the same calendar period. Measures requiring follow-up (issue and pull-request response or decision times) exclude records that do not have the full follow-up interval observable before the cutoff.")

heading("2. Data Processing and Metric Construction")
add("The master service names were normalized for joining across component outputs. Each component output was checked for duplicate or unmatched service keys before integration. Categorical values were mapped according to explicit scoring tables. For Community Engagement, fields 938 and 940 supplement their parent selections only when free text supports a distinct Other category; field 043 is evidence for Code of Conduct status rather than an additional metric. For Governance, field 950 supports classification of Other board descriptions. Narrative organizational-support data are coded consistently under the Organizational Capacity rubric.")
add("Development Activity separates local Git-history measures from GitHub API measures. The resulting repository data are aggregated to the infrastructure level under the repository mapping rules. GitHub observations are platform-specific; absence of API-visible events is not treated as evidence about activity on other platforms.")

heading("3. Metric Framework and Weights")
add("All component metrics are represented on a 0–100 scale. Within a component, each listed metric has equal weight. The component score is the arithmetic mean of its metric scores, using the fixed component denominator. The four component scores are then averaged equally, so each component contributes 25% regardless of its number of metrics. This hierarchical design prevents the 13-metric Development Activity component from dominating merely because it contains more indicators.")
tab(["Component", "Metrics", "Metric count", "Within-component weight", "OISI weight"], [
    ["Community Engagement", "E1, E2, C1, C2, C3, N1", "6", "1/6 each", "25%"],
    ["Governance Practices", "G1–G7", "7", "1/7 each", "25%"],
    ["Organizational Capacity", "M1, M2, H1, I2, O1", "5", "1/5 each", "25%"],
    ["Development Activity", "m1–m13", "13", "1/13 each", "25%"],
], [2400, 3100, 1100, 1500, 1260])

heading("4. Component Indicators and Scoring Rules")
heading("4.1 Community Engagement", 2)
add("The six metrics measure reported engagement status, breadth of engagement activities, breadth of contribution routes, contribution guidance, public repository listing and Code of Conduct status. The 0/33/67/100 status ladder assigns evenly spaced values to ordered response categories; it is a transparent scoring convention.")
tab(["Code and metric", "Source", "Scoring rule"], [
    ["E1 Engagement status", "037", "Implemented/available 100; in progress 67; considering 33; not planning 0. Unknown/blank unavailable, zero in all-entry view."],
    ["E2 Engagement activity breadth", "038 + coded 938", "Distinct supported activity categories scored against the 13-category maximum. Twelve named categories plus Other only when selected and supported by distinct 938 text; deduplicate. See equation below."],
    ["C1 Contribution opportunity breadth", "040 + coded 940", "Distinct supported routes scored against the nine-route maximum. Eight named routes plus Other only when selected and supported by distinct 940 text; exclude Not applicable; deduplicate. See equation below."],
    ["C2 Contribution guidance", "041", "URL containing contribut, get-involved, giveortake or newcomers scores 100 unless suspect. Another populated, non-suspect URL scores 50. Suspect patterns: discord, terms_of_use, account-types, bute_code or semicolon. Blank/suspect unavailable, zero-filled in all-entry view."],
    ["C3 Public contribution repository", "019", "URL containing github.com, gitlab.com, bitbucket.org, codeberg.org or sr.ht scores 100; other populated field scores 50; blank scores 0 in all-entry view. Host matching is not a live access test."],
    ["N1 Code of Conduct status", "042; 043 evidence only", "Implemented/available 100; in progress 67; considering 33; not planning 0. Unknown/blank unavailable, zero-filled in all-entry view."],
], [2400, 1600, 5360])
eq("Breadth metric equations", ["E2", " = 100 × ", mfrac("distinct supported activity types", "13"), "     C1 = 100 × ", mfrac("distinct supported contribution routes", "9")])
eq("Community Engagement component score", ["CE", " = ", mfrac("E1 + E2 + C1 + C2 + C3 + N1", "6")])
add("Field 116 is retained as evidence, not scored as a separate indicator. Field 048 belongs to Governance Practices, not Community Engagement. Fields 938, 940 and 043 support coding of other selected metrics rather than earning separate points.")

heading("4.2 Governance Practices", 2)
add("The seven indicators represent reported governance arrangements and commitments. POSI provides conceptual context for accountable organizational practices, stakeholder governance and inclusion. The selected fields measure reported presence or status, not the quality or effectiveness of governance.")
tab(["Code and metric", "Source", "Scoring rule"], [
    ["G1 Governance structure status", "024", "Implemented/available 100; in progress 67; considering 33; not planning 0; unknown/blank/unrecognized 0 in final zero-filled score."],
    ["G2 Bylaws or equivalent formal rules", "108", "Same four-level status ladder as G1."],
    ["G3 Board or steering-body presence", "050 + 950", "Recognized advisory/steering or multi-board structure 100; None 0; Other 100 only when 950 substantiates a relevant body, otherwise provisional 0 and review flag."],
    ["G4 Web accessibility status", "029", "Same four-level status ladder as G1."],
    ["G5 Equity commitment status", "035", "Same four-level status ladder as G1."],
    ["G6 Privacy commitment status", "044", "Same four-level status ladder as G1."],
    ["G7 Community governance formality", "048", "Formal 100; ad hoc 50; none 0; unknown/blank/unrecognized 0 in final zero-filled score."],
], [2400, 1600, 5360])
eq("Governance Practices component score", ["G", " = ", mfrac("G1 + G2 + G3 + G4 + G5 + G6 + G7", "7")])
add("Fields 026 (governance records) and 033 (transparent pricing) are not included in this final metric set. Field 050 is retained as G3; field 950 is supporting classification evidence, not an eighth metric. A reported board/body does not establish independence, representation or effectiveness.")

heading("4.3 Organizational Capacity", 2)
add("The five metrics represent technical readiness, maintenance, human resources, institutional affiliation/support and service delivery. NASA's nine-stage Technology Readiness Level framework informs M1. Staffing and affiliation scores are based on the source fields and explicit coding rules; each is one metric in the component mean.")
tab(["Code and metric", "Source", "Scoring rule"], [
    ["M1 Technology readiness", "079", "Normalize the recorded stage on the source nine-level scale; examples: levels 9, 8 and 5 yield 100, 88.89 and 55.56. Unknown/blank/unrecognized 0. See equation below."],
    ["M2 Maintenance status", "080", "Actively maintained 100; minimally maintained 50; inactive/not maintained 0; unknown/blank 0."],
    ["H1 Human-resource capacity", "054 + 055", "Score full-time and volunteer fields separately: 0 people 0; 1–5 25; 6–10 50; 11–20 75; >20 100. Missing subfield 0; average the two."],
    ["I2 Institutional affiliation/support", "058", "Formal host or explicit operational/financial support 100; substantive affiliation without stated operational support 50; explicit absence/placeholders/blank/unknown 0. Free-text coding is reviewable."],
    ["O1 Hosted/service-delivery capacity", "076", "Any reported hosted/service route 100; no services available 0; not applicable/unknown/blank 0. Measures reported availability, not service quality or reliability."],
], [2400, 1600, 5360])
eq("Human-resource capacity metric", ["H1", " = ", mfrac("Full-time staffing score + Volunteer staffing score", "2")])
eq("Technology readiness normalization", ["M1 = 100 × ", mfrac("readiness level", "9")])
eq("Organizational Capacity component score", ["OC", " = ", mfrac("M1 + M2 + H1 + I2 + O1", "5")])
add("Infrastructure age (M3) and formal organizational host (I1) are excluded from the final metric set. Age is not a direct measure of current capacity; formal-host status may overlap with institutional-support information. Other fields remain descriptive or are not sufficiently comparable under this component's operational definition.")

heading("4.4 Development Activity", 2)
add("The 13 equally weighted indicators summarize repository continuity, contributors, responsiveness and activity signals over the fixed 12-month window. CHAOSS project-health concepts and open-source sustainability research informed construct selection. Borges and Valente support interpreting stars as an attention/visibility signal, not software quality. The specific caps, bins, thresholds and windows below are study-defined operational rules.")
tab(["Code and metric", "Raw measure", "0–100 scoring rule"], [
    ["m1 Human commits", "c = qualifying non-merge, non-bot commits", "Log-scaled count; cap at 1,000 commits. See equation below."],
    ["m2 Active weeks", "w = distinct active weeks", "Proportion of 52 weeks with qualifying activity, expressed on 0–100. See equation below."],
    ["m3 Commit recency", "d = days since latest qualifying commit", "Full score for activity within 30 days, declining to zero at 365 days. No activity in a complete Git extract = 0. See equation below."],
    ["m4 Release activity", "Git tags in window", "0 tags 0; 1 tag 50; 2–3 tags 75; 4+ tags 100."],
    ["m5 Active contributors", "n = distinct qualifying human contributors", "0–1 contributors score 0; then a log-scaled score capped at 32 contributors. See equation below."],
    ["m6 Effective contributors", "E = inverse concentration of commit shares", "Concentration-adjusted contributor score; 0 if E ≤ 1, rising on a log scale and capped at E=8. See equation below."],
    ["m7 Contributor retention", "Contributors active in both half-windows", "0 contributors 0; 1 = 25; 2 = 50; 3–4 = 75; 5+ = 100."],
    ["m8 Issue response ≤14 days", "Eligible issues responded to within 14 days", "Share receiving a timely first response, only if n ≥ 10; otherwise unavailable and zero-filled."],
    ["m9 PR response ≤7 days", "Eligible PRs responded to within 7 days", "Share receiving a timely first response, only if n ≥ 10; otherwise unavailable and zero-filled."],
    ["m10 Issue closure ≤90 days", "Eligible issues closed within 90 days", "Share closed within 90 days, only if n ≥ 10; otherwise unavailable and zero-filled."],
    ["m11 PR decision ≤30 days", "Eligible PRs merged or closed within 30 days", "Share merged or closed within 30 days, only if n ≥ 10; measures decision, not acceptance."],
    ["m12 Stars gained", "s = stars gained in window", "Log-scaled attention signal; cap at 1,000 stars. See equation below."],
    ["m13 Forks created", "f = forks created in window", "Log-scaled potential reuse signal; cap at 100 forks. See equation below."],
], [2100, 3100, 4160])
eq("m1 human-commit score", ["m1 = 100 × min(1, ", mfrac([msub("log", "10"), "(1+c)"], [msub("log", "10"), "(1001)"]), ")"])
eq("m2 active-week score", ["m2 = 100 × ", mfrac("w", "52")])
eq("m3 commit-recency score", ["m3 = 100, d ≤ 30; otherwise max(0, 100 × ", mfrac("365 − d", "335"), "). If complete Git history has no qualifying activity, m3 = 0."])
eq("m5 active-contributor score", ["m5 = 0, n ≤ 1; otherwise 100 × min(1, ", mfrac([msub("log", "2"), "(n)"], "5"), "); cap at n = 32."])
eq("Effective contributor count", ["E = ", mfrac("1", "Σ pᵢ²"), "; pᵢ is contributor i's share of qualifying human commits."])
eq("m6 effective-contributor score", ["m6 = 0, E ≤ 1; otherwise 100 × min(1, ", mfrac([msub("log", "2"), "(E)"], "3"), "); cap at E = 8."])
eq("m8 issue-response rate", ["m8 = 100 × ", mfrac("eligible issues receiving a response within 14 days", "eligible issues"), "; scored as observed only when n ≥ 10."])
eq("m9 pull-request response rate", ["m9 = 100 × ", mfrac("eligible PRs receiving a response within 7 days", "eligible PRs"), "; scored as observed only when n ≥ 10."])
eq("m10 issue closure rate", ["m10 = 100 × ", mfrac("eligible issues closed within 90 days", "eligible issues"), "; scored as observed only when n ≥ 10."])
eq("m11 pull-request decision rate", ["m11 = 100 × ", mfrac("eligible PRs merged or closed within 30 days", "eligible PRs"), "; scored as observed only when n ≥ 10."])
eq("m12 stars-gained score", ["m12 = 100 × min(1, ", mfrac([msub("log", "10"), "(1+s)"], [msub("log", "10"), "(1001)"]), "); cap at s = 1,000."])
eq("m13 forks-created score", ["m13 = 100 × min(1, ", mfrac([msub("log", "10"), "(1+f)"], [msub("log", "10"), "(101)"]), "); cap at f = 100."])
eq("Development Activity component score", ["DA", " = ", mfrac("m1 + m2 + m3 + m4 + m5 + m6 + m7 + m8 + m9 + m10 + m11 + m12 + m13", "13")])
add("The minimum of 10 eligible items applies only to m8–m11 rates; it is not an overall score-eligibility rule. Eligible issues and pull requests must be authored by non-bots and have the full response or decision period observable before the window closes. For m8, the calculation first uses non-maintainer issues and falls back to all eligible issues if fewer than 10 are available. Qualifying issue response includes a non-author, non-bot comment, label, assignment or close event; qualifying PR response includes a non-author, non-bot comment/review or a merge/close by someone other than the author. m10 and m11 measure follow-through/decision under their definitions. Closed pull requests are not necessarily accepted. API-visible GitHub activity does not capture private activity or activity on other hosting platforms.")

heading("5. Aggregation, Missing Data and Ranking")
add("For the all-entry score, each metric is scored under its component rule and unavailable metrics are assigned zero. The component denominator remains fixed at the number of metrics in that component. Each component contributes one quarter of the overall index.")
eq("Overall equal-component index", ["OISI", " = ", mfrac("CE + G + OC + DA", "4"), " = 0.25CE + 0.25G + 0.25OC + 0.25DA"])
add("Three complementary views are reported. The all-entry zero-filled OISI includes all entries and is interpreted as evidence-adjusted because it reflects both measured scores and missing information. The complete-data comparison includes only records with all 31 metrics observed (six cases in the current integrated workbook), so it is a small-cohort diagnostic. Separately, the Community Engagement component has 49 complete cases across its six metrics. The available-component diagnostic averages observed metrics within each component, then averages components with at least one observed metric; because its denominator varies, it is descriptive rather than a directly comparable ranking.")
eq("Component observed coverage", ["Coverage", " = 100 × ", mfrac("number of observed metrics", "number of metrics in component")])
eq("Observed-only component mean", ["S", msub("obs", "component"), " = ", mfrac("sum of observed metric scores", "number of observed metrics")])
add("A measured zero, an unavailable metric that is zero-filled, and a missing component result are distinct states. The workbook retains observed flags and data-quality notes. There is no overall coverage eligibility gate and no 70-point performance band. Rank ties use average ranks.")

heading("6. Why These Metrics and Not Others?")
add("Metric selection balanced construct relevance, availability in the source data, interpretability and ability to score consistently across infrastructures. Community Engagement retains indicators of engagement status and breadth, contribution routes/guidance, repository listing and conduct status. Governance retains structure, formal rules, board/body presence, accessibility, equity, privacy and community governance. Organizational Capacity retains readiness, maintenance, workforce, institutional support and service delivery. Development Activity retains measures of activity volume and continuity, contributor distribution/retention, responsiveness/follow-through and repository attention or reuse signals.")
add("Fields that duplicate supporting evidence were not scored separately: 938/940 support Other coding; 043 supports N1; 950 supports G3. Fields 116 and 048 are not Community Engagement metrics; 048 is a Governance metric. Governance fields 026 and 033 were excluded from the final set. Organizational Capacity excludes M3 infrastructure age and I1 formal host, and leaves descriptive or conditional fields out of the comparable core score. Development data not selected as independent indicators were excluded when they duplicated or closely overlapped a retained measure. These exclusions are specific to the present construct and data design, not claims that the fields are unimportant in every study.")

heading("7. Interpretation")
add("OISI is a transparent, equal-weight composite of selected reported and observed indicators. It is not an externally validated direct measure of infrastructure quality, impact, security, user satisfaction or long-term sustainability. Equal weights are a design choice rather than empirically estimated importance. URL patterns, repository hosts, Git tags, stars and forks are proxies with construct-specific meaning. Interpret rankings alongside component scores, metric coverage and repository/platform context.")
add("Composite-indicator guidance motivates explicit normalization, weighting, aggregation and missing-data decisions and robustness checks. The present workbook compares complete-data, zero-filled and available-component views; it does not establish that the exact weights, thresholds or scoring cutoffs are uniquely correct.")

heading("References")
refs = [
    "Invest in Open Infrastructure. (2026). Data for Infra Finder, version 3.0. Zenodo record 18225266. https://zenodo.org/records/18225266",
    "Nardo, M., Saisana, M., Saltelli, A., Tarantola, S., Hoffmann, A., & Giovannini, E. (2008). Handbook on Constructing Composite Indicators: Methodology and User Guide. OECD Publishing / European Commission Joint Research Centre. https://doi.org/10.1787/9789264043466-en",
    "Saisana, M., Saltelli, A., & Tarantola, S. (2005). Uncertainty and sensitivity analysis techniques as tools for the quality assessment of composite indicators. Journal of the Royal Statistical Society: Series A, 168(2), 307–323. https://doi.org/10.1111/j.1467-985X.2005.00350.x",
    "Principles of Open Scholarly Infrastructure (POSI), version 1.0. https://doi.org/10.24343/C34W2H",
    "NASA. Technology Readiness Levels. https://www.nasa.gov/directorates/somd/space-communications-navigation-program/technology-readiness-levels/",
    "CHAOSS. Starter Project Health Metrics Model. https://www.chaoss.community/starter-project-health-metrics-model/",
    "Avelino, G., et al. (2019). On the abandonment and survival of open source projects: An empirical investigation. ESEM 2019. https://doi.org/10.1109/ESEM.2019.8870181",
    "Borges, H., & Valente, M. T. (2018). What’s in a GitHub Star? Understanding Repository Starring Practices in a Social Coding Platform. Journal of Systems and Software. https://doi.org/10.1016/j.jss.2018.09.016",
    "Kalliamvakou, E., et al. (2014). The Promises and Perils of Mining GitHub. Proceedings of the 11th Working Conference on Mining Software Repositories. https://doi.org/10.1145/2597073.2597074",
    "Yehudi, Y., Goble, C., & Jay, C. (2023). Individual context-free online community health indicators fail to identify open source software sustainability. arXiv:2309.12120. https://arxiv.org/abs/2309.12120",
]
for i, ref in enumerate(refs, 1):
    add(f"{i}. {ref}")


def make_styles():
    styles = el(W, "styles")
    default = sub(styles, W, "docDefaults")
    rpr_default = sub(default, W, "rPrDefault")
    rpr = sub(rpr_default, W, "rPr")
    sub(rpr, W, "rFonts", {f"{{{W}}}ascii": "Aptos", f"{{{W}}}hAnsi": "Aptos", f"{{{W}}}eastAsia": "Aptos"})
    sub(rpr, W, "sz", {f"{{{W}}}val": "21"})
    ppr_default = sub(default, W, "pPrDefault")
    sub(ppr_default, W, "pPr")
    for sid, name, size, color in [
        ("Normal", "Normal", 21, "24312D"),
        ("Title", "Title", 36, "244D42"),
        ("Subtitle", "Subtitle", 22, "567067"),
        ("Heading1", "Heading 1", 28, "244D42"),
        ("Heading2", "Heading 2", 23, "315E50"),
        ("Caption", "Caption", 18, "567067"),
    ]:
        s = sub(styles, W, "style", {f"{{{W}}}type": "paragraph", f"{{{W}}}styleId": sid})
        sub(s, W, "name", {f"{{{W}}}val": name})
        rp = sub(s, W, "rPr")
        sub(rp, W, "sz", {f"{{{W}}}val": str(size)})
        sub(rp, W, "color", {f"{{{W}}}val": color})
        if sid.startswith("Heading") or sid == "Title":
            sub(rp, W, "b")
        if sid.startswith("Heading"):
            pp = sub(s, W, "pPr")
            sub(pp, W, "keepNext")
    return styles


def write_docx():
    doc = el(W, "document")
    b = sub(doc, W, "body")
    for item in body:
        b.append(item)
    sect = sub(b, W, "sectPr")
    sub(sect, W, "pgSz", {f"{{{W}}}w": "12240", f"{{{W}}}h": "15840"})
    sub(sect, W, "pgMar", {f"{{{W}}}top": "900", f"{{{W}}}right": "1440", f"{{{W}}}bottom": "900", f"{{{W}}}left": "1440", f"{{{W}}}header": "420", f"{{{W}}}footer": "420", f"{{{W}}}gutter": "0"})
    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>'''
    rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'''
    docrels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'''
    core = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>OISI Research Methodology</dc:title><dc:subject>Data sources, metric construction, scoring and aggregation</dc:subject><dc:creator>OISI Research Team</dc:creator></cp:coreProperties>'''
    with ZipFile(OUTPUT, "w", ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/_rels/document.xml.rels", docrels)
        z.writestr("word/document.xml", ET.tostring(doc, encoding="utf-8", xml_declaration=True))
        z.writestr("word/styles.xml", ET.tostring(make_styles(), encoding="utf-8", xml_declaration=True))
        z.writestr("docProps/core.xml", core)
    print(f"Created {OUTPUT}")
    print(f"OMML equations: {len(doc.findall('.//{%s}oMathPara' % M))}")


if __name__ == "__main__":
    write_docx()
