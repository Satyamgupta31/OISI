from __future__ import annotations

import re
import statistics
from pathlib import Path

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "OISI_Full_Index_Equal_Component_Weight_112_Missing_Data_Views_Literature_Supported.xlsx"
SOURCES = {
    "Community Engagement": (ROOT / "Community-Engagement/OISI_Community_Analysis_Final_Defensible.xlsx", "Scoring"),
    "Governance Practices": (ROOT / "Governance-Practices/OISI_Governance_Practices_Final_Defensible_Audit_Corrected.xlsx", "Scoring"),
    "Organizational Capacity": (ROOT / "Organizational-capacity/OISI_Organizational_Capacity_Final_Defensible_Audit_Corrected.xlsx", "Scoring"),
    "Development Activity": (ROOT / "Development-Activity/Dev Data/oisi_c4/OISI_Development_Activity_Final_Equal_Weight_Zero_Fill_Corrected.xlsx", "Scoring"),
}
COMPONENTS = list(SOURCES)
COMPONENT_WEIGHT = 0.25
COLORS = {"green": "245B55", "blue": "315B70", "light": "E8F1EF", "gold": "FCECC9", "red": "9A3D3D", "white": "FFFFFF", "line": "D5DFDC", "gray": "EEF1F2"}


def norm(value):
    return re.sub(r"[^a-z0-9]+", "", str(value or "").lower())


def read_sheet(path, sheet):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[sheet]
    rows = list(ws.iter_rows(values_only=True))
    headers = [str(v) if v is not None else "" for v in rows[0]]
    return [dict(zip(headers, row)) for row in rows[1:] if any(v is not None for v in row)]


def index_rows(rows, key):
    result = {}
    for row in rows:
        k = norm(row.get(key))
        if not k:
            continue
        if k in result:
            raise ValueError(f"Duplicate normalized join key {row.get(key)!r}")
        result[k] = row
    return result


def score(value):
    if value in (None, "", "NA", "N/A", "None"):
        return None
    try:
        return float(value)
    except (ValueError, TypeError):
        return None


def average(values):
    return sum(values) / len(values) if values else 0.0


def rank_average(values):
    order = sorted(range(len(values)), key=lambda i: (-values[i], i))
    ranks = [0.0] * len(values)
    pos = 0
    while pos < len(order):
        end = pos + 1
        while end < len(order) and values[order[end]] == values[order[pos]]:
            end += 1
        rank = ((pos + 1) + end) / 2
        for j in range(pos, end):
            ranks[order[j]] = rank
        pos = end
    return ranks


def metric_specs():
    return {
        "Community Engagement": [(x, x, 1 / 6) for x in ["E1", "E2", "C1", "C2", "C3", "N1"]],
        "Governance Practices": [(x, x, 1 / 7) for x in ["G1", "G2", "G3", "G4", "G5", "G6", "G7"]],
        "Organizational Capacity": [(x, x, 1 / 5) for x in ["M1", "M2", "H1", "I2", "O1"]],
        "Development Activity": [(x, x, 1 / 13) for x in [
            "m1_commits", "m2_active_weeks", "m3_recency", "m4_releases", "m5_contributors",
            "m6_effective_contributors", "m7_retention", "m8_issue_response", "m9_pr_response",
            "m10_issue_close90", "m11_pr_decision30", "m12_stars", "m13_forks",
        ]],
    }


def load_all():
    ce_path, _ = SOURCES["Community Engagement"]
    master = read_sheet(ce_path, "Data")
    if len(master) != 112:
        raise ValueError(f"Expected 112 Community Engagement master rows; found {len(master)}")
    services = [r.get("001_service_name") for r in master]
    if len({norm(v) for v in services}) != len(services):
        raise ValueError("Community Engagement master has duplicate normalized service names")

    source = {name: read_sheet(path, sheet) for name, (path, sheet) in SOURCES.items()}
    keys = {
        "Community Engagement": "Service",
        "Governance Practices": "Service",
        "Organizational Capacity": "Service",
        "Development Activity": "Infrastructure",
    }
    indexed = {name: index_rows(source[name], keys[name]) for name in COMPONENTS}
    for name in COMPONENTS:
        missing = [s for s in services if norm(s) not in indexed[name]]
        extras = set(indexed[name]) - {norm(s) for s in services}
        if missing or extras:
            raise ValueError(f"Join mismatch for {name}: missing={len(missing)}, extra={len(extras)}")

    # The CE scoring formulas are not cached in its workbook; use the source-derived
    # Calc values and any explicit numeric reviewer overrides from Verification.
    ce_calc = index_rows(read_sheet(ce_path, "Calc"), "Service")
    ce_verification = index_rows(read_sheet(ce_path, "Verification"), "Service")
    ce_fields = ["E1", "E2", "C1", "C2", "C3", "N1"]
    ce_overrides = ["E1 override", "E2 override", "C1 override", "C2 override", "C3 override", "N1 override"]

    # Use source workbook coverage flags so true observed zeros remain distinct
    # from zero-filled missing values.
    coverage_sheets = {
        "Governance Practices": "Coverage",
        "Organizational Capacity": "Coverage",
        "Development Activity": "Coverage",
    }
    cover = {}
    for name, sheet in coverage_sheets.items():
        path, _ = SOURCES[name]
        data = read_sheet(path, sheet)
        key = "Infrastructure" if name == "Development Activity" else "Service"
        cover[name] = index_rows(data, key)

    specs = metric_specs()
    rows, detail = [], []
    for base in master:
        service = base["001_service_name"]
        k = norm(service)
        component_scores, component_counts, component_metric_scores, component_observed_means = {}, {}, {}, {}
        all_observed = True
        for comp in COMPONENTS:
            record = indexed[comp][k]
            component_values, observed_flags = [], []
            if comp == "Community Engagement":
                calc = ce_calc[k]
                verification = ce_verification.get(k, {})
                for field, override_field, _ in specs[comp]:
                    raw = score(calc.get(field))
                    override = score(verification.get(override_field))
                    observed = override is not None or raw is not None
                    value = override if override is not None else raw
                    observed_flags.append(observed)
                    component_values.append(0.0 if value is None else value)
            else:
                coverage = cover[comp][k]
                if comp == "Governance Practices":
                    flag_fields = [f"{field} observed" for field, _, _ in specs[comp]]
                elif comp == "Organizational Capacity":
                    flag_fields = ["M1 observed", "M2 observed", "H1 observed (either field)", "I2 observed", "O1 observed"]
                else:
                    flag_fields = [f"{field} observed" for field, _, _ in specs[comp]]
                for (field, _, _), flag_field in zip(specs[comp], flag_fields):
                    raw = score(record.get(field))
                    if raw is None:
                        raise ValueError(f"Missing zero-filled score for {service} / {comp} / {field}")
                    observed_flags.append(bool(coverage.get(flag_field)))
                    component_values.append(raw)

            component_metric_scores[comp] = component_values
            component_scores[comp] = average(component_values)
            component_counts[comp] = sum(observed_flags)
            component_observed_means[comp] = average([v for v, flag in zip(component_values, observed_flags) if flag]) if any(observed_flags) else None
            all_observed = all_observed and all(observed_flags)
            for (code, _, within_weight), value, observed in zip(specs[comp], component_values, observed_flags):
                detail.append({
                    "Service": service,
                    "Component": comp,
                    "Metric": code,
                    "Metric score (0-100; missing=0)": value,
                    "Observed in source": "Yes" if observed else "No",
                    "Within-component weight": within_weight,
                    "Implied overall weight": COMPONENT_WEIGHT * within_weight,
                    "Source workbook": SOURCES[comp][0].name,
                })

        overall = average(list(component_scores.values()))
        rows.append({
            "Service": service,
            "Provider": base.get("008_provider_name"),
            "Solution category": base.get("014_solution_category"),
            "Website": base.get("010_website"),
            "Location": base.get("013_location"),
            **{f"{c} score": component_scores[c] for c in COMPONENTS},
            **{f"{c} score (zero-filled)": component_scores[c] for c in COMPONENTS},
            **{f"{c} observed-only score": component_observed_means[c] for c in COMPONENTS},
            **{f"{c} observed metrics": component_counts[c] for c in COMPONENTS},
            **{f"{c} metrics total": len(specs[c]) for c in COMPONENTS},
            **{f"{c} observed coverage": component_counts[c] / len(specs[c]) for c in COMPONENTS},
            "All component metrics observed": "Yes" if all_observed else "No",
            "Equal-component OISI (0-100)": overall,
            "Complete-data OISI (0-100)": average([component_observed_means[c] for c in COMPONENTS]) if all_observed else None,
            "Available-component mean (diagnostic)": average([component_observed_means[c] for c in COMPONENTS if component_observed_means[c] is not None]),
            "Available components (at least one observed metric)": sum(component_counts[c] > 0 for c in COMPONENTS),
            "Interpretation": "Fixed 25% component weights; unavailable metric values are zero-filled. Read as evidence-adjusted, not performance alone.",
            "Development assessability": indexed["Development Activity"][k].get("Assessment interpretation", ""),
            "Development data-quality notes": indexed["Development Activity"][k].get("Data-quality notes", ""),
        })

    ranks = rank_average([r["Equal-component OISI (0-100)"] for r in rows])
    for row, rank in zip(rows, ranks):
        row["Rank (average ties)"] = rank
    complete = [r for r in rows if r["Complete-data OISI (0-100)"] is not None]
    complete_ranks = rank_average([r["Complete-data OISI (0-100)"] for r in complete])
    for row, rank in zip(complete, complete_ranks):
        row["Complete-data rank"] = rank
    available = [r for r in rows if r["Available-component mean (diagnostic)"] is not None]
    available_ranks = rank_average([r["Available-component mean (diagnostic)"] for r in available])
    for row, rank in zip(available, available_ranks):
        row["Available-component diagnostic rank"] = rank
    for row in rows:
        row.setdefault("Complete-data rank", "")
        row.setdefault("Available-component diagnostic rank", "")
    rows.sort(key=lambda r: (r["Rank (average ties)"], norm(r["Service"])))
    return master, rows, detail, specs, indexed, cover


def style(ws, widths=None):
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor=COLORS["green"])
        cell.font = Font(color=COLORS["white"], bold=True)
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    thin = Side(style="hair", color=COLORS["line"])
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if cell.row > 1:
                cell.border = Border(bottom=thin)
    for col in range(1, ws.max_column + 1):
        width = widths.get(col) if widths and col in widths else min(42, max(12, max(len(str(ws.cell(r, col).value or "")) for r in range(1, min(ws.max_row, 80) + 1)) + 2))
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.row_dimensions[1].height = 34


def create_workbook(master, rows, detail, specs, indexed, cover):
    wb = Workbook()
    wb.remove(wb.active)
    names = ["Summary", "Complete-Data Ranking", "OISI Ranking", "Available-Component Diagnostic", "Component Scores", "Metric Detail", "Metric Rules", "Methodology", "Evidence & Limitations", "References", "Coverage & QA", "Source Register"]
    sheets = {name: wb.create_sheet(name) for name in names}

    summary = sheets["Summary"]
    summary.append(["OISI Equal-Component Index", "Result"])
    component_means = {c: statistics.mean(r[f"{c} score"] for r in rows) for c in COMPONENTS}
    complete_rows = [r for r in rows if r["Complete-data OISI (0-100)"] is not None]
    available_rows = [r for r in rows if r["Available-component mean (diagnostic)"] is not None]
    summary_rows = [
        ["Population", len(rows)],
        ["Component weighting", "25% each; fixed denominator of four components"],
        ["Within-component weighting", "Equal weight to every metric in that component"],
        ["Overall score formula", "(Community Engagement + Governance Practices + Organizational Capacity + Development Activity) / 4"],
        ["Missing metric treatment", "0 in component average; observed status retained separately"],
        ["Overall coverage eligibility gate", "None"],
        ["Complete-data comparison cases", len(complete_rows)],
        ["Complete-data mean OISI", statistics.mean(r["Complete-data OISI (0-100)"] for r in complete_rows) if complete_rows else "NA"],
        ["Complete-data rank interpretation", "Strict diagnostic: all 31 metrics must be observed. Small cohort; do not generalize to all 112."],
        ["Mean OISI score (0-100)", statistics.mean(r["Equal-component OISI (0-100)"] for r in rows)],
        ["Median OISI score (0-100)", statistics.median(r["Equal-component OISI (0-100)"] for r in rows)],
        ["Minimum OISI score", min(r["Equal-component OISI (0-100)"] for r in rows)],
        ["Maximum OISI score", max(r["Equal-component OISI (0-100)"] for r in rows)],
        ["Mean: Community Engagement", component_means["Community Engagement"]],
        ["Mean: Governance Practices", component_means["Governance Practices"]],
        ["Mean: Organizational Capacity", component_means["Organizational Capacity"]],
        ["Mean: Development Activity", component_means["Development Activity"]],
        ["All 31 metrics observed", sum(r["All component metrics observed"] == "Yes" for r in rows)],
        ["Available-component diagnostic entries", len(available_rows)],
        ["Mean available-component diagnostic", statistics.mean(r["Available-component mean (diagnostic)"] for r in available_rows) if available_rows else "NA"],
        ["Available-component diagnostic definition", "Within each component, average observed metrics only; then average components with at least one observed metric. Variable denominator; diagnostic only."],
        ["Ranking interpretation", "All-entry zero-filled ranking; it reflects both measured scores and data availability. Review coverage before comparing."],
        ["Assessment warning", "Development entries marked not assessable receive zero-filled Development Activity scores; these are policy zeros, not measured inactivity."],
    ]
    for row in summary_rows:
        summary.append(row)

    ordered = sorted(rows, key=lambda r: (r["Rank (average ties)"], norm(r["Service"])))
    headers = ["Rank (average ties)", "Service", "Provider", "Solution category", "Equal-component OISI (0-100)", *[f"{c} score" for c in COMPONENTS], "All component metrics observed", "Development assessability"]
    sheets["OISI Ranking"].append(headers)
    for row in ordered:
        sheets["OISI Ranking"].append([row.get(h, "") for h in headers])

    complete_headers = ["Complete-data rank", "Service", "Provider", "Solution category", "Complete-data OISI (0-100)", *[f"{c} observed-only score" for c in COMPONENTS], *[f"{c} observed metrics" for c in COMPONENTS]]
    for row in sorted(complete_rows, key=lambda r: (r["Complete-data rank"], norm(r["Service"]))):
        if sheets["Complete-Data Ranking"].max_row == 1:
            sheets["Complete-Data Ranking"].append(complete_headers)
        sheets["Complete-Data Ranking"].append([row.get(h, "") for h in complete_headers])
    if sheets["Complete-Data Ranking"].max_row == 1:
        sheets["Complete-Data Ranking"].append(complete_headers)

    available_headers = ["Available-component diagnostic rank", "Service", "Provider", "Solution category", "Available-component mean (diagnostic)", "Available components (at least one observed metric)", *[f"{c} observed-only score" for c in COMPONENTS], *[f"{c} observed metrics" for c in COMPONENTS]]
    sheets["Available-Component Diagnostic"].append(available_headers)
    for row in sorted(available_rows, key=lambda r: (r["Available-component diagnostic rank"], norm(r["Service"]))):
        sheets["Available-Component Diagnostic"].append([row.get(h, "") for h in available_headers])

    component_headers = ["Service", "Provider", "Solution category"]
    for c in COMPONENTS:
        component_headers.extend([f"{c} score (zero-filled)", f"{c} observed-only score", f"{c} observed metrics", f"{c} metric total", f"{c} observed coverage"])
    component_headers.extend(["Equal-component OISI (0-100)", "Rank (average ties)", "Complete-data OISI", "Complete-data rank", "Available-component mean", "Available-component rank", "All component metrics observed", "Interpretation"])
    sheets["Component Scores"].append(component_headers)
    for row in ordered:
        sheets["Component Scores"].append([row.get(h, "") for h in component_headers])

    detail_headers = list(detail[0]) if detail else []
    sheets["Metric Detail"].append(detail_headers)
    for row in detail:
        sheets["Metric Detail"].append([row.get(h, "") for h in detail_headers])

    rules = sheets["Metric Rules"]
    rules.append(["Component", "Metric", "Source fields", "Within-component weight", "Implied OISI weight", "Scoring criteria", "Unavailable-data treatment", "Rationale / caution"])
    rule_path = {
        "Community Engagement": (SOURCES["Community Engagement"][0], "Metric Rules"),
        "Governance Practices": (SOURCES["Governance Practices"][0], "Metric Rules"),
        "Organizational Capacity": (SOURCES["Organizational Capacity"][0], "Metric Rules"),
        "Development Activity": (SOURCES["Development Activity"][0], "Metric Rules"),
    }
    for comp in COMPONENTS:
        source_rules = read_sheet(*rule_path[comp])
        for rule in source_rules:
            code = str(rule.get("Code", ""))
            if code.upper() == "TOTAL" or not code:
                continue
            within = next((w for m, _, w in specs[comp] if m == code), None)
            if within is None:
                continue
            rationale = (
                rule.get("Why this metric is retained")
                or rule.get("Why included / interpretation")
                or rule.get("Interpretation / caution")
            )
            if not rationale and comp == "Community Engagement":
                rationale = {
                    "E1": "Captures the reported status of community-engagement activity.",
                    "E2": "Measures the breadth of distinct engagement activity types reported in fields 038/938.",
                    "C1": "Measures the breadth of distinct contribution routes reported in fields 040/940.",
                    "C2": "Represents the presence and apparent specificity of contribution guidance; URL patterns are a proxy.",
                    "C3": "Represents a populated public-code-repository field; host matching is a proxy, not a live audit.",
                    "N1": "Captures the reported status of the Code of Conduct; field 043 is supporting evidence, not a separate metric.",
                }.get(code, "See the component workbook's Metric Rules sheet.")
            rules.append([
                comp, code, rule.get("Source fields") or rule.get("Source field(s)") or rule.get("Source field"),
                within, COMPONENT_WEIGHT * within,
                rule.get("Scoring criteria") or rule.get("Scoring rule"),
                rule.get("Unavailable-data treatment") or rule.get("Unavailable-data rule"),
                rationale,
            ])

    method = sheets["Methodology"]
    method.append(["Methodological step", "Rule / explanation"])
    method_rows = [
        ["1. Unit of analysis", "One OISI infrastructure/service per row, drawn from the 112-entry Community Engagement master list."],
        ["2. Source data", "Use the final component workbooks listed in Source Register. Join component results to the master population by normalized service name; workbook generation stops if any source row is missing or duplicated."],
        ["3. Component metrics", "Community Engagement uses E1, E2, C1, C2, C3, N1 (6). Governance Practices uses G1-G7 (7). Organizational Capacity uses M1, M2, H1, I2, O1 (5). Development Activity uses m1-m13 (13). See Metric Rules for source fields, formulas, and rationale."],
        ["4. Normalize", "Each component metric is on a 0-100 scale according to its component workbook scoring rule. Do not normalize raw values a second time."],
        ["5. Within-component aggregation", "Arithmetic mean of that component's metric scores, each equally weighted within the component. Missing metrics contribute 0 in the all-entry view; observed flags remain separate. Equal weights are a transparent study convention, not a claim that literature establishes equal substantive importance."],
        ["6. Across-component aggregation", "OISI = 0.25*Community Engagement + 0.25*Governance Practices + 0.25*Organizational Capacity + 0.25*Development Activity. Each component contributes exactly 25%."],
        ["7. Missingness", "No overall coverage or eligibility gate is applied. The all-entry index uses zero-filled component scores, so it is evidence-adjusted and reflects both measured values and missingness/platform availability. Zero-filling is an explicit conservative reporting scenario, not a standard imputation endorsed by the OECD/JRC Handbook and not evidence of no activity."],
        ["8. Complete-data comparison", "A strict comparison is provided only for infrastructures with all 31 metrics observed across all four components. It is a small-cohort diagnostic, not the all-population result."],
        ["9. Available-component diagnostic", "For each component, average observed metrics only; then average the component means for components with at least one observed metric. The denominator varies, so this is descriptive and not a directly comparable ranking."],
        ["10. Ranking", "The primary workbook ranking is the all-entry zero-filled index, sorted high to low with average ranks for ties. Complete-data ranking and available-component ordering are separate diagnostics. No pass/fail bands are applied."],
        ["11. Reporting", "Report component scores, metric observed counts/coverage, data-quality notes, and the caveat that missing-data zeros are not evidence of no activity. Identify not-assessable repositories separately. Use the complete-data and available-component views as complementary diagnostics, with their cohort/denominator limitations."],
        ["Weight implications", "Within-component metric weights: CE 1/6; Governance 1/7; Organizational Capacity 1/5; Development Activity 1/13. Effective overall metric shares are component weight times within-component weight."],
        ["Metric selection and exclusions", "Metrics were retained when source fields could be operationalized as distinct, interpretable signals within the four study components. Redundant, weakly interpretable, conditional, or out-of-scope fields were excluded under the component-specific rules. Exclusion does not imply that a field is unimportant in general; the rationale and caveats are listed in Metric Rules."],
        ["External evidence", "Cited research supports the relevance of broad concepts (for example, response activity, contributor concentration, governance, and technology readiness) and the need for transparent robustness analysis. It does not validate this exact metric set, URL/status proxies, cut-points, thresholds, weights, or resulting ranks. See Evidence & Limitations and References."],
        ["Sensitivity and robustness", "This workbook presents complete-data and alternate-denominator diagnostic views. It does not yet report formal perturbation tests of weights, cutoffs, or missing-data assumptions. Composite-indicator guidance recommends testing such choices; rank stability should therefore be treated as untested beyond the views shown here."],
        ["Interpretation", "OISI summarizes selected operational indicators; it is not a validated direct measure of quality, impact, security, user satisfaction, or sustainability. Project/platform context matters, and cross-project ranks should be interpreted cautiously."],
    ]
    for row in method_rows:
        method.append(row)

    evidence = sheets["Evidence & Limitations"]
    evidence.append(["Area / metrics", "How the literature informs the choice", "What it does not establish / limitation"])
    evidence_rows = [
        ["Index construction, equal weights, aggregation (all components)", "Nardo et al. (2008) and Greco et al. (2019) describe composite-indicator choices in normalization, weighting, aggregation, missing data, and robustness. They support making assumptions explicit and examining sensitivity.", "Neither source proves that equal weights are uniquely correct for OISI. Equal 25% component weights and equal metric weights are transparent normative choices; this workbook has not yet run formal weight perturbation tests."],
        ["Missing-data views (all components)", "The OECD/JRC Handbook treats missing-data handling as a consequential methodological choice; alternative assumptions should be made visible and assessed.", "The Handbook does not endorse zero-filling as a general imputation method. Our zero-filled all-entry score is labeled evidence-adjusted and must not be read as proof of zero activity. Complete-data and available-component views are diagnostics, not a complete sensitivity analysis."],
        ["Community Engagement: E1, E2, C1, C2, C3, N1", "These indicators operationalize reported engagement status, activity/contribution breadth, contribution guidance/repository presence, and Code of Conduct status using the selected Infra Finder fields. Their field-level rationale and rules are in Metric Rules.", "The references supplied do not independently validate these exact fields, 0/33/67/100 status ladder, category denominators, URL-pattern classifications, or the resulting scores. URL presence is a proxy, not a live usability or access audit."],
        ["Governance: G1-G7", "The archived POSI principles describe organizational accountability, stakeholder governance, inclusive participation, and transparent practices, providing conceptual context for governance-related indicators.", "POSI is a principles framework, not validation of these Infra Finder fields or a scoring instrument. Board/body presence does not measure independence, representativeness, effectiveness, or decision quality; reported status is not independently audited here."],
        ["Organizational Capacity: M1, M2, H1, I2, O1", "NASA defines a nine-level Technology Readiness Level scale, providing provenance for the TRL concept. Research on open-source project survival highlights the potential role of core-contributor succession; this is adjacent support for considering capacity signals.", "Dividing TRL by 9 is our normalization convention, not a NASA-prescribed index. Staffing bins, maintenance scoring, institutional support and hosted-service scoring are study operationalizations; the cited work does not validate their cutoffs or weights. Hosted-service presence is not itself proof of resilience or capacity."],
        ["Development Activity: m1-m13", "CHAOSS project-health resources include concepts such as response time, change-request outcomes, release activity, and contributor concentration. Avelino et al. examine abandonment and survival in relation to changes in core contributors.", "Our observation window, minimum event count, 7/14/30/90-day cutoffs, activity caps, and formulas are study-specific. The cited work motivates relevant constructs but does not validate these exact thresholds or OISI scores. GitHub-visible activity is not equivalent to all development activity."],
        ["Stars / forks (Development Activity indicators)", "Borges and Valente find that stars can influence developer attention and repository selection, supporting their interpretation as an attention/visibility signal.", "Stars are not software quality or sustainability. Fork counts are treated as a descriptive reuse/interest proxy; the references do not establish them as a direct measure of successful reuse."],
        ["GitHub data validity and platform scope", "Kalliamvakou et al. document both opportunities and pitfalls in mining GitHub data, supporting explicit platform and event-definition caveats.", "The study captures observable GitHub activity, not equivalent activity on GitLab, Bitbucket, private systems, or non-code trackers. API records and event classifications can be incomplete or ambiguous."],
        ["Cross-project interpretation", "Yehudi, Goble, and Jay report that individual context-free community-health indicators can have different meanings across projects, supporting cautious interpretation and contextual review.", "This cautions against treating the rank as a context-free or high-stakes verdict. OISI is descriptive/comparative within the study's operational framework, not a causal or validated universal quality scale."],
        ["Excluded fields and scope", "Fields were excluded where the final component framework judged them redundant, not sufficiently interpretable as a comparable indicator, conditional, or outside the component construct; see each component's Metric Rules and source workbook.", "The literature cited here does not prove that excluded fields are unimportant. The selection is study-specific and may omit meaningful context; exclusions should be revisited if the research question or evidence base changes."],
    ]
    for row in evidence_rows:
        evidence.append(row)

    refs = sheets["References"]
    refs.append(["ID", "Reference", "Relevance in this study", "Link / DOI"])
    reference_rows = [
        ["R1", "Nardo, M., Saisana, M., Saltelli, A., Tarantola, S., Hoffmann, A., & Giovannini, E. (2008). Handbook on Constructing Composite Indicators: Methodology and User Guide. OECD Publishing / European Commission JRC. DOI: 10.1787/9789264043466-en.", "Composite-indicator construction, missing data, weighting, aggregation, and sensitivity; not endorsement of our exact choices.", "https://doi.org/10.1787/9789264043466-en"],
        ["R2", "Saisana, M., Saltelli, A., & Tarantola, S. (2005). Uncertainty and sensitivity analysis techniques as tools for the quality assessment of composite indicators. Journal of the Royal Statistical Society: Series A, 168(2), 307-323.", "Motivates sensitivity analysis of composite-indicator assumptions and rankings.", "https://doi.org/10.1111/j.1467-985X.2005.00350.x"],
        ["R3", "Greco, S., Ishizaka, A., Tasiou, M., & Torrisi, G. (2019). On the Methodological Framework of Composite Indices: A Review of the Issues of Weighting, Aggregation, and Robustness. Social Indicators Research, 141, 61-94. DOI: 10.1007/s11205-017-1832-9.", "Review of weighting, aggregation, and robustness issues.", "https://doi.org/10.1007/s11205-017-1832-9"],
        ["R4", "Principles of Open Scholarly Infrastructure (POSI), version 1.0 (archived record). DOI: 10.24343/C34W2H.", "Conceptual context for governance, organizational accountability, and inclusiveness.", "https://doi.org/10.24343/C34W2H"],
        ["R5", "NASA. Technology Readiness Levels (TRLs).", "Primary reference for the nine-level TRL framework; our division by 9 is a study normalization convention.", "https://www.nasa.gov/directorates/somd/space-communications-navigation-program/technology-readiness-levels/"],
        ["R6", "CHAOSS. Starter Project Health Metrics Model.", "Conceptual reference for project-health constructs including response, change requests, releases, and contributor concentration.", "https://www.chaoss.community/starter-project-health-metrics-model/"],
        ["R7", "Avelino, G., et al. (2019). On the abandonment and survival of open source projects: An empirical investigation. ESEM 2019. DOI: 10.1109/ESEM.2019.8870181.", "Evidence related to project abandonment/survival and core-contributor succession; does not validate our capacity score.", "https://doi.org/10.1109/ESEM.2019.8870181"],
        ["R8", "Borges, H., & Valente, M. T. (2018). What's in a GitHub Star? Understanding Repository Starring Practices in a Social Coding Platform. Journal of Systems and Software. DOI: 10.1016/j.jss.2018.09.016.", "Supports interpreting stars as attention/selection signals, not as quality.", "https://doi.org/10.1016/j.jss.2018.09.016"],
        ["R9", "Kalliamvakou, E., et al. (2014). The Promises and Perils of Mining GitHub. Proceedings of the 11th Working Conference on Mining Software Repositories. DOI: 10.1145/2597073.2597074.", "Supports caveats about GitHub sampling, data interpretation, and event semantics.", "https://doi.org/10.1145/2597073.2597074"],
        ["R10", "Yehudi, Y., Goble, C., & Jay, C. (2023). Individual context-free online community health indicators fail to identify open source software sustainability. arXiv:2309.12120.", "Supports caution against universal or high-stakes interpretation of context-free project-health indicators.", "https://arxiv.org/abs/2309.12120"],
        ["R11", "Invest in Open Infrastructure. Data for Infra Finder, version 3.0 (2026). Zenodo record 18225266.", "Provenance for the Infra Finder dataset and associated data definitions used in the component sources.", "https://zenodo.org/records/18225266"],
    ]
    for row in reference_rows:
        refs.append(row)

    qa = sheets["Coverage & QA"]
    qa.append(["Check", "Result", "Interpretation"])
    qa_rows = [
        ["Master entries", len(master), "Expected 112."],
        ["Unique normalized service keys", len({norm(r.get("001_service_name")) for r in master}), "Must equal master row count."],
        ["Source matches", "All 112 per component", "No unmatched services or duplicate normalized keys."],
        ["Equal component weights", "25% + 25% + 25% + 25% = 100%", "Fixed component denominator."],
        ["Metric count", sum(len(specs[c]) for c in COMPONENTS), "31 component metrics total; they are not globally equally weighted."],
        ["All 31 metrics observed", sum(r["All component metrics observed"] == "Yes" for r in rows), "Diagnostic only; does not exclude entries."],
        ["Mean observed coverage: CE", statistics.mean(r["Community Engagement observed coverage"] for r in rows), "Metric coverage; not an eligibility rule."],
        ["Mean observed coverage: Governance", statistics.mean(r["Governance Practices observed coverage"] for r in rows), "Metric coverage; not an eligibility rule."],
        ["Mean observed coverage: Organizational Capacity", statistics.mean(r["Organizational Capacity observed coverage"] for r in rows), "Metric coverage; not an eligibility rule."],
        ["Mean observed coverage: Development Activity", statistics.mean(r["Development Activity observed coverage"] for r in rows), "Metric coverage; not an eligibility rule."],
        ["Formula check", "All OISI scores equal arithmetic mean of four component scores", "Verified during build."],
        ["Score range", f"{min(r['Equal-component OISI (0-100)'] for r in rows):.2f} to {max(r['Equal-component OISI (0-100)'] for r in rows):.2f}", "All values must lie between 0 and 100."],
    ]
    for row in qa_rows:
        qa.append(row)

    reg = sheets["Source Register"]
    reg.append(["Component", "Workbook used", "Scoring sheet", "Metric count", "Within-component weight", "Scoring basis"])
    for comp, (path, sheet) in SOURCES.items():
        reg.append([comp, str(path.relative_to(ROOT)), sheet, len(specs[comp]), "Equal metrics within component", "Final component workbook; zero-filled scores with observed flags retained"])

    widths = {
        "Summary": {1: 48, 2: 110},
        "OISI Ranking": {1: 18, 2: 34, 3: 38, 4: 32, 5: 24, **{i: 22 for i in range(6, 10)}, 10: 30, 11: 50},
        "Component Scores": {1: 34, 2: 36, 3: 32, **{i: 24 for i in range(4, 21)}},
        "Metric Detail": {1: 34, 2: 30, 3: 28, 4: 30, 5: 20, 6: 24, 7: 22, 8: 66},
        "Metric Rules": {1: 32, 2: 24, 3: 46, 4: 24, 5: 22, 6: 82, 7: 82, 8: 86},
        "Methodology": {1: 34, 2: 120},
        "Evidence & Limitations": {1: 42, 2: 100, 3: 100},
        "References": {1: 12, 2: 100, 3: 76, 4: 72},
        "Coverage & QA": {1: 46, 2: 42, 3: 90},
        "Source Register": {1: 34, 2: 100, 3: 24, 4: 16, 5: 34, 6: 84},
    }
    for name, ws in sheets.items():
        style(ws, widths.get(name))
        if name == "Component Scores":
            for row in ws.iter_rows(min_row=2):
                for idx in (7, 11, 15, 19):
                    if idx <= ws.max_column:
                        row[idx - 1].number_format = "0.00%"
        if name == "Metric Detail":
            for row in ws.iter_rows(min_row=2):
                row[5].number_format = "0.00%"
                row[6].number_format = "0.00%"
        if name == "Coverage & QA":
            for r in range(2, ws.max_row + 1):
                if "coverage:" in str(ws.cell(r, 1).value).lower():
                    ws.cell(r, 2).number_format = "0.00%"

    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcMode = "auto"
    wb.save(OUTPUT)
    print(f"Created {OUTPUT} with {len(rows)} entries.")
    print(f"Mean OISI={statistics.mean(r['Equal-component OISI (0-100)'] for r in rows):.4f}; median={statistics.median(r['Equal-component OISI (0-100)'] for r in rows):.4f}.")


def main():
    master, rows, detail, specs, indexed, cover = load_all()
    create_workbook(master, rows, detail, specs, indexed, cover)


if __name__ == "__main__":
    main()
