from __future__ import annotations

import math
import re
from collections import Counter
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


SOURCE = Path("OISI_112_Organizational_Capacity_Verification.xlsx")
OUTPUT = Path("OISI_Organizational_Capacity_Scoring_v3.xlsx")
AS_OF_YEAR = 2026
AGE_CAP = 15


METRICS = [
    ("M1", "Technology readiness", "079_technology_readiness", "Maturity & Maintenance", 0.10),
    ("M2", "Maintenance status", "080_maintenance_status", "Maturity & Maintenance", 0.20),
    ("M3", "Infrastructure age", "056_launch_date", "Maturity & Maintenance", 0.05),
    ("I1", "Formal organizational host", "051_business_form", "Institutional Support", 0.15),
    ("I2", "Affiliations / institutional support", "058_affiliations", "Institutional Support", 0.15),
    ("I3", "SCOSS participation", "099_scoss", "Institutional Support", 0.05),
    ("O1", "Hosted/SaaS service availability", "076_saas_options", "Operational Capacity", 0.20),
    ("O2", "Registered service-provider ecosystem", "077_registered_service_providers", "Operational Capacity", 0.10),
]


DATA_COLUMNS = [
    "001_service_name",
    "008_provider_name",
    "010_website",
    "013_location",
    "014_solution_category",
    "019_open_code_repository_url",
    "017_funding_needs",
    "051_business_form",
    "951_business_form_other",
    "052_nonprofit_status",
    "053_shareholders",
    "054_staffing_full_time",
    "055_staffing_volunteer",
    "056_launch_date",
    "057_organizational_history",
    "058_affiliations",
    "059_founders",
    "061_primary_funding_source",
    "961_primary_funding_source_other",
    "076_saas_options",
    "077_registered_service_providers",
    "079_technology_readiness",
    "080_maintenance_status",
    "099_scoss",
]


def clean(value) -> str:
    if value is None:
        return ""
    text = re.sub(r"<[^>]+>", " ", str(value))
    return re.sub(r"\s+", " ", text).strip()


def lower(value) -> str:
    return clean(value).lower()


def technology_score(value):
    text = lower(value)
    if text == "actual system proven in operational environment":
        return 1.0
    if text == "system complete and qualified":
        return 8 / 9
    if text == "technology validated in relevant environment":
        return 5 / 9
    return None


def maintenance_score(value):
    text = lower(value)
    if text == "actively maintained":
        return 1.0
    if text == "minimally maintained":
        return 0.5
    if text in {"inactive", "not maintained"}:
        return 0.0
    return None


def age_score(value):
    text = clean(value)
    if not text:
        return None
    match = re.match(r"(\d{4})", text)
    if not match:
        return None
    age = max(0, AS_OF_YEAR - int(match.group(1)))
    return min(age, AGE_CAP) / AGE_CAP


def host_score(value):
    text = lower(value)
    if not text or text == "other":
        return None
    if "volunteer community" in text:
        return 0.5
    if any(token in text for token in ["non-profit", "academic institution", "commercial", "fiscal sponsorship"]):
        return 1.0
    return None


def affiliation_score(value):
    text = lower(value)
    if not text:
        return None
    if "no affiliation" in text:
        return 0.0
    return 1.0


def scoss_score(value):
    text = lower(value)
    if text == "true":
        return 1.0
    if text == "false":
        return 0.0
    return None


def saas_score(value):
    text = lower(value)
    if not text or text in {"unknown", "not applicable"}:
        return None
    if text == "no services available":
        return 0.0
    if text.startswith("through "):
        return 1.0
    return None


def provider_applicable(saas_value):
    text = lower(saas_value)
    return text in {"through solution provider or third party vendor", "through third party vendor only"}


def provider_score(provider_value, saas_value):
    if provider_applicable(saas_value):
        return 1.0 if clean(provider_value) else 0.0
    return None


def score_row(row):
    scores = {
        "M1": technology_score(row["079_technology_readiness"]),
        "M2": maintenance_score(row["080_maintenance_status"]),
        "M3": age_score(row["056_launch_date"]),
        "I1": host_score(row["051_business_form"]),
        "I2": affiliation_score(row["058_affiliations"]),
        "I3": scoss_score(row["099_scoss"]),
        "O1": saas_score(row["076_saas_options"]),
        "O2": provider_score(row["077_registered_service_providers"], row["076_saas_options"]),
    }

    weight_by_code = {code: weight for code, _, _, _, weight in METRICS}
    group_by_code = {code: group for code, _, _, group, _ in METRICS}
    applicable = {code: True for code, *_ in METRICS}
    if lower(row["076_saas_options"]) == "not applicable":
        applicable["O1"] = False
    if not provider_applicable(row["076_saas_options"]):
        applicable["O2"] = False

    weight_scored = sum(weight_by_code[code] for code, value in scores.items() if value is not None)
    weight_applicable = sum(weight_by_code[code] for code, is_applicable in applicable.items() if is_applicable)
    weighted_sum = sum(value * weight_by_code[code] for code, value in scores.items() if value is not None)
    groups = {group_by_code[code] for code, value in scores.items() if value is not None}
    completeness = weight_scored / weight_applicable if weight_applicable else 0
    eligible = completeness >= 0.70 and len(groups) >= 2
    final_score = weighted_sum / weight_scored * 100 if eligible else None

    flags = []
    if scores["M1"] is not None and scores["M1"] < 1:
        flags.append("Technology readiness below operational-proven")
    if scores["M2"] is not None and scores["M2"] < 1:
        flags.append("Maintenance below active")
    if scores["I2"] is None and clean(row["058_affiliations"]):
        flags.append("Affiliation text not coded")
    if provider_applicable(row["076_saas_options"]) and not clean(row["077_registered_service_providers"]):
        flags.append("Third-party service channel but no provider list")
    if scores["O1"] is None and lower(row["076_saas_options"]) == "unknown":
        flags.append("SaaS options unknown")

    return scores, {
        "weight_scored": weight_scored,
        "weight_applicable": weight_applicable,
        "weighted_sum": weighted_sum,
        "completeness": completeness,
        "groups": len(groups),
        "subcomponents": "".join(
            label
            for label, group in [
                ("M", "Maturity & Maintenance"),
                ("I", "Institutional Support"),
                ("O", "Operational Capacity"),
            ]
            if group in groups
        ),
        "eligible": eligible,
        "score": final_score,
        "status": "Reportable" if eligible and completeness >= 0.85 else ("Provisional" if eligible else "Not scored"),
        "lower": weighted_sum / weight_applicable * 100 if weight_applicable else None,
        "upper": (weighted_sum + (weight_applicable - weight_scored)) / weight_applicable * 100 if weight_applicable else None,
        "flags": "; ".join(flags),
    }


def subscore(scores, codes):
    weight_by_code = {code: weight for code, _, _, _, weight in METRICS}
    scored = [(scores[code], weight_by_code[code]) for code in codes if scores[code] is not None]
    if not scored:
        return "NA"
    return sum(value * weight for value, weight in scored) / sum(weight for _, weight in scored)


def recompute(scores, threshold=0.70, group_min=2, overrides=None):
    values = dict(scores)
    if overrides:
        values.update(overrides)
    weight_by_code = {code: weight for code, _, _, _, weight in METRICS}
    group_by_code = {code: group for code, _, _, group, _ in METRICS}
    weight_scored = sum(weight_by_code[code] for code, value in values.items() if value is not None)
    weighted_sum = sum(value * weight_by_code[code] for code, value in values.items() if value is not None)
    groups = {group_by_code[code] for code, value in values.items() if value is not None}
    if weight_scored >= threshold and len(groups) >= group_min:
        return weighted_sum / weight_scored * 100
    return "NA"


def median(values):
    if not values:
        return None
    values = sorted(values)
    mid = len(values) // 2
    if len(values) % 2:
        return values[mid]
    return (values[mid - 1] + values[mid]) / 2


def pct(values, predicate):
    return sum(1 for value in values if predicate(value)) / len(values) if values else None


def autosize(ws):
    for col in range(1, ws.max_column + 1):
        width = 10
        for row in range(1, ws.max_row + 1):
            value = ws.cell(row, col).value
            width = max(width, len(str(value)) + 2 if value is not None else 0)
        ws.column_dimensions[get_column_letter(col)].width = min(width, 55)


def style_sheet(ws):
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A2"
    autosize(ws)


def main():
    wb_source = openpyxl.load_workbook(SOURCE, data_only=True)
    source = wb_source["Organizational Capacity"]
    headers = [source.cell(2, c).value for c in range(1, source.max_column + 1)]

    rows = []
    for r in range(4, source.max_row + 1):
        source_row = {header: source.cell(r, i + 1).value for i, header in enumerate(headers)}
        row = {column: source_row.get(column) for column in DATA_COLUMNS}
        scores, meta = score_row(row)
        rows.append((row, scores, meta))

    ranked = sorted([item for item in rows if item[2]["score"] is not None], key=lambda item: item[2]["score"], reverse=True)
    ranks = {item[0]["001_service_name"]: idx for idx, item in enumerate(ranked, 1)}
    score_values = [meta["score"] for _, _, meta in rows if meta["score"] is not None]
    statuses = Counter(meta["status"] for _, _, meta in rows)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    method = wb.create_sheet("Method")
    parameters = wb.create_sheet("Parameters")
    scoring_rules = wb.create_sheet("Scoring Rules")
    scores_ws = wb.create_sheet("Scores")
    diagnostics = wb.create_sheet("Diagnostics")
    sensitivity = wb.create_sheet("Sensitivity")

    method_rows = [
        ["OISI Component 3 - Organizational Capacity v3: scoring workbook"],
        ["Source", "OISI_112_Organizational_Capacity_Verification.xlsx"],
        ["Purpose", "Measures whether an infrastructure has maintained technology, institutional support, and operational capacity for continuity."],
        ["Design", "Weights adjusted to 35% Maturity & Maintenance, 35% Institutional Support, 30% Operational Capacity."],
        ["Reason for adjustment", "Technology readiness and maintenance are retained, but reduced because the data are near-constant and otherwise inflate scores."],
        ["Formula", "Score = 100 x sum(metric score x metric weight) / sum(scored metric weights). Missing values are excluded."],
        ["Completeness", "A score is produced when at least 70% of weight is scored and at least two sub-components are present."],
    ]
    for row in method_rows:
        method.append(row)

    parameters.append(["Code", "Metric", "Variable", "Sub-component", "Weight", "Scoring rule"])
    rules = {
        "M1": "Actual operational system=1; complete/qualified=8/9; validated relevant environment=5/9; otherwise NA.",
        "M2": "Actively maintained=1; minimally maintained=0.5; inactive/not maintained=0; unknown/blank=NA.",
        "M3": f"Age = {AS_OF_YEAR} minus launch year; capped at {AGE_CAP} years; score=min(age,{AGE_CAP})/{AGE_CAP}.",
        "I1": "Formal organizational host=1; volunteer community=0.5; Other/blank=NA.",
        "I2": "Named affiliation/support=1; explicit no affiliations=0; blank=NA.",
        "I3": "SCOSS TRUE=1; FALSE=0; unknown/blank=NA.",
        "O1": "Hosted/SaaS channel available=1; no services available=0; not applicable/unknown=NA.",
        "O2": "If third-party service channel applies: provider list=1, no provider list=0; otherwise NA.",
    }
    for code, metric, variable, group, weight in METRICS:
        parameters.append([code, metric, variable, group, weight, rules[code]])
    parameters.append(["", "", "", "Total", sum(weight for *_, weight in METRICS), ""])
    parameters.append([])
    parameters.append(["Setting", "Value", "Note"])
    parameters.append(["Minimum completeness", 0.70, "Minimum scored weight needed for a result"])
    parameters.append(["Reportable threshold", 0.85, "Below this threshold, eligible scores are provisional"])
    parameters.append(["Minimum sub-components", 2, "At least two of Maturity, Institutional, Operational"])
    parameters.append(["As-of year", AS_OF_YEAR, "Used for launch-date age score"])
    parameters.append(["Age cap", AGE_CAP, "Years required for full age score"])

    scoring_rules.append(["Section", "Rule", "Details / verification guidance"])
    scoring_rule_rows = [
        [
            "Overall formula",
            "Weighted average over scored metrics",
            "Organizational Capacity Score = 100 x sum(metric score x metric weight) / sum(scored metric weights).",
        ],
        [
            "Missing-data rule",
            "Unknown, blank, and structurally not applicable values are not automatically scored as zero",
            "They are excluded from the scored-weight denominator unless the field is applicable and absence itself is meaningful.",
        ],
        [
            "Applicable denominator",
            "Completeness is measured against applicable metric weight",
            "Example: if 076_saas_options is Not applicable, O1 is removed from the applicable denominator. If third-party service delivery is not applicable, O2 is also removed.",
        ],
        [
            "Eligibility rule",
            "A score is produced only when applicable completeness is at least 70% and at least two sub-components are represented",
            "Sub-components are Maturity & Maintenance, Institutional Support, and Operational Capacity.",
        ],
        [
            "Result status",
            "Reportable or Provisional",
            "Reportable = eligible and applicable completeness >= 85%. Provisional = eligible and applicable completeness >= 70% but < 85%. Not scored = below eligibility threshold.",
        ],
        [
            "M1 Technology readiness",
            "079_technology_readiness, weight 10%",
            "Actual system proven in operational environment = 1.00; System complete and qualified = 8/9; Technology validated in relevant environment = 5/9; otherwise NA.",
        ],
        [
            "M2 Maintenance status",
            "080_maintenance_status, weight 20%",
            "Actively Maintained = 1.00; Minimally Maintained = 0.50; Inactive/not maintained = 0.00; Unknown/blank = NA.",
        ],
        [
            "M3 Infrastructure age",
            "056_launch_date, weight 5%",
            f"Age = {AS_OF_YEAR} minus launch year. Score = min(age, {AGE_CAP}) / {AGE_CAP}. Full credit at {AGE_CAP}+ years.",
        ],
        [
            "I1 Formal organizational host",
            "051_business_form, weight 15%",
            "Formal organizational backing, including non-profit, academic fiscal sponsorship, non-profit fiscal sponsorship, or commercial organization = 1.00; Volunteer community = 0.50; Other/blank = NA.",
        ],
        [
            "I2 Affiliations / institutional support",
            "058_affiliations, weight 15%",
            "Named affiliation or institutional support = 1.00; explicit No affiliations = 0.00; blank = NA.",
        ],
        [
            "I3 SCOSS participation",
            "099_scoss, weight 5%",
            "TRUE = 1.00; FALSE = 0.00; unknown/blank = NA. This is a small support signal, not a core capacity requirement.",
        ],
        [
            "O1 Hosted/SaaS service availability",
            "076_saas_options, weight 20%",
            "Through solution provider only, through third-party vendor only, or through solution provider/third-party vendor = 1.00; No services available = 0.00; Not applicable/Unknown/blank = NA.",
        ],
        [
            "O2 Registered service-provider ecosystem",
            "077_registered_service_providers, weight 10%",
            "Applicable only when 076 indicates third-party service delivery. Provider list/evidence present = 1.00; third-party delivery but no provider list = 0.00; otherwise NA.",
        ],
        [
            "Lower bound",
            "Missing applicable metrics treated as zero",
            "Lower bound = weighted sum / applicable weight x 100.",
        ],
        [
            "Upper bound",
            "Missing applicable metrics treated as one",
            "Upper bound = (weighted sum + applicable missing weight) / applicable weight x 100.",
        ],
        [
            "Descriptive fields",
            "017, 052, 053, 054, 055, 057, 059, 061 are not in the core score",
            "These fields can support review, narrative interpretation, subgrouping, or future sensitivity analysis, but they do not directly increase or decrease the v3 score.",
        ],
    ]
    for rule_row in scoring_rule_rows:
        scoring_rules.append(rule_row)

    scores_headers = [
        "Service",
        "Provider",
        "Solution category",
        "079 Technology readiness",
        "080 Maintenance status",
        "056 Launch date",
        "051 Business form",
        "058 Affiliations",
        "099 SCOSS",
        "076 SaaS options",
        "077 Registered service providers",
        "061 Primary funding (descriptive)",
        "054 Staffing FT (descriptive)",
        "055 Staffing volunteer (descriptive)",
        *[f"{code} {metric}" for code, metric, *_ in METRICS],
        "Maturity & Maintenance (0-1)",
        "Institutional Support (0-1)",
        "Operational Capacity (0-1)",
        "Weight scored",
        "Weight applicable",
        "Weighted sum",
        "Completeness",
        "Sub-components used",
        "Eligible?",
        "ORG CAPACITY SCORE (0-100)",
        "Lower bound (missing=0)",
        "Upper bound (missing=1)",
        "Result status",
        "Rank",
        "Data-quality flags",
    ]
    scores_ws.append(scores_headers)
    for row, row_scores, meta in rows:
        scores_ws.append(
            [
                row["001_service_name"],
                row["008_provider_name"],
                row["014_solution_category"],
                row["079_technology_readiness"],
                row["080_maintenance_status"],
                row["056_launch_date"],
                row["051_business_form"],
                row["058_affiliations"],
                row["099_scoss"],
                row["076_saas_options"],
                row["077_registered_service_providers"],
                row["061_primary_funding_source"],
                row["054_staffing_full_time"],
                row["055_staffing_volunteer"],
                *[row_scores[code] for code, *_ in METRICS],
                subscore(row_scores, ["M1", "M2", "M3"]),
                subscore(row_scores, ["I1", "I2", "I3"]),
                subscore(row_scores, ["O1", "O2"]),
                meta["weight_scored"],
                meta["weight_applicable"],
                meta["weighted_sum"],
                meta["completeness"],
                meta["subcomponents"],
                "Yes" if meta["eligible"] else "No",
                meta["score"],
                meta["lower"],
                meta["upper"],
                meta["status"],
                ranks.get(row["001_service_name"], "NA"),
                meta["flags"],
            ]
        )

    diagnostics.append(["Metric", "Weight", "Scored values (n)", "Coverage", "Mean", "Std dev", "% at 1.0", "% at 0", "Reading"])
    metric_columns = {code: [] for code, *_ in METRICS}
    for _, row_scores, _ in rows:
        for code in metric_columns:
            if isinstance(row_scores[code], (int, float)):
                metric_columns[code].append(row_scores[code])
    for code, metric, _, _, weight in METRICS:
        values = metric_columns[code]
        mean = sum(values) / len(values) if values else None
        sd = math.sqrt(sum((value - mean) ** 2 for value in values) / len(values)) if values else None
        if code in {"M1", "M2"}:
            reading = "Near-constant; kept for construct validity but weight reduced."
        elif code == "I3":
            reading = "Rare positive signal; behaves like a small institutional-support bonus."
        elif code == "O2":
            reading = "Sparse but discriminating where third-party provider ecosystem is applicable."
        else:
            reading = "Useful capacity signal."
        diagnostics.append(
            [
                f"{code} {metric}",
                weight,
                len(values),
                len(values) / len(rows),
                mean,
                sd,
                pct(values, lambda value: abs(value - 1) < 1e-9),
                pct(values, lambda value: abs(value) < 1e-9),
                reading,
            ]
        )
    diagnostics.append([])
    diagnostics.append(["Result summary", "Value"])
    diagnostics.append(["Total services", len(rows)])
    diagnostics.append(["Services scored", len(score_values)])
    diagnostics.append(["Reportable", statuses["Reportable"]])
    diagnostics.append(["Provisional", statuses["Provisional"]])
    diagnostics.append(["Not scored", statuses["Not scored"]])
    diagnostics.append(["Median score", median(score_values)])
    diagnostics.append(["Mean score", sum(score_values) / len(score_values) if score_values else None])

    sensitivity.append(["Variant", "Services scored", "Mean score", "Median score", "Interpretation"])
    variants = [
        ("Baseline v3", lambda scores: recompute(scores), "Reference"),
        ("Lower completeness threshold 60%", lambda scores: recompute(scores, threshold=0.60), "Shows effect of allowing weaker evidence base"),
        ("Higher completeness threshold 85%", lambda scores: recompute(scores, threshold=0.85), "Shows high-completeness-only results"),
        ("Maturity near-constant fields dropped", lambda scores: recompute({**scores, "M1": None, "M2": None}, threshold=0.50), "Tests dependence on 079/080"),
        ("Provider ecosystem dropped", lambda scores: recompute({**scores, "O2": None}), "Tests sensitivity to sparse 077"),
    ]
    for name, fn, note in variants:
        vals = [value for _, row_scores, _ in rows if isinstance((value := fn(row_scores)), (int, float))]
        sensitivity.append([name, len(vals), sum(vals) / len(vals) if vals else None, median(vals), note])

    for ws in [method, parameters, scoring_rules, scores_ws, diagnostics, sensitivity]:
        style_sheet(ws)
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, float):
                    cell.number_format = "0.00"

    wb.save(OUTPUT)


if __name__ == "__main__":
    main()
