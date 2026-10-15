from __future__ import annotations

import math
import os
import re
import statistics
import tempfile
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter


HERE = Path(__file__).resolve().parent
GOV_SOURCE = HERE / "OISI_112_Governance_Practices_Verification.xlsm"
CE_SOURCE = HERE.parent / "Community-Engagement" / "OISI_112_Community_Engagement_Verification (1).xlsx"
OUTPUT = HERE / "OISI_Governance_Practices_Final_Defensible_Audit_Corrected.xlsx"
GOV_SHEET = "Governance Practices"
CE_SHEET = "Community Engagement"
DATA_START = 4

GOV_FIELDS = [
    "001_service_name", "008_provider_name", "010_website", "013_location",
    "014_solution_category", "019_open_code_repository_url", "024_governance_structure_status",
    "025_governance_structure_url", "026_governance_records_status", "027_governance_summary",
    "028_governance_records_url", "050_board_structure", "950_board_structure_other",
    "100_board_level", "900_board_level_other", "108_bylaws_status",
    "029_web_accessibility_status", "030_web_accessibility_url", "031_web_accessibility_applicability",
    "033_transparent_pricing_status", "034_transparent_pricing_url", "035_commitment_to_equity_status",
    "036_commitment_to_equity_url", "044_commitment_to_privacy_status", "045_commitment_to_privacy_url",
    "046_values_frameworks", "verification_status", "verification_notes",
]
STATUS_FIELDS = {
    "G1": "024_governance_structure_status",
    "G2": "108_bylaws_status",
    "G4": "029_web_accessibility_status",
    "G5": "035_commitment_to_equity_status",
    "G6": "044_commitment_to_privacy_status",
}
METRICS = [
    ("G1", "Governance structure status", "024_governance_structure_status"),
    ("G2", "Bylaws or equivalent formal rules", "108_bylaws_status"),
    ("G3", "Board or steering-body presence", "050_board_structure + 950_board_structure_other"),
    ("G4", "Web accessibility status", "029_web_accessibility_status"),
    ("G5", "Equity commitment status", "035_commitment_to_equity_status"),
    ("G6", "Privacy commitment status", "044_commitment_to_privacy_status"),
    ("G7", "Community governance formality", "048_community_governance_status"),
]
STATUS_SCORE = {"implemented/available": 100, "in progress": 67, "considering": 33, "not planning": 0}
GOV_COLORS = {"teal": "245B55", "white": "FFFFFF", "gold": "FCECC9", "light": "E8F1EF"}


def clean(value) -> str:
    if value is None:
        return ""
    text = re.sub(r"<[^>]+>", " ", str(value))
    return re.sub(r"\s+", " ", text).strip()


def norm(value) -> str:
    return clean(value).casefold()


def status_score(value) -> int:
    return STATUS_SCORE.get(norm(value), 0)


def status_observed(value) -> bool:
    return norm(value) in STATUS_SCORE


def community_governance_score(value) -> int:
    return {"formal": 100, "ad hoc": 50, "none": 0}.get(norm(value), 0)


def community_governance_observed(value) -> bool:
    return norm(value) in {"formal", "ad hoc", "none"}


def board_score(board_type, other_text):
    value = norm(board_type)
    other = norm(other_text)
    if value == "none":
        return 0, "Explicitly reports no board/steering body", False
    if value in {"advisory board or steering committee", "multi-board governance", "advisory board or steering committee | other"}:
        return 100, "Recognized board or steering-body category", False
    if value == "other":
        terms = ("board of directors", "foundation board", "advisory board", "steering committee", "governing board", "executive board", "board of trustees", "governance board")
        if any(term in other for term in terms):
            return 100, "Other field names a board/committee; review its relationship to the infrastructure", True
        return 0, "Other response needs review; no qualifying governance body established", True
    return 0, "Blank or unrecognized board-structure response", False


def load_source(path: Path, sheet_name: str):
    wb = openpyxl.load_workbook(path, data_only=True, keep_vba=path.suffix.lower() == ".xlsm")
    ws = wb[sheet_name]
    headers = [ws.cell(2, c).value for c in range(1, ws.max_column + 1)]
    records = []
    by_service = {}
    for r in range(DATA_START, ws.max_row + 1):
        service = clean(ws.cell(r, headers.index("001_service_name") + 1).value)
        if not service:
            continue
        record = {key: ws.cell(r, headers.index(key) + 1).value for key in headers if key}
        records.append(record)
        by_service[service] = record
    return wb, ws, records, by_service


def compute(record, community_record):
    scores = {code: status_score(record[field]) for code, field in STATUS_FIELDS.items()}
    g3, g3_reason, g3_review = board_score(record.get("050_board_structure"), record.get("950_board_structure_other"))
    g7 = community_governance_score(community_record.get("048_community_governance_status"))
    scores.update({"G3": g3, "G7": g7})
    observed = {code: status_observed(record[field]) for code, field in STATUS_FIELDS.items()}
    observed["G3"] = bool(clean(record.get("050_board_structure")))
    observed["G7"] = community_governance_observed(community_record.get("048_community_governance_status"))
    count = sum(observed.values())
    coverage = count / len(METRICS)
    observed_vals = [scores[code] for code, *_ in METRICS if observed[code]]
    observed_mean = sum(observed_vals) / len(observed_vals) if observed_vals else 0
    final = sum(scores[code] for code, *_ in METRICS) / len(METRICS)
    flags = []
    if g3_review:
        flags.append("Review 050 Other response; override available")
    for code, field in STATUS_FIELDS.items():
        if not status_observed(record.get(field)):
            flags.append(f"{field} unavailable/unknown scored 0")
    if not community_governance_observed(community_record.get("048_community_governance_status")):
        flags.append("048 community governance missing/unrecognized scored 0")
    if not clean(record.get("verification_status")) and not clean(record.get("verification_notes")):
        flags.append("No independent verification recorded")
    return {
        "scores": scores, "observed": observed, "count": count, "coverage": coverage,
        "observed_mean": observed_mean, "final": final, "g3_reason": g3_reason,
        "g3_review": g3_review, "flags": "; ".join(flags),
    }


def average_rank(values):
    indexed = sorted(enumerate(values), key=lambda pair: pair[1], reverse=True)
    ranks = [0.0] * len(values)
    start = 0
    while start < len(indexed):
        end = start + 1
        while end < len(indexed) and indexed[end][1] == indexed[start][1]:
            end += 1
        average = (start + 1 + end) / 2
        for pos in range(start, end):
            ranks[indexed[pos][0]] = average
        start = end
    return ranks


def spearman(x, y):
    if len(x) < 2:
        return None
    rx, ry = average_rank(x), average_rank(y)
    mx, my = statistics.mean(rx), statistics.mean(ry)
    denominator = math.sqrt(sum((a - mx) ** 2 for a in rx) * sum((b - my) ** 2 for b in ry))
    return sum((a - mx) * (b - my) for a, b in zip(rx, ry)) / denominator if denominator else 1.0


def rank_delta(reference, variant):
    ref_rank, var_rank = average_rank(reference), average_rank(variant)
    deltas = [abs(a - b) for a, b in zip(ref_rank, var_rank)]
    return statistics.mean(deltas) if deltas else 0, max(deltas, default=0), sum(delta >= 10 for delta in deltas)


def set_style(ws, widths=None):
    ws.freeze_panes = "A2"
    if ws.max_row > 1 and ws.max_column > 1:
        ws.auto_filter.ref = ws.dimensions
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor=GOV_COLORS["teal"])
        cell.font = Font(color=GOV_COLORS["white"], bold=True)
        cell.alignment = Alignment(vertical="top", wrap_text=True)
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    for col in range(1, ws.max_column + 1):
        if widths and col in widths:
            width = widths[col]
        else:
            longest = max((len(str(ws.cell(r, col).value or "")) for r in range(1, min(ws.max_row, 100) + 1)), default=12)
            width = min(max(longest + 2, 12), 48)
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.row_dimensions[1].height = 32


def add_formula_caches(path: Path, cache_map: dict[str, dict[str, float]]):
    ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
          "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
          "pkg": "http://schemas.openxmlformats.org/package/2006/relationships"}
    ET.register_namespace("", ns["main"])
    ET.register_namespace("r", ns["rel"])
    temp_path = None
    with ZipFile(path, "r") as original:
        workbook = ET.fromstring(original.read("xl/workbook.xml"))
        relationships = ET.fromstring(original.read("xl/_rels/workbook.xml.rels"))
        targets = {item.attrib["Id"]: item.attrib["Target"] for item in relationships.findall("pkg:Relationship", ns)}
        sheet_paths = {}
        for sheet in workbook.findall("main:sheets/main:sheet", ns):
            target = targets[sheet.attrib[f"{{{ns['rel']}}}id"]]
            sheet_paths[sheet.attrib["name"]] = target.lstrip("/") if target.startswith("/") else f"xl/{target}"
        replacements = {}
        for sheet_name, values in cache_map.items():
            xml_path = sheet_paths[sheet_name]
            root = ET.fromstring(original.read(xml_path))
            cells = {cell.attrib.get("r"): cell for cell in root.findall(".//main:c", ns)}
            for address, value in values.items():
                cell = cells[address]
                if cell.find("main:f", ns) is None:
                    raise ValueError(f"Expected formula in {sheet_name}!{address}")
                node = cell.find("main:v", ns)
                if node is None:
                    node = ET.SubElement(cell, f"{{{ns['main']}}}v")
                node.text = format(float(value), ".15g")
                cell.attrib.pop("t", None)
            replacements[xml_path] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        fd, temp_path = tempfile.mkstemp(suffix=".xlsx", dir=path.parent)
        os.close(fd)
        with ZipFile(temp_path, "w", ZIP_DEFLATED) as updated:
            for item in original.infolist():
                updated.writestr(item, replacements.get(item.filename, original.read(item.filename)))
    try:
        os.replace(temp_path, path)
    finally:
        if temp_path and os.path.exists(temp_path):
            os.remove(temp_path)


def main():
    gov_wb, gov_ws, records, gov_by_service = load_source(GOV_SOURCE, GOV_SHEET)
    ce_wb, ce_ws, _, ce_by_service = load_source(CE_SOURCE, CE_SHEET)
    if set(gov_by_service) != set(ce_by_service):
        raise ValueError("Governance and Community source service names do not match exactly.")
    community_records = {name: ce_by_service[name] for name in gov_by_service}
    computed = [compute(record, community_records[clean(record["001_service_name"])]) for record in records]
    n = len(records)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    sheet_names = ["Summary", "Scoring", "Metric Rules", "Methodology", "Parameters", "Diagnostics", "Sensitivity", "Coverage", "Verification", "Review Queue", "Calc", "Data", "Field Guide"]
    sheets = {name: wb.create_sheet(name) for name in sheet_names}

    rules = sheets["Metric Rules"]
    rules.append(["Code", "Metric", "Source field(s)", "Exact weight", "Displayed weight", "Scoring criteria", "Unavailable-data treatment", "Why this metric is retained"])
    reasons = {
        "G1": "Reports whether a governance structure is implemented or at an earlier stated stage; the data show responses across the status ladder.",
        "G2": "Captures written bylaws or equivalent rules as a distinct formal-governance element; report its comparatively high missingness.",
        "G3": "Captures whether a board/steering body exists. Corrects the former error that credited the explicit response None.",
        "G4": "Captures the reported maturity of web accessibility status.",
        "G5": "Captures the reported maturity of explicit equity commitments.",
        "G6": "Captures the reported maturity of privacy commitments relevant to users and their data.",
        "G7": "Measures formal versus ad hoc community participation in governance; all source service names match the governance dataset.",
    }
    criteria = {
        "G1": "Implemented/available=100; In progress=67; Considering=33; Not planning=0.",
        "G2": "Implemented/available=100; In progress=67; Considering=33; Not planning=0.",
        "G3": "Recognized advisory/steering or multi-board category=100; None=0; Other=100 only when field 950 substantiates a relevant board/committee; otherwise 0 pending review.",
        "G4": "Implemented/available=100; In progress=67; Considering=33; Not planning=0.",
        "G5": "Implemented/available=100; In progress=67; Considering=33; Not planning=0.",
        "G6": "Implemented/available=100; In progress=67; Considering=33; Not planning=0.",
        "G7": "Formal=100; Ad hoc=50; None=0.",
    }
    for code, title, fields in METRICS:
        unavailable = "Unknown, not applicable, blank, or unrecognized=0 in final score; omitted from observed-only diagnostics. Explicit negative response counts as observed zero."
        if code == "G3":
            unavailable = "Blank=0 and unavailable. Other is scored 0 provisionally and flagged unless field 950 supports a relevant body; reviewer override 0/100 available. None is an observed zero."
        if code == "G4":
            unavailable = "Blank or unrecognized=0 and unavailable. Explicit None counts as observed zero."
        rules.append([code, title, fields, 1 / 7, "14.29%", criteria[code], unavailable, reasons[code]])
    rules.append(["Total", "", "", 1.0, "100.00%", "", "", "Exact weight per metric is 1/7; 14.29% is display rounding."])

    params = sheets["Parameters"]
    params.append(["Parameter", "Value", "Explanation"])
    params_rows = [
        ["Number of metrics", 7, "G1 through G7"],
        ["Exact metric weight", 1 / 7, "Each metric contributes equally; exact weight sums to 100%."],
        ["Displayed metric weight", "14.29%", "Rounded display only."],
        ["Final score denominator", 7, "Fixed for every infrastructure."],
        ["Coverage eligibility threshold", "None", "Coverage is diagnostic only; all entries are scored."],
        ["Unavailable-data treatment", 0, "Unknown, N/A, blank, or unrecognized values contribute zero to the final score."],
        ["Status ladder", "100 / 67 / 33 / 0", "Implemented / In progress / Considering / Not planning."],
        ["Community governance ladder", "100 / 50 / 0", "Formal / Ad hoc / None."],
        ["050 Other handling", "Review + override", "Use 950 description; score 100 only when a relevant board or committee is substantiated."],
    ]
    for row in params_rows:
        params.append(row)

    method = sheets["Methodology"]
    method.append(["Methodology and interpretation", "Details"])
    method_rows = [
        ["Scope", "Seven indicators represent governance-structure status, formal rules, governing-body presence, community participation in governance, accessibility, equity, and privacy."],
        ["Equal weighting", "Each top-level metric receives exactly 1/7 of the score (14.2857%; displayed as 14.29%). Equal weighting is a transparent design choice, not an empirically estimated claim of equal causal importance."],
        ["Final formula", "Governance Practices Score = (G1 + G2 + G3 + G4 + G5 + G6 + G7) / 7. Each metric is scored from 0 to 100."],
        ["No coverage gate", "Every infrastructure receives a score. There is no minimum coverage or eligibility rule; the denominator stays seven."],
        ["Zero-fill assumption", "Unknown, Not applicable, blank, and unrecognized responses score zero in the final composite. This makes the score reflect reported performance and data completeness; zero does not prove absence of the underlying practice."],
        ["Coverage diagnostics", "Observed metric count and observed coverage (count/7) are reported but do not gate scoring. Explicit negative responses such as Not planning or None count as observed zero. Unknown, N/A, blank, and unresolved status values are excluded from observed coverage."],
        ["Observed-only mean", "Average the metric scores with observed data only. This is descriptive and uses a variable denominator; it is not the final score or primary zero-filled ranking."],
        ["048 source merge", "Community governance fields 048 and 049 are imported from the Community Engagement verification workbook and matched by exact service name. All 112 names match; the current Community Engagement index does not score 048 as a separate metric."],
        ["050 correction", "The prior binary rule awarded points to any nonblank value and therefore treated None as positive. The revised rule scores None=0 and requires field 950 evidence to count Other as a governing body. Other cases are flagged for review and may be overridden in Verification."],
        ["Fields excluded from score", "026 Governance records and 033 Transparent pricing are excluded as requested. Field 046 Values frameworks and 100 Board level remain descriptive; 046 contains heterogeneous frameworks and is blank for 76 entries, while 100 has only 15 responses. URLs and summaries are evidence/context, not additional metrics."],
        ["Why 033 is excluded", "The prior source has 59 No direct costs and 47 Implemented/available responses. Treating both as fully transparent pricing places 94.6% at the score ceiling and conflates absence of charge with price transparency."],
        ["Sensitivity analysis", "The zero-filled seven-metric score is the main result. Observed-only, complete-case, and six-metric results excluding G3 are included only as sensitivity checks; none exclude entries from the primary score."],
        ["Evidence limitation", "The source verification_status and verification_notes fields are empty for all entries. Scores are therefore based on recorded fields and have not been independently verified through every linked source."],
        ["Interpretation caution", "G1, G2, and G3 are related structural-governance indicators. Retaining all three gives the formal governance cluster 3/7 of the total metric weight. Sensitivity results show how much G3 influences rankings."],
        ["No bands", "No 70-point cutoff or high/medium/low performance bands are applied."],
    ]
    for row in method_rows:
        method.append(row)

    # Preserve raw governance source fields; append 048/049 joined by exact service name.
    data = sheets["Data"]
    data_fields = GOV_FIELDS + ["048_community_governance_status", "049_community_governance_url"]
    data.append(data_fields)
    for record in records:
        community = community_records[clean(record["001_service_name"])]
        data.append([record.get(field) for field in GOV_FIELDS] + [community.get("048_community_governance_status"), community.get("049_community_governance_url")])

    guide = sheets["Field Guide"]
    gov_guide = gov_wb["Field_Guide"]
    for row in gov_guide.iter_rows(values_only=True):
        guide.append(list(row))
    ce_guide = ce_wb["Field_Guide"]
    ce_guide_headers = [ce_guide.cell(1, c).value for c in range(1, ce_guide.max_column + 1)]
    for r in range(2, ce_guide.max_row + 1):
        code = ce_guide.cell(r, 1).value
        if code in {"048_community_governance_status", "049_community_governance_url"}:
            guide.append([ce_guide.cell(r, c).value for c in range(1, len(ce_guide_headers) + 1)])

    scoring = sheets["Scoring"]
    scoring.append(["Service", "Provider", "Solution category", "024 structure status", "108 bylaws status", "050 board structure", "950 other detail", "048 community governance", "029 accessibility status", "035 equity status", "044 privacy status", "G1", "G2", "G3", "G4", "G5", "G6", "G7", "Final score / 100", "Rank", "Available metrics (of 7)", "Observed coverage", "Observed-only mean", "Data-quality notes"])
    calc = sheets["Calc"]
    calc.append(["Service", "G1", "G2", "G3 provisional", "G4", "G5", "G6", "G7", "G3 coding rationale", "G3 review flag", "Python-checked zero-filled score", "Available metrics", "Observed coverage", "Observed-only mean", "Data-quality flags"])
    verification = sheets["Verification"]
    verification.append(["Service", "050 board structure", "950 other detail", "Provisional G3", "Reviewer override (0 or 100)", "Final G3", "Review status", "Reviewer notes"])
    coverage = sheets["Coverage"]
    coverage.append(["Service", "G1 observed", "G2 observed", "G3 observed", "G4 observed", "G5 observed", "G6 observed", "G7 observed", "Available metrics (of 7)", "Observed coverage", "Observed-only mean", "Zero-filled final score"])
    review = sheets["Review Queue"]
    review.append(["Service", "Metric / field", "Raw value", "Provisional treatment", "Review reason / action", "Reviewer notes"])

    status_codes = list(STATUS_FIELDS.keys())
    for excel_row, (record, result) in enumerate(zip(records, computed), start=2):
        service = clean(record["001_service_name"])
        community = community_records[service]
        g3, g3_reason, g3_review = board_score(record.get("050_board_structure"), record.get("950_board_structure_other"))
        review_status = "Review Other response" if g3_review else "No G3 review flagged"
        verification.append([service, record.get("050_board_structure"), record.get("950_board_structure_other"), g3, None, f'=IF(ISNUMBER(E{excel_row}),E{excel_row},D{excel_row})', review_status, "Allowed override values are 0 or 100; use only after checking the described body's role."])
        score_map = result["scores"]
        calc.append([service, score_map["G1"], score_map["G2"], score_map["G3"], score_map["G4"], score_map["G5"], score_map["G6"], score_map["G7"], result["g3_reason"], "Yes" if g3_review else "No", result["final"], result["count"], result["coverage"], result["observed_mean"], result["flags"]])
        coverage_flags = [int(result["observed"][code]) for code, *_ in METRICS]
        coverage.append([service, *coverage_flags, result["count"], result["coverage"], f"=Scoring!W{excel_row}", f"=Scoring!S{excel_row}"])
        raw_fields = [record.get(field) for field in STATUS_FIELDS.values()]
        scoring.append([
            service, record.get("008_provider_name"), record.get("014_solution_category"),
            record.get("024_governance_structure_status"), record.get("108_bylaws_status"),
            record.get("050_board_structure"), record.get("950_board_structure_other"),
            community.get("048_community_governance_status"), record.get("029_web_accessibility_status"),
            record.get("035_commitment_to_equity_status"), record.get("044_commitment_to_privacy_status"),
            f"=Calc!B{excel_row}", f"=Calc!C{excel_row}", f"=Verification!F{excel_row}",
            f"=Calc!E{excel_row}", f"=Calc!F{excel_row}", f"=Calc!G{excel_row}", f"=Calc!H{excel_row}",
            f"=AVERAGE(L{excel_row}:R{excel_row})",
            f'=RANK(S{excel_row},$S$2:$S${n+1},0)+(COUNTIF($S$2:$S${n+1},S{excel_row})-1)/2',
            f"=Coverage!I{excel_row}", f"=Coverage!J{excel_row}",
            f'=IFERROR(SUMPRODUCT(L{excel_row}:R{excel_row},Coverage!B{excel_row}:H{excel_row})/U{excel_row},0)', result["flags"],
        ])
        if g3_review:
            review.append([service, "G3: 050/950", f'{clean(record.get("050_board_structure"))} | {clean(record.get("950_board_structure_other"))}', g3, "Confirm the described body governs, oversees, or advises this infrastructure; then enter 0 or 100 on Verification.", ""])
        for code, status_field, url_field in [
            ("G1", "024_governance_structure_status", "025_governance_structure_url"),
            ("G4", "029_web_accessibility_status", "030_web_accessibility_url"),
            ("G5", "035_commitment_to_equity_status", "036_commitment_to_equity_url"),
            ("G6", "044_commitment_to_privacy_status", "045_commitment_to_privacy_url"),
        ]:
            if norm(record.get(status_field)) == "implemented/available" and not clean(record.get(url_field)):
                review.append([service, f"{code}: {status_field}", record.get(status_field), "Status scored 100; no paired URL supplied", f"Review official evidence where available; this is a documentation check, not an eligibility rule.", ""])
        if norm(community.get("048_community_governance_status")) == "formal" and not clean(community.get("049_community_governance_url")):
            review.append([service, "G7: 048/049", community.get("048_community_governance_status"), "Formal scored 100; no paired URL supplied", "Review official evidence where available; this is a documentation check, not an eligibility rule.", ""])

    # Metric diagnostics use observed responses only for the observed mean; zero-filled means use all records.
    diagnostics = sheets["Diagnostics"]
    diagnostics.append(["Metric", "Source field(s)", "Observed n", "Observed coverage", "Mean among observed", "Mean under zero-fill", "SD among observed", "% at 100 among observed", "% at 0 among observed", "Reading"])
    for code, title, fields in METRICS:
        observed_values = [result["scores"][code] for result in computed if result["observed"][code]]
        all_values = [result["scores"][code] for result in computed]
        observed_n = len(observed_values)
        missing_reading = {
            "G1": "Broad status variation; unknown values are zero-filled.",
            "G2": "29 unknown values; zero-fill materially affects its mean.",
            "G3": "Other values need review; explicit None correctly scores zero.",
            "G4": "Accessibility applicability field is sparse; use status as recorded.",
            "G5": "Equity commitment has good spread across status levels.",
            "G6": "High implemented share; still retains non-ceiling variation.",
            "G7": "Community governance is sourced from the matched CE sheet; six blanks.",
        }[code]
        diagnostics.append([
            f"{code} {title}", fields, observed_n, observed_n / n,
            sum(observed_values) / observed_n if observed_n else 0,
            sum(all_values) / n,
            statistics.pstdev(observed_values) if observed_n > 1 else 0,
            sum(v == 100 for v in observed_values) / observed_n if observed_n else 0,
            sum(v == 0 for v in observed_values) / observed_n if observed_n else 0,
            missing_reading,
        ])
    finals = [result["final"] for result in computed]
    diagnostics.append([])
    diagnostics.append(["Component summary", "Value"])
    diagnostics.append(["Entries scored", n])
    diagnostics.append(["Metrics", 7])
    diagnostics.append(["Weight per metric", "1/7 = 14.2857% (display 14.29%)"])
    diagnostics.append(["Mean zero-filled score", statistics.mean(finals)])
    diagnostics.append(["Median zero-filled score", statistics.median(finals)])
    diagnostics.append(["Score standard deviation", statistics.pstdev(finals)])
    diagnostics.append(["Minimum score", min(finals)])
    diagnostics.append(["Maximum score", max(finals)])
    diagnostics.append(["Mean observed coverage", statistics.mean(result["coverage"] for result in computed)])
    diagnostics.append(["All seven metrics observed", sum(result["count"] == 7 for result in computed)])
    diagnostics.append(["Verification data", "verification_status and verification_notes are blank across the source entries."])

    baseline = [result["final"] for result in computed]
    observed_only = [result["observed_mean"] for result in computed]
    complete_idx = [i for i, result in enumerate(computed) if result["count"] == len(METRICS)]
    complete_scores = [baseline[i] for i in complete_idx]
    no_g3 = [sum(result["scores"][code] for code in ("G1", "G2", "G4", "G5", "G6", "G7")) / 6 for result in computed]
    baseline_ranks = average_rank(baseline)
    sensitivity = sheets["Sensitivity"]
    sensitivity.append(["Sensitivity variant", "Entries", "Spearman rho vs baseline", "Mean absolute score change", "Max absolute score change", "Entries moving >=10 ranks", "Interpretation"])
    obs_delta = rank_delta(baseline, observed_only)
    sensitivity.append(["Baseline: seven-metric zero-filled", n, 1.0, 0, 0, 0, "Primary result; all entries scored."])
    sensitivity.append(["Observed-only mean", n, spearman(baseline, observed_only), statistics.mean(abs(a-b) for a,b in zip(baseline, observed_only)), max(abs(a-b) for a,b in zip(baseline, observed_only)), obs_delta[2], "Diagnostic only; variable denominator and partial coverage."])
    if complete_idx:
        base_complete = [baseline[i] for i in complete_idx]
        complete_rho = spearman(base_complete, complete_scores)
        sensitivity.append(["Complete-data subset, same seven metrics", len(complete_idx), complete_rho, 0, 0, 0, "Sensitivity only; does not exclude other entries from the main score."])
    else:
        sensitivity.append(["Complete-data subset, same seven metrics", 0, None, None, None, None, "No complete cases."])
    no_g3_rho = spearman(baseline, no_g3)
    no_g3_delta = rank_delta(baseline, no_g3)
    sensitivity.append(["Six-metric variant excluding G3", n, no_g3_rho, statistics.mean(abs(a-b) for a,b in zip(baseline, no_g3)), max(abs(a-b) for a,b in zip(baseline, no_g3)), no_g3_delta[2], "Sensitivity check for the influence of board-presence metric; not the official metric set."])
    sensitivity.append([])
    sensitivity.append(["Note", "G3 sensitivity uses provisional 050/950 coding. Refresh this workbook after reviewer overrides to update all result and sensitivity tables."])

    summary = sheets["Summary"]
    summary.append(["Governance Practices: Final Equal-Weight Results", "Value"])
    summary_rows = [
        ["Entries", n], ["Final metrics", "G1-G7"], ["Weight per metric", "1/7 = 14.29%"],
        ["Coverage gate", "None"], ["Unavailable data", "Zero in final score"],
        ["Mean score", "=AVERAGE(Scoring!S2:S%d)" % (n+1)],
        ["Median score", "=MEDIAN(Scoring!S2:S%d)" % (n+1)],
        ["Standard deviation", statistics.pstdev(finals)],
        ["Minimum score", "=MIN(Scoring!S2:S%d)" % (n+1)],
        ["Maximum score", "=MAX(Scoring!S2:S%d)" % (n+1)],
        ["Mean observed coverage", "=AVERAGE(Coverage!J2:J%d)" % (n+1)],
        ["Entries with all seven metrics observed", "=COUNTIF(Coverage!I2:I%d,7)" % (n+1)],
        ["Interpretation", "Final scores reflect reported values and missingness because unavailable values score zero. Coverage is diagnostic only."],
        ["Metric review", "See Metric Rules, Verification, Review Queue, Diagnostics, and Sensitivity."],
    ]
    for row in summary_rows:
        summary.append(row)

    # Formatting.
    for ws in sheets.values():
        set_style(ws)
    set_style(summary, {1: 42, 2: 110})
    set_style(scoring, {1: 37, 2: 28, 3: 32, 4: 25, 5: 20, 6: 32, 7: 55, 8: 24, 9: 26, 10: 22, 11: 22, **{c: 14 for c in range(12, 20)}, 20: 10, 21: 22, 22: 18, 23: 20, 24: 70})
    set_style(rules, {1: 10, 2: 38, 3: 42, 4: 14, 5: 16, 6: 88, 7: 85, 8: 82})
    set_style(method, {1: 30, 2: 125})
    set_style(params, {1: 33, 2: 40, 3: 95})
    set_style(diagnostics, {1: 42, 2: 44, 3: 14, 4: 18, 5: 23, 6: 23, 7: 21, 8: 25, 9: 25, 10: 86})
    set_style(sensitivity, {1: 48, 2: 14, 3: 22, 4: 25, 5: 24, 6: 24, 7: 90})
    set_style(coverage, {1: 37, **{c: 17 for c in range(2, 9)}, 9: 22, 10: 18, 11: 20, 12: 22})
    set_style(verification, {1: 38, 2: 35, 3: 58, 4: 18, 5: 25, 6: 18, 7: 25, 8: 76})
    set_style(review, {1: 38, 2: 32, 3: 76, 4: 38, 5: 88, 6: 70})
    set_style(calc, {1: 38, **{c: 18 for c in range(2, 9)}, 9: 76, 10: 20, 11: 23, 12: 20, 13: 20, 14: 22, 15: 90})
    set_style(data, {c: min(max(len(str(data.cell(1, c).value or "")) + 3, 18), 42) for c in range(1, data.max_column + 1)})
    set_style(guide, {1: 42, 2: 34, 3: 78, 4: 82})

    for r in range(2, n + 2):
        for c in range(12, 20):
            scoring.cell(r, c).number_format = "0.00"
        scoring.cell(r, 19).number_format = "0.00"
        scoring.cell(r, 22).number_format = "0.00%"
        scoring.cell(r, 23).number_format = "0.00"
        coverage.cell(r, 10).number_format = "0.00%"
        coverage.cell(r, 11).number_format = "0.00"
        coverage.cell(r, 12).number_format = "0.00"
        verification.cell(r, 4).number_format = "0.00"
        verification.cell(r, 6).number_format = "0.00"
    for r in range(2, 9):
        for c in (4, 8, 9):
            diagnostics.cell(r, c).number_format = "0.00%"
        for c in (5, 6, 7):
            diagnostics.cell(r, c).number_format = "0.00"
    override_dv = DataValidation(type="list", formula1='"0,100"', allow_blank=True)
    override_dv.error = "Enter 0 or 100 only."
    override_dv.errorTitle = "Invalid G3 override"
    sheets["Verification"].add_data_validation(override_dv)
    override_dv.add(f"E2:E{n+1}")
    for r in range(2, n + 2):
        cell = verification.cell(r, 5)
        cell.fill = PatternFill("solid", fgColor=GOV_COLORS["gold"])
        cell.comment = Comment("Optional reviewer correction after checking field 950 and evidence. Allowed values: 0 or 100.", "OISI")

    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcMode = "auto"
    wb.save(OUTPUT)

    # Cache formula results so workbook previews display populated results without requiring Excel recalculation.
    final_values = [item["final"] for item in computed]
    rank_values = average_rank(final_values)
    cache_map = {"Summary": {}, "Scoring": {}, "Verification": {}, "Coverage": {}}
    cache_map["Summary"].update({
        "B7": statistics.mean(final_values),
        "B8": statistics.median(final_values),
        "B10": min(final_values),
        "B11": max(final_values),
        "B12": statistics.mean(item["coverage"] for item in computed),
        "B13": sum(item["count"] == 7 for item in computed),
    })
    for r, (record, result) in enumerate(zip(records, computed), start=2):
        for col, code in zip("LMOPQR", ("G1", "G2", "G4", "G5", "G6", "G7")):
            cache_map["Scoring"][f"{col}{r}"] = result["scores"][code]
        cache_map["Scoring"][f"N{r}"] = result["scores"]["G3"]
        cache_map["Scoring"][f"S{r}"] = result["final"]
        cache_map["Scoring"][f"T{r}"] = rank_values[r - 2]
        cache_map["Scoring"][f"U{r}"] = result["count"]
        cache_map["Scoring"][f"V{r}"] = result["coverage"]
        cache_map["Scoring"][f"W{r}"] = result["observed_mean"]
        cache_map["Verification"][f"F{r}"] = result["scores"]["G3"]
        cache_map["Coverage"][f"K{r}"] = result["observed_mean"]
        cache_map["Coverage"][f"L{r}"] = result["final"]
    add_formula_caches(OUTPUT, cache_map)
    print(f"Created {OUTPUT} with {n} entries, seven equal-weight metrics, and {review.max_row-1} review items.")
    print(f"Mean zero-filled score: {statistics.mean(final_values):.4f}; mean observed coverage: {statistics.mean(item['coverage'] for item in computed):.4f}.")
    print(f"Complete-data sensitivity entries: {len(complete_idx)}; 048 exact-name join: {len(gov_by_service)} of {len(ce_by_service)} matched.")


if __name__ == "__main__":
    main()
