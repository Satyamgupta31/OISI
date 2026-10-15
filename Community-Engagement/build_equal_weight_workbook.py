from __future__ import annotations

import importlib.util
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation


HERE = Path(__file__).resolve().parent
SOURCE = HERE / "OISI_Community_Engagement_v2_new_metrics.xlsx"
OUTPUT = HERE / "OISI_Community_Engagement_Equal_Weight_All_Scored_Missing_Zero.xlsx"

spec = importlib.util.spec_from_file_location("ce_previous", HERE / "build_improved_metrics.py")
ce = importlib.util.module_from_spec(spec)
assert spec and spec.loader
spec.loader.exec_module(ce)

METRICS = [
    ("E1", "Engagement status", "037", "Engagement & Participation"),
    ("E2", "Engagement activity breadth", "038 + reviewed/coded 938", "Engagement & Participation"),
    ("C1", "Contribution opportunity breadth", "040 + reviewed 940", "Contribution Mechanisms"),
    ("C2", "Contribution guidance", "041", "Contribution Mechanisms"),
    ("C3", "Public contribution repository", "019", "Contribution Mechanisms"),
    ("N1", "Code of Conduct status", "042; 043 evidence only", "Community Norms"),
]
WEIGHT = 1 / len(METRICS)
ACTIVITY_DENOMINATOR = 13
CONTRIBUTION_DENOMINATOR = 9
STATUS = {"implemented/available": 100.0, "in progress": 67.0, "considering": 33.0, "not planning": 0.0}
PUBLIC_FORGES = ("github.com", "gitlab.com", "bitbucket.org", "codeberg.org", "sr.ht")
SUSPECT_GUIDANCE = ce.SUSPECT_GUIDANCE


def number(score):
    return None if score is None else round(float(score), 4)


def status_score(value):
    return STATUS.get(ce.lower(value))


def guidance_candidate(url):
    text = ce.lower(url)
    if not text or any(pattern in text for pattern in SUSPECT_GUIDANCE):
        return None, "Blank or suspect URL; review required"
    if any(pattern in text for pattern in ("contribut", "get-involved", "giveortake", "newcomers")):
        return 100.0, "Candidate dedicated contribution guide; verify URL"
    return 50.0, "Candidate general participation page; verify URL"


def repository_candidate(url):
    text = ce.lower(url)
    if not text:
        return None, "No repository URL supplied"
    if any(host in text for host in PUBLIC_FORGES):
        return 100.0, "Available public-forge repository URL"
    return 50.0, "Repository field populated on another/unspecified host"


def row_scores(row):
    checklist, coded, union = ce.engagement_categories(
        row.get("038_community_engagement_activities"),
        row.get("938_community_engagement_activities_other"),
    )
    has_other_938 = "other" in ce.lower(row.get("038_community_engagement_activities")) and bool(
        ce.clean(row.get("938_community_engagement_activities_other"))
    )
    activity_count = len(union)
    activity_count = min(activity_count, ACTIVITY_DENOMINATOR)

    contribution_items = {
        item for item in ce.split_pipe(row.get("040_contribution_opportunities"))
        if item.lower() not in {"other", "not applicable", "unknown"}
    }
    other_text = ce.clean(row.get("940_contribution_opportunities_other"))
    has_other_route = (
        "other" in {x.lower() for x in ce.split_pipe(row.get("040_contribution_opportunities"))}
        and bool(other_text)
    )
    contribution_count = min(len(contribution_items) + int(has_other_route), CONTRIBUTION_DENOMINATOR)

    guidance, guidance_note = guidance_candidate(row.get("041_contribution_guidelines"))
    repository, repository_note = repository_candidate(row.get("019_open_code_repository_url"))
    e2 = 100.0 * activity_count / ACTIVITY_DENOMINATOR if activity_count else None
    c1 = 100.0 * contribution_count / CONTRIBUTION_DENOMINATOR if contribution_count else None
    scores = [
        status_score(row.get("037_community_engagement_status")),
        e2,
        c1,
        guidance,
        repository,
        status_score(row.get("042_code_of_conduct_status")),
    ]
    notes = {
        "activity_count": activity_count,
        "activity_types": "; ".join(sorted(union | ({"Other"} if has_other_938 else set()))),
        "activity_coded_count": len(coded),
        "contribution_count": contribution_count,
        "contribution_types": "; ".join(sorted(contribution_items | ({"Other"} if has_other_route else set()))),
        "guidance_note": guidance_note,
        "repository_note": repository_note,
        "guidance_candidate_score": guidance,
        "repository_candidate_score": repository,
    }
    flags = []
    if ce.clean(row.get("938_community_engagement_activities_other")):
        flags.append("938 coding requires reviewer confirmation")
    if ce.clean(row.get("940_contribution_opportunities_other")):
        flags.append("940 Other route requires reviewer confirmation")
    if ce.clean(row.get("041_contribution_guidelines")):
        flags.append("041 scored from available URL tier")
    if ce.clean(row.get("019_open_code_repository_url")):
        flags.append("019 scored from populated repository field")
    if ce.clean(row.get("043_code_of_conduct_url")):
        flags.append("043 evidence URL should be checked; not scored separately")
    if ce.clean(row.get("116_community_engagement_url")):
        flags.append("116 retained as source evidence only; not scored")
    if ce.lower(row.get("037_community_engagement_status")) in {"considering", "not planning"} and activity_count:
        flags.append("037 status conflicts with listed 038/938 activity; both scored independently")
    missing_count = sum(v is None for v in scores)
    if missing_count:
        flags.append(f"{missing_count} unavailable metric(s) scored as 0")
    return scores, notes, "; ".join(flags)


def style(ws, widths=None):
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions if ws.max_row > 1 else "A1"
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor="245B55")
        cell.font = Font(color="FFFFFF", bold=True)
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    for col in range(1, ws.max_column + 1):
        vals = [len(str(ws.cell(r, col).value or "")) for r in range(1, min(ws.max_row, 80) + 1)]
        ws.column_dimensions[get_column_letter(col)].width = (widths or {}).get(col, min(max(max(vals, default=10) + 2, 12), 48))


def main():
    source_wb = openpyxl.load_workbook(SOURCE, data_only=True, read_only=True)
    src = source_wb["Data"]
    headers = [c.value for c in src[1]]
    records = []
    for cells in src.iter_rows(min_row=2, values_only=True):
        raw = dict(zip(headers, cells))
        records.append(raw)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    names = ["Summary", "Scoring", "Metric Rules", "Parameters", "Verification", "Activity Coding", "Contribution Coding", "Calc", "Review Queue", "Data", "Formula Checks"]
    sheets = {name: wb.create_sheet(name) for name in names}

    rules = [
        ["OISI Community Analysis / Community Engagement: Equal-Weight Scoring Rules"],
        ["Scoring set", "Six metrics E1, E2, C1, C2, C3, N1. Each is equally weighted."],
        ["Weights", "Each metric = 1/6 = 16.6667%; displayed as 16.67%. Calculations use the exact 1/6 fraction."],
        ["Final score", "Arithmetic mean of all six metric scores on a 0-100 scale."],
        ["Coverage rule", "No coverage eligibility gate. Every infrastructure receives a score."],
        ["Missing values", "Unavailable, blank, Unknown, or ambiguous metric values are scored as 0. This zero-filled score is not proof of zero underlying performance."],
        ["Performance bands", "None are used."],
        ["038 and 938", "Twelve named 038 activity types plus a thirteenth Other category, only when Other is selected and 938 supports a distinct unlisted activity. Duplicate categories are counted once."],
        ["040 and 940", "Eight named 040 contribution routes plus a ninth Other category, only when Other is selected and 940 supports a distinct unlisted route."],
        ["116", "Evidence only; not a scored metric. It is not included in E2 counts."],
        ["042 and 043", "042 is one status metric. 043 is supporting evidence for reviewer verification and is not scored separately."],
        ["Review status", "C2 and C3 are scored automatically when their source fields are populated. URL review notes are audit context, not a scoring prerequisite. Free-text coding and Other categories remain reviewable."],
        [],
        ["Code", "Metric", "Source fields", "Equal weight", "Scoring rule", "Missing / verification treatment"],
        ["E1", "Engagement status", "037", WEIGHT, "Implemented/available=100; In progress=67; Considering=33; Not planning=0.", "Unknown/blank=0."],
        ["E2", "Engagement activity breadth", "038 + reviewed/coded 938", WEIGHT, "100 x distinct supported activity types / 13. Twelve named types plus valid Other; categories deduplicated.", "No supported type=0. Review 938 coding and Other where needed."],
        ["C1", "Contribution opportunity breadth", "040 + reviewed 940", WEIGHT, "100 x distinct supported routes / 9. Eight named routes plus valid Other; categories deduplicated.", "No supported route=0. Review 940 Other where needed."],
        ["C2", "Contribution guidance", "041", WEIGHT, "Available dedicated contribution-guide URL=100; available relevant general participation URL=50.", "Blank or suspect URL=0."],
        ["C3", "Public contribution repository", "019", WEIGHT, "Available public-forge URL=100; populated URL on another/unspecified host=50.", "Blank field=0."],
        ["N1", "Code of Conduct status", "042; 043 evidence only", WEIGHT, "Implemented/available=100; In progress=67; Considering=33; Not planning=0.", "Unknown/blank=0. Use 043 as evidence only; no separate URL bonus/penalty."],
    ]
    for row in rules:
        sheets["Metric Rules"].append(row)

    param = sheets["Parameters"]
    param.append(["Code", "Metric", "Source fields", "Weight (exact)", "Weight (display)", "Dimension"])
    for code, name, fields, dimension in METRICS:
        param.append([code, name, fields, WEIGHT, "16.67%", dimension])
    param.append(["Total", "", "", sum([WEIGHT] * 6), "100.00%", ""])
    param.append([])
    param.append(["Breadth denominator", "Available indicators", "Count rule"])
    param.append(["038", 13, "12 named types + valid Other; no duplicate category count"])
    param.append(["040", 9, "8 named routes + valid Other; no duplicate route count"])
    param.append(["Coverage eligibility gate", "None", "All infrastructures receive an equal-weight score; unavailable metrics contribute 0"])

    coding = sheets["Activity Coding"]
    activity_categories = list(ce.ENGAGEMENT_KEYWORDS)
    coding.append(["Service", *[f"038: {c}" for c in activity_categories], *[f"938 coded: {c}" for c in activity_categories], "038 count", "938 coded count", "Other supported", "Union count / 13", "Coded union"])
    contrib = sheets["Contribution Coding"]
    routes = ["Contribute funds", "Contribute to code", "Contribute to documentation", "Contribute to education or training", "Contribute to governance", "Contribute to peer reviews", "Contribute to user research or user testing", "Contribute to working groups or interest groups"]
    contrib.append(["Service", *[f"040: {r}" for r in routes], "940 text", "Other selected and described", "Route count / 9", "Coded routes"])

    calc = sheets["Calc"]
    calc.append(["Service", "E1 default", "E2 preliminary", "C1 preliminary", "C2 URL score", "C3 repository-field score", "N1 default", "E2 count", "C1 count", "C2 scoring note", "C3 scoring note", "Data flags"])
    verification = sheets["Verification"]
    verification.append(["Service", "E1 override (0-100)", "E2 override (0-100)", "C1 override (0-100)", "C2 override (0-100)", "C3 override (0-100)", "N1 override (0-100)", "938 reviewed?", "940 reviewed?", "041 verified?", "019 verified public repo?", "043 checked?", "Reviewer notes"])
    dv = DataValidation(type="decimal", operator="between", formula1="0", formula2="100", allow_blank=True)
    verification.add_data_validation(dv)

    review = sheets["Review Queue"]
    review.append(["Service", "Review item", "Source value / coded proposal", "Reviewer decision", "Reviewer notes"])
    scoring = sheets["Scoring"]
    scoring.append(["Service", "Provider", "Solution category", "E1", "E2", "C1", "C2", "C3", "N1", "Metrics with source data (of 6)", "Community Analysis score (0-100; missing=0)", "Missing metrics scored as 0", "Review status", "Data-quality flags"])

    for ix, row in enumerate(records, start=2):
        service = row.get("001_service_name") or ""
        scores, note, flags = row_scores(row)
        checklist, coded, union = ce.engagement_categories(row.get("038_community_engagement_activities"), row.get("938_community_engagement_activities_other"))
        other_support = "Other" in ce.lower(row.get("038_community_engagement_activities")) and bool(ce.clean(row.get("938_community_engagement_activities_other")))
        coded_names = union | ({"Other"} if other_support else set())
        coding.append([service, *[int(c in checklist) for c in activity_categories], *[int(c in coded) for c in activity_categories], len(checklist), len(coded), int(other_support), len(coded_names), "; ".join(sorted(coded_names))])

        selected_routes = {x for x in ce.split_pipe(row.get("040_contribution_opportunities")) if x.lower() not in {"other", "not applicable", "unknown"}}
        other_route = "other" in {x.lower() for x in ce.split_pipe(row.get("040_contribution_opportunities"))} and bool(ce.clean(row.get("940_contribution_opportunities_other")))
        contrib.append([service, *[int(r in selected_routes) for r in routes], row.get("940_contribution_opportunities_other"), int(other_route), note["contribution_count"], "; ".join(sorted(selected_routes | ({"Other"} if other_route else set())))])

        calc.append([service, *[number(x) for x in [scores[0], scores[1], scores[2], scores[3], scores[4], scores[5]]], note["activity_count"], note["contribution_count"], f'{note["guidance_candidate_score"]}: {note["guidance_note"]}', f'{note["repository_candidate_score"]}: {note["repository_note"]}', flags])
        verification.append([service, None, None, None, None, None, None, "Review coding" if ce.clean(row.get("938_community_engagement_activities_other")) else "N/A", "Review Other" if ce.clean(row.get("940_contribution_opportunities_other")) else "N/A", "Scored from available URL" if scores[3] is not None else "No score", "Scored from available field" if scores[4] is not None else "No score", "Evidence only" if ce.clean(row.get("043_code_of_conduct_url")) else "No URL", None])
        # Reviewer numeric overrides remain optional; otherwise source-derived values score directly.
        score_row_num = ix
        calc_row_num = ix
        metric_formulas = [
            f'=IF(ISNUMBER(Verification!{col}{ix}),Verification!{col}{ix},IF(Calc!{calc_col}{calc_row_num}="",0,Calc!{calc_col}{calc_row_num}))'
            for col, calc_col in zip("BCDEFG", "BCDEFG")
        ]
        observed_formula = f'=COUNT(Calc!B{calc_row_num}:G{calc_row_num})'
        final_formula = f'=AVERAGE(D{score_row_num}:I{score_row_num})'
        missing_formula = f'=6-J{score_row_num}'
        review_formula = f'=IF(COUNTIF(Verification!H{ix}:L{ix},"Review*")>0,"Scored - coding review","Scored")'
        scoring.append([service, row.get("008_provider_name"), row.get("014_solution_category"), *metric_formulas,
                        observed_formula, final_formula, missing_formula, review_formula, flags])

        if ce.clean(row.get("938_community_engagement_activities_other")):
            review.append([service, "938 free-text activity coding", row.get("938_community_engagement_activities_other"), "; ".join(sorted(coded)), "Confirm or correct categories; add valid Other only if distinct from the 12 named types."])
        if ce.clean(row.get("940_contribution_opportunities_other")):
            review.append([service, "940 Other contribution route", row.get("940_contribution_opportunities_other"), "Other" if other_route else "Not counted", "Confirm this is selected in 040 and describes a distinct route."])
        if ce.clean(row.get("041_contribution_guidelines")):
            review.append([service, "041 contribution guidance URL", row.get("041_contribution_guidelines"), f'{note["guidance_candidate_score"]}: {note["guidance_note"]}', "Scored from the available URL; inspect only if the URL appears suspect or the automatic tier is inaccurate."])
        if ce.clean(row.get("019_open_code_repository_url")):
            review.append([service, "019 public repository", row.get("019_open_code_repository_url"), f'{note["repository_candidate_score"]}: {note["repository_note"]}', "Scored from the populated repository field; inspect only if the URL or host classification is clearly malformed."])
        if ce.clean(row.get("043_code_of_conduct_url")):
            review.append([service, "043 Code of Conduct evidence", row.get("043_code_of_conduct_url"), "Evidence only", "Check whether the linked policy applies to this community; do not score URL separately."])

    # Formula-driven results are recalculated by Excel/LibreOffice when opened.
    summary = sheets["Summary"]
    summary.append(["Community Analysis equal-metric scoring summary"])
    summary.append(["Services in source snapshot", len(records)])
    summary.append(["Metric set", "E1, E2, C1, C2, C3, N1; six equally weighted metrics"])
    summary.append(["Weight per metric", "1/6 = 16.6667% (displayed 16.67%)"])
    summary.append(["Scored services", '=COUNT(Scoring!K2:K%d)' % (len(records) + 1)])
    summary.append(["Mean score", '=IFERROR(AVERAGE(Scoring!K2:K%d),"")' % (len(records) + 1)])
    summary.append(["Scoring rule", "Equal arithmetic mean of all six metrics; every unavailable metric is scored as 0. No coverage eligibility gate."])
    summary.append(["Interpretation", "Zero-filled scores are conservative evidence-adjusted results; zero due to unavailable data does not establish zero underlying performance."])

    raw = sheets["Data"]
    raw.append(headers)
    for record in records:
        raw.append([record.get(h) for h in headers])

    checks = sheets["Formula Checks"]
    checks.append(["Check", "Expected", "Workbook value / rule"])
    checks.append(["Number of metrics", 6, len(METRICS)])
    checks.append(["Total exact weight", 1, sum([WEIGHT] * 6)])
    checks.append(["038 denominator", 13, "12 named activity types + valid Other"])
    checks.append(["040 denominator", 9, "8 named contribution routes + valid Other"])
    checks.append(["Standalone 116 metric", 0, "Removed; evidence only"])
    checks.append(["048 community governance metric", 0, "Not included"])
    checks.append(["Performance bands", 0, "None"])

    for r in range(2, len(records) + 2):
        dv.add(f"B{r}:G{r}")
    for ws in sheets.values():
        style(ws)
    widths = {
        "Metric Rules": {1: 18, 2: 45, 3: 30, 4: 18, 5: 75, 6: 70},
        "Scoring": {1: 32, 2: 32, 3: 35, 13: 35, 14: 70},
        "Verification": {1: 32, 13: 55},
        "Review Queue": {1: 32, 2: 34, 3: 70, 4: 60, 5: 75},
        "Data": {i + 1: 36 for i, _ in enumerate(headers)},
    }
    for name, ws in sheets.items():
        # Apply targeted widths after general autosizing.
        for col, width in widths.get(name, {}).items():
            ws.column_dimensions[get_column_letter(col)].width = width
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcMode = "auto"
    wb.save(OUTPUT)
    print(f"Created {OUTPUT}")
    print(f"Source records: {len(records)}; metrics: {len(METRICS)}; review items: {review.max_row - 1}")


if __name__ == "__main__":
    main()
