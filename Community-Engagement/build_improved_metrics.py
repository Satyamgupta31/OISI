from __future__ import annotations

import re
from collections import Counter
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


SOURCE = Path("OISI_Community_Engagement_v2_new_metrics.xlsx")
OUTPUT = Path("OISI_Community_Engagement_v3_multi_sheet.xlsx")

METRICS = [
    ("E1", "Engagement availability", "037_community_engagement_status", "Engagement & Participation", 0.10),
    ("E2", "Engagement activity breadth", "038 + coded 938", "Engagement & Participation", 0.25),
    ("E3", "Engagement evidence strength", "116 + 938", "Engagement & Participation", 0.05),
    ("C1", "Contribution opportunity breadth", "040 + 940", "Contribution Mechanisms", 0.25),
    ("C2", "Contribution guidance", "041", "Contribution Mechanisms", 0.15),
    ("C3", "Open contribution infrastructure", "019", "Contribution Mechanisms", 0.05),
    ("N1", "Community norms / Code of Conduct", "042 + 043", "Community Norms", 0.15),
]

STATUS_SCORE = {
    "implemented/available": 1.0,
    "in progress": 0.67,
    "considering": 0.33,
    "not planning": 0.0,
}

ENGAGEMENT_KEYWORDS = {
    "Annual meetings": ["annual meeting", "annual general meeting", "agm", "annual summit", "global summit"],
    "Blogs": ["blog", "community spotlight", "news post", "updates"],
    "Community calls": ["community call", "community conversation", "office hour", "calendar", "meeting", "meetings"],
    "Conference participation": ["conference", "workshop", "presentation", "symposium", "event"],
    "Development sprints": ["sprint", "hackathon", "code sprint"],
    "Interest, working, user, or advisory groups": [
        "working group",
        "interest group",
        "user group",
        "advisory",
        "council",
        "committee",
        "stakeholder",
        "shared responsibility",
        "roles",
        "community network",
    ],
    "Mailing lists and discussion forums (including Slack)": [
        "forum",
        "slack",
        "discourse",
        "mailing list",
        "google group",
        "listserv",
        "chat",
        "discussion",
    ],
    "Social media": ["social media", "twitter", "x.com", "mastodon", "linkedin"],
    "Staff roles with responsibility for community engagement": [
        "community manager",
        "community engagement manager",
        "engagement manager",
        "community coordinator",
        "staff",
    ],
    "User research": ["user research", "user testing", "feedback", "survey", "consultation", "interview"],
    "Volunteer or ambassador network": ["volunteer", "ambassador", "champion", "affiliate", "advocate"],
    "Webinars and training": ["webinar", "training", "school", "onboarding", "education", "documentation hub"],
}

SUSPECT_GUIDANCE = ["discord", "terms_of_use", "account-types", "bute_code", ";"]
PUBLIC_FORGES = ["github.com", "gitlab.com", "bitbucket.org", "codeberg.org"]

DATA_COLUMNS = [
    "001_service_name",
    "008_provider_name",
    "010_website",
    "013_location",
    "014_solution_category",
    "019_open_code_repository_url",
    "037_community_engagement_status",
    "038_community_engagement_activities",
    "938_community_engagement_activities_other",
    "116_community_engagement_url",
    "039_contribution_pathways",
    "040_contribution_opportunities",
    "940_contribution_opportunities_other",
    "041_contribution_guidelines",
    "042_code_of_conduct_status",
    "043_code_of_conduct_url",
    "078_membership_numbers",
    "107_membership_program_url",
]


def clean(value) -> str:
    if value is None:
        return ""
    text = re.sub(r"<[^>]+>", " ", str(value))
    return re.sub(r"\s+", " ", text).strip()


def lower(value) -> str:
    return clean(value).lower()


def split_pipe(value) -> list[str]:
    text = clean(value)
    if not text:
        return []
    return [part.strip() for part in text.split("|") if part.strip()]


def status_score(value):
    return STATUS_SCORE.get(lower(value), None)


def engagement_categories(checklist, other_text):
    checklist_categories = set()
    for item in split_pipe(checklist):
        if item.lower() not in {"other", "not applicable"}:
            checklist_categories.add(item)

    coded_categories = set()
    text = lower(other_text)
    if text:
        for category, keywords in ENGAGEMENT_KEYWORDS.items():
            if any(keyword in text for keyword in keywords):
                coded_categories.add(category)

    categories = checklist_categories | coded_categories
    if "other" in lower(checklist) and text:
        categories.add("Other")
    return checklist_categories, coded_categories, categories


def contribution_count(checklist, other_text):
    items = [item for item in split_pipe(checklist) if item.lower() not in {"other", "not applicable"}]
    if clean(other_text):
        items.append("Other")
    return len(set(items))


def guidance_score(url):
    text = lower(url)
    if not text:
        return None, "Blank"
    if any(pattern in text for pattern in SUSPECT_GUIDANCE):
        return None, "CHECK: suspect URL"
    if any(pattern in text for pattern in ["contribut", "get-involved", "giveortake", "newcomers"]):
        return 1.0, "Contribution guide"
    return 0.5, "Other participation page"


def repo_score(url):
    text = lower(url)
    if not text:
        return None
    return 1.0 if any(host in text for host in PUBLIC_FORGES) else 0.5


def coc_score(status, url):
    base = status_score(status)
    if base is None:
        return None
    if base == 1.0 and not clean(url):
        return 0.75
    return base


def score_row(row):
    checklist_cats, coded_938, all_cats = engagement_categories(
        row["038_community_engagement_activities"],
        row["938_community_engagement_activities_other"],
    )
    c1_count = contribution_count(row["040_contribution_opportunities"], row["940_contribution_opportunities_other"])
    c2, c2_class = guidance_score(row["041_contribution_guidelines"])

    scores = {
        "E1": status_score(row["037_community_engagement_status"]),
        "E2": min(len(all_cats), 8) / 8 if all_cats else None,
        "E3": 1.0
        if clean(row["116_community_engagement_url"])
        else (0.5 if clean(row["938_community_engagement_activities_other"]) else None),
        "C1": min(c1_count, 6) / 6 if c1_count else None,
        "C2": c2,
        "C3": repo_score(row["019_open_code_repository_url"]),
        "N1": coc_score(row["042_code_of_conduct_status"], row["043_code_of_conduct_url"]),
    }

    weight_scored = sum(weight for code, _, _, _, weight in METRICS if scores[code] is not None)
    weighted_sum = sum(scores[code] * weight for code, _, _, _, weight in METRICS if scores[code] is not None)
    groups = {group for code, _, _, group, _ in METRICS if scores[code] is not None}
    eligible = weight_scored >= 0.60 and len(groups) >= 3
    final = weighted_sum / weight_scored * 100 if eligible else None

    flags = []
    if clean(row["938_community_engagement_activities_other"]) and not coded_938:
        flags.append("938 text present but not coded")
    if c2_class == "CHECK: suspect URL":
        flags.append("041 suspect URL")
    if lower(row["039_contribution_pathways"]) in {"considering", "not planning"} and (
        c1_count > 0 or clean(row["041_contribution_guidelines"])
    ):
        flags.append("039 conflicts with contribution evidence")
    if not clean(row["019_open_code_repository_url"]):
        flags.append("No code repo URL")

    meta = {
        "checklist_count": len(checklist_cats),
        "coded_938_count": len(coded_938),
        "union_count": len(all_cats),
        "categories": "; ".join(sorted(all_cats)),
        "c1_count": c1_count,
        "c2_class": c2_class,
        "weight_scored": weight_scored,
        "weighted_sum": weighted_sum,
        "metrics_scored": sum(1 for v in scores.values() if v is not None),
        "groups": len(groups),
        "eligible": eligible,
        "score": final,
        "lower": weighted_sum * 100,
        "upper": (weighted_sum + (1 - weight_scored)) * 100,
        "status": "Reportable" if eligible and weight_scored >= 0.85 else ("Provisional" if eligible else "Not scored"),
        "flags": "; ".join(flags),
        "coded_categories": "; ".join(sorted(coded_938)),
    }
    return scores, meta


def append_rows(ws, rows):
    for row in rows:
        ws.append(row)


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


def median(values):
    if not values:
        return None
    ordered = sorted(values)
    mid = len(ordered) // 2
    if len(ordered) % 2:
        return ordered[mid]
    return (ordered[mid - 1] + ordered[mid]) / 2


def main():
    source_wb = openpyxl.load_workbook(SOURCE, data_only=True)
    source_data = source_wb["Data"]
    source_headers = [source_data.cell(1, c).value for c in range(1, source_data.max_column + 1)]

    rows = []
    for r in range(2, source_data.max_row + 1):
        source_row = {h: source_data.cell(r, i + 1).value for i, h in enumerate(source_headers)}
        row = {h: source_row.get(h) for h in DATA_COLUMNS}
        scores, meta = score_row(row)
        rows.append((row, scores, meta))

    ranked = sorted([x for x in rows if x[2]["score"] is not None], key=lambda x: x[2]["score"], reverse=True)
    ranks = {row["001_service_name"]: i for i, (row, _, _) in enumerate(ranked, 1)}
    scores_list = [meta["score"] for _, _, meta in rows if meta["score"] is not None]
    statuses = Counter(meta["status"] for _, _, meta in rows)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    sheets = {
        name: wb.create_sheet(name)
        for name in [
            "Codebook",
            "Summary",
            "Scoring",
            "Verification",
            "Coding",
            "Calc",
            "Sensitivity",
            "Descriptive",
            "Audit",
            "Dictionary",
            "Parameters",
            "Data",
        ]
    }

    append_rows(
        sheets["Codebook"],
        [
            ["OISI Component 1 - Community Engagement v3: codebook and method"],
            ["Dataset", "Infra Finder community-engagement fields copied from the supplied workbook."],
            ["Score formula", "100 x sum(weight x metric score) / sum(scored weights), after missing-data exclusion."],
            ["Missing data rule", "Unknown, blank, and Not applicable are excluded from numerator and denominator."],
            ["Completeness rule", "Score requires at least 60% metric weight and at least one scored metric in each sub-component."],
            ["Important change", "Community governance is not part of this score. The old governance metric has been removed."],
            ["039 treatment", "039 is not scored because the source field has many Unknown values; it is used only for consistency flags."],
            [],
            ["Metric", "Definition"],
            *[[code, name, fields, group, weight] for code, name, fields, group, weight in METRICS],
        ],
    )

    two_i2c = next(meta for row, _, meta in rows if row["001_service_name"] == "2i2c")
    append_rows(
        sheets["Summary"],
        [
            ["OISI Component 1 - Community Engagement v3 results summary"],
            ["Metric set", "E1, E2, E3, C1, C2, C3, N1"],
            ["Total services", len(rows)],
            ["Services scored", len(scores_list)],
            ["Reportable", statuses["Reportable"]],
            ["Provisional", statuses["Provisional"]],
            ["Not scored", statuses["Not scored"]],
            ["Median score", median(scores_list)],
            ["Mean score", sum(scores_list) / len(scores_list) if scores_list else None],
            ["2i2c score", two_i2c["score"]],
            ["2i2c status", two_i2c["status"]],
            ["2i2c coded 938 categories", two_i2c["coded_categories"]],
        ],
    )

    scoring_headers = [
        "Service",
        "Provider",
        "Solution category",
        *[f"{code} {name}" for code, name, *_ in METRICS],
        "Engagement & Participation (0-100)",
        "Contribution Mechanisms (0-100)",
        "Community Norms (0-100)",
        "Weight scored (share)",
        "Weighted sum",
        "Metrics scored (of 7)",
        "Sub-components covered (of 3)",
        "Eligible for score?",
        "COMMUNITY ENGAGEMENT SCORE (0-100)",
        "Lower bound (missing metrics = 0)",
        "Upper bound (missing metrics = 1)",
        "Result status",
        "Rank",
        "Data-quality flags",
    ]
    sheets["Scoring"].append(scoring_headers)
    for row, scores, meta in rows:
        engagement = subscore(scores, ["E1", "E2", "E3"], [0.10, 0.25, 0.05])
        contribution = subscore(scores, ["C1", "C2", "C3"], [0.25, 0.15, 0.05])
        norms = subscore(scores, ["N1"], [0.15])
        sheets["Scoring"].append(
            [
                row["001_service_name"],
                row["008_provider_name"],
                row["014_solution_category"],
                *[scores[code] for code, *_ in METRICS],
                engagement,
                contribution,
                norms,
                meta["weight_scored"],
                meta["weighted_sum"],
                meta["metrics_scored"],
                meta["groups"],
                "Yes" if meta["eligible"] else "No",
                meta["score"],
                meta["lower"],
                meta["upper"],
                meta["status"],
                ranks.get(row["001_service_name"], "NA"),
                meta["flags"],
            ]
        )

    sheets["Verification"].append(
        ["Service", "E1 override", "E2 override", "E3 override", "C1 override", "C2 override", "C3 override", "N1 override", "Reviewer notes"]
    )
    for row, _, _ in rows:
        sheets["Verification"].append([row["001_service_name"], "", "", "", "", "", "", "", ""])

    category_names = list(ENGAGEMENT_KEYWORDS)
    coding_headers = (
        ["Service"]
        + [f"038: {c}" for c in category_names]
        + [f"938: {c}" for c in category_names]
        + ["Checklist categories (n)", "Coded 938 categories (n)", "Union categories (n)", "Union categories"]
    )
    sheets["Coding"].append(coding_headers)
    for row, _, meta in rows:
        checklist, coded, union = engagement_categories(row["038_community_engagement_activities"], row["938_community_engagement_activities_other"])
        sheets["Coding"].append(
            [row["001_service_name"]]
            + [1 if c in checklist else 0 for c in category_names]
            + [1 if c in coded else 0 for c in category_names]
            + [len(checklist), len(coded), len(union), "; ".join(sorted(union))]
        )

    sheets["Calc"].append(
        [
            "Service",
            "037 score",
            "E2 union count",
            "E2 score",
            "E3 evidence score",
            "040/940 count",
            "C1 score",
            "041 URL class",
            "C2 score",
            "C3 repo score",
            "N1 CoC score",
            "Flags",
        ]
    )
    for row, scores, meta in rows:
        sheets["Calc"].append(
            [
                row["001_service_name"],
                scores["E1"],
                meta["union_count"],
                scores["E2"],
                scores["E3"],
                meta["c1_count"],
                scores["C1"],
                meta["c2_class"],
                scores["C2"],
                scores["C3"],
                scores["N1"],
                meta["flags"],
            ]
        )

    sheets["Sensitivity"].append(["Service", "Base score", "Threshold 50%", "Threshold 70%", "No 938 in E2"])
    for row, scores, meta in rows:
        no_938_scores = dict(scores)
        checklist, _, _ = engagement_categories(row["038_community_engagement_activities"], "")
        no_938_scores["E2"] = min(len(checklist), 8) / 8 if checklist else None
        sheets["Sensitivity"].append(
            [
                row["001_service_name"],
                meta["score"],
                recompute(no_938_scores, 0.50),
                recompute(scores, 0.70),
                recompute(no_938_scores, 0.60),
            ]
        )

    sheets["Descriptive"].append(
        ["Service", "Provider", "078 Membership numbers", "107 Membership program URL", "Membership evidence", "Community Engagement score"]
    )
    for row, _, meta in rows:
        member_number = row["078_membership_numbers"]
        membership_evidence = "Yes" if clean(member_number) or clean(row["107_membership_program_url"]) else "No"
        sheets["Descriptive"].append(
            [
                row["001_service_name"],
                row["008_provider_name"],
                member_number,
                row["107_membership_program_url"],
                membership_evidence,
                meta["score"],
            ]
        )

    sheets["Audit"].append(["Service", "938 text extract", "Coded categories", "Reviewer verdict", "Reviewer notes"])
    audit_rows = sorted(
        [x for x in rows if clean(x[0]["938_community_engagement_activities_other"])],
        key=lambda x: x[0]["001_service_name"] or "",
    )[:30]
    for row, _, meta in audit_rows:
        sheets["Audit"].append(
            [
                row["001_service_name"],
                clean(row["938_community_engagement_activities_other"])[:900],
                meta["coded_categories"],
                "",
                "",
            ]
        )

    sheets["Dictionary"].append(["Category", "Keyword"])
    for category, keywords in ENGAGEMENT_KEYWORDS.items():
        for keyword in keywords:
            sheets["Dictionary"].append([category, keyword])

    sheets["Parameters"].append(["Code", "Metric", "Variable", "Sub-component", "Weight", "Note"])
    for code, name, fields, group, weight in METRICS:
        sheets["Parameters"].append([code, name, fields, group, weight, "v3 scoring metric"])
    sheets["Parameters"].append(["", "", "", "Total", sum(w for *_, w in METRICS), "Must equal 100%"])
    sheets["Parameters"].append([])
    sheets["Parameters"].append(["Setting", "Value", "Note"])
    sheets["Parameters"].append(["Completeness threshold", 0.60, "Minimum scored metric weight required"])
    sheets["Parameters"].append(["Reportable threshold", 0.85, "High-completeness reporting band"])
    sheets["Parameters"].append(["Activity cap", 8, "E2 categories for full score"])
    sheets["Parameters"].append(["Opportunity cap", 6, "C1 categories for full score"])

    sheets["Data"].append(DATA_COLUMNS)
    for row, _, _ in rows:
        sheets["Data"].append([row[h] for h in DATA_COLUMNS])

    for sheet in sheets.values():
        style_sheet(sheet)
    for ws_name in ["Scoring", "Calc", "Sensitivity", "Descriptive", "Summary", "Parameters"]:
        for row in sheets[ws_name].iter_rows():
            for cell in row:
                if isinstance(cell.value, float):
                    cell.number_format = "0.00"

    wb.save(OUTPUT)


def subscore(scores, codes, weights):
    scored = [(scores[c], w) for c, w in zip(codes, weights) if scores[c] is not None]
    if not scored:
        return "NA"
    return sum(v * w for v, w in scored) / sum(w for _, w in scored) * 100


def recompute(scores, threshold):
    weight_by_code = {code: weight for code, _, _, _, weight in METRICS}
    group_by_code = {code: group for code, _, _, group, _ in METRICS}
    weight_scored = sum(weight_by_code[c] for c, value in scores.items() if value is not None)
    weighted_sum = sum(value * weight_by_code[c] for c, value in scores.items() if value is not None)
    groups = {group_by_code[c] for c, value in scores.items() if value is not None}
    if weight_scored >= threshold and len(groups) >= 3:
        return weighted_sum / weight_scored * 100
    return "NA"


if __name__ == "__main__":
    main()
