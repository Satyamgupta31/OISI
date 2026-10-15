from __future__ import annotations

import csv
import importlib.util
import math
import os
import statistics
import tempfile
import xml.etree.ElementTree as ET
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


HERE = Path(__file__).resolve().parent
OUTPUT = HERE / "OISI_Development_Activity_Final_Equal_Weight_Zero_Fill_Corrected.xlsx"
INPUTS = {
    "Repository register": HERE / "repos.csv",
    "Git metrics": HERE / "git_metrics.csv",
    "GitHub/API metrics": HERE / "gh_metrics.csv",
    "Manual metrics": HERE / "manual_metrics.csv",
    "Shared scoring rules": HERE / "score.py",
    "Extraction rules": HERE / "rules.py",
}

sys_path = str(HERE)
import sys
if sys_path not in sys.path:
    sys.path.insert(0, sys_path)
import rules as pipeline_rules

spec = importlib.util.spec_from_file_location("oisi_c4_score", HERE / "score.py")
score_module = importlib.util.module_from_spec(spec)
assert spec.loader is not None
spec.loader.exec_module(score_module)

METRIC_INFO = [
    ("m1_commits", "Human commits", "Continuity", "git_metrics.csv::human_commits", "Non-merge, non-bot commits during the window.", "100 * min(1, log10(1+c) / log10(1001)); 1,000 or more commits = 100.", "Measures recorded development volume. Log scaling limits domination by very large repositories."),
    ("m2_active_weeks", "Active weeks", "Continuity", "git_metrics.csv::active_weeks", "Distinct weeks, out of 52, with at least one human non-merge commit.", "100 * active_weeks / 52.", "Measures temporal spread of work rather than total volume."),
    ("m3_recency", "Commit recency", "Continuity", "git_metrics.csv::days_since_last", "Days from the latest human non-merge commit in the window to the cutoff.", "100 if d <= 30; otherwise max(0, 100 * (365-d)/335). No activity in a complete Git extract = 0.", "Distinguishes recent recorded development from activity that ended earlier in the window."),
    ("m4_releases", "Release activity", "Continuity", "git_metrics.csv::tags_in_window", "Number of Git tags created during the window; tags are a proxy for releases.", "0 tags = 0; 1 = 50; 2-3 = 75; 4+ = 100.", "Captures versioned outputs; projects that release without Git tags may be undercounted."),
    ("m5_contributors", "Active contributors", "Contributors", "git_metrics.csv::contributors", "Resolved human identities with at least one non-merge commit in the window.", "0-1 contributors = 0; otherwise 100 * min(1, log2(n)/5); 32+ = 100.", "Measures breadth of human participation, beyond commit volume alone."),
    ("m6_effective_contributors", "Effective contributors", "Contributors", "git_metrics.csv::eff_contributors", "Inverse concentration of human commit shares: E = 1 / sum(p_i^2).", "0 if E <= 1; otherwise 100 * min(1, log2(E)/3); E >= 8 = 100.", "Adjusts contributor count for whether work is concentrated in a small number of people."),
    ("m7_retention", "Contributor retention", "Contributors", "git_metrics.csv::retained", "Number of contributors active in both halves of the observation window.", "0 = 0; 1 = 25; 2 = 50; 3-4 = 75; 5+ = 100.", "Captures continued participation across the year rather than one-time contribution."),
    ("m8_issue_response", "Issue response within 14 days", "Responsiveness", "gh_metrics.csv::m8_issue_response_rate; manual_metrics.csv override", "Share of eligible issues receiving a qualifying response within 14 days.", "100 * response rate; at least 10 eligible issues required for an observed rate.", "Measures timely handling of reported issues. A response is not the same as resolution."),
    ("m9_pr_response", "PR response within 7 days", "Responsiveness", "gh_metrics.csv::m9_pr_response_rate", "Share of eligible non-bot pull requests receiving a qualifying response within 7 days.", "100 * response rate; at least 10 eligible PRs required for an observed rate.", "Measures timely engagement with proposed contributions."),
    ("m10_issue_close90", "Issue closure within 90 days", "Responsiveness", "gh_metrics.csv::m10_issue_close90_rate; manual_metrics.csv override", "Share of eligible human-authored issues closed within 90 days.", "100 * closure rate; at least 10 eligible issues required for an observed rate.", "Measures follow-through on issues, distinct from an initial response."),
    ("m11_pr_decision30", "PR decision within 30 days", "Responsiveness", "gh_metrics.csv::m11_pr_decision30_rate", "Share of eligible non-bot PRs merged or closed within 30 days.", "100 * decision rate; at least 10 eligible PRs required for an observed rate.", "Measures whether proposed changes reach a decision, not only an initial response."),
    ("m12_stars", "Stars gained", "Adoption", "gh_metrics.csv::m12_stars_gained", "New GitHub stars recorded inside the window.", "100 * min(1, log10(1+stars) / log10(1001)); 1,000+ = 100.", "A window-specific attention/adoption signal, not a direct quality measure; GitHub-only coverage."),
    ("m13_forks", "Forks created", "Adoption", "gh_metrics.csv::m13_forks_created", "GitHub forks created inside the window.", "100 * min(1, log10(1+forks) / log10(101)); 100+ = 100.", "A possible code-reuse signal distinct from stars, but correlated with them and GitHub-specific."),
]
METRIC_KEYS = [x[0] for x in METRIC_INFO]
DIMENSIONS = ["Continuity", "Contributors", "Responsiveness", "Adoption"]
WEIGHT = 1 / len(METRIC_KEYS)
COLORS = {"green": "245B55", "green_light": "E8F1EF", "blue": "315B70", "gold": "FCECC9", "red": "9A3D3D", "dark": "263238", "white": "FFFFFF", "line": "D5DFDC", "gray": "EEF1F2"}


def read_csv(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as stream:
        return list(csv.DictReader(stream))


def index_by_id(rows):
    return {row["oisi_id"]: row for row in rows}


def clean(value):
    return "" if value is None else str(value).strip()


def avg_rank(values):
    ordered = sorted(values, reverse=True)
    ranks = {}
    i = 0
    while i < len(ordered):
        j = i
        while j + 1 < len(ordered) and ordered[j + 1] == ordered[i]:
            j += 1
        ranks[ordered[i]] = (i + 1 + j + 1) / 2
        i = j + 1
    return [ranks[value] for value in values]


def spearman(a, b):
    if len(a) < 2:
        return 0.0
    ra, rb = avg_rank(a), avg_rank(b)
    ma, mb = statistics.mean(ra), statistics.mean(rb)
    num = sum((x-ma)*(y-mb) for x, y in zip(ra, rb))
    den = math.sqrt(sum((x-ma)**2 for x in ra) * sum((y-mb)**2 for y in rb))
    return num / den if den else 0.0


def quality_notes(repo, git, gh, manual, observed):
    notes = []
    if repo.get("scored") != "yes":
        notes.append("Repository register marks entry not assessable/excluded; zero-filled score is not evidence of inactivity")
    if git.get("git_status") != "complete":
        notes.append("Git extraction incomplete or unavailable")
    if gh.get("gh_status") != "complete":
        notes.append("GitHub/API extraction incomplete or no GitHub source")
    for code, nfield in (("m8_issue_response", "m8_n"), ("m9_pr_response", "m9_n"), ("m10_issue_close90", "m10_n"), ("m11_pr_decision30", "m11_n")):
        if observed[code]:
            continue
        raw_n = score_module.num(gh.get(nfield))
        if raw_n is not None and raw_n < pipeline_rules.MIN_ITEMS:
            notes.append(f"{code} below minimum eligible sample (n={int(raw_n)}); scored 0")
        elif manual.get("m8_issue_response_rate" if code == "m8_issue_response" else "m10_issue_close90_rate"):
            continue
        else:
            notes.append(f"{code} unavailable; scored 0")
    if not observed["m12_stars"] or not observed["m13_forks"]:
        notes.append("GitHub adoption fields unavailable where no supported GitHub source exists; scored 0")
    return "; ".join(dict.fromkeys(notes)) or "No missing metric inputs flagged by pipeline availability checks"


def write_formula_caches(path: Path, caches):
    ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
          "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
          "pkg": "http://schemas.openxmlformats.org/package/2006/relationships"}
    ET.register_namespace("", ns["main"])
    ET.register_namespace("r", ns["rel"])
    replacements = {}
    fd, temp_path = tempfile.mkstemp(suffix=".xlsx", dir=path.parent)
    os.close(fd)
    try:
        with ZipFile(path, "r") as source_zip:
            workbook_root = ET.fromstring(source_zip.read("xl/workbook.xml"))
            rel_root = ET.fromstring(source_zip.read("xl/_rels/workbook.xml.rels"))
            targets = {r.attrib["Id"]: r.attrib["Target"] for r in rel_root.findall("pkg:Relationship", ns)}
            sheet_paths = {}
            for sheet in workbook_root.findall("main:sheets/main:sheet", ns):
                target = targets[sheet.attrib[f"{{{ns['rel']}}}id"]]
                sheet_paths[sheet.attrib["name"]] = target.lstrip("/") if target.startswith("/") else f"xl/{target}"
            for sheet_name, values in caches.items():
                xml_path = sheet_paths[sheet_name]
                root = ET.fromstring(source_zip.read(xml_path))
                cells = {cell.attrib.get("r"): cell for cell in root.findall(".//main:c", ns)}
                for address, value in values.items():
                    cell = cells[address]
                    if cell.find("main:f", ns) is None:
                        raise ValueError(f"Expected formula at {sheet_name}!{address}")
                    value_node = cell.find("main:v", ns)
                    if value_node is None:
                        value_node = ET.SubElement(cell, f"{{{ns['main']}}}v")
                    value_node.text = format(float(value), ".15g")
                    cell.attrib.pop("t", None)
                replacements[xml_path] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            with ZipFile(temp_path, "w", ZIP_DEFLATED) as target_zip:
                for item in source_zip.infolist():
                    target_zip.writestr(item, replacements.get(item.filename, source_zip.read(item.filename)))
        os.replace(temp_path, path)
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)


def style_sheet(ws, widths=None, freeze="A2", tab=None):
    ws.freeze_panes = freeze
    if ws.max_row > 1 and ws.max_column > 1:
        ws.auto_filter.ref = ws.dimensions
    if tab:
        ws.sheet_properties.tabColor = tab
    thin = Side(style="hair", color=COLORS["line"])
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor=COLORS["green"])
        cell.font = Font(color=COLORS["white"], bold=True)
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if cell.row > 1:
                cell.border = Border(bottom=thin)
    for col in range(1, ws.max_column + 1):
        if widths and col in widths:
            width = widths[col]
        else:
            samples = [len(str(ws.cell(r, col).value or "")) for r in range(1, min(ws.max_row, 100) + 1)]
            width = min(max(max(samples, default=10) + 2, 12), 44)
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.row_dimensions[1].height = 34


def main():
    repos = read_csv(INPUTS["Repository register"])
    git_rows = read_csv(INPUTS["Git metrics"])
    gh_rows = read_csv(INPUTS["GitHub/API metrics"])
    manual_rows = read_csv(INPUTS["Manual metrics"])
    git_by_id, gh_by_id, manual_by_id = map(index_by_id, (git_rows, gh_rows, manual_rows))
    n = len(repos)
    if n != 112:
        raise ValueError(f"Expected 112 register rows; found {n}")
    if len({r["oisi_id"] for r in repos}) != n:
        raise ValueError("Duplicate oisi_id in repository register")

    records = []
    for repo in repos:
        oid = repo["oisi_id"]
        git = git_by_id.get(oid, {})
        gh = gh_by_id.get(oid, {})
        manual = manual_by_id.get(oid, {})
        raw = {**git, **gh, **{k: v for k, v in manual.items() if clean(v)}}
        scored = score_module.score_row(raw)
        metric_scores = {key: (0.0 if scored[key] is None else float(scored[key])) for key in METRIC_KEYS}
        observed = {key: scored[key] is not None for key in METRIC_KEYS}
        available_count = sum(observed.values())
        coverage = available_count / len(METRIC_KEYS)
        observed_vals = [float(scored[key]) for key in METRIC_KEYS if observed[key]]
        observed_mean = statistics.mean(observed_vals) if observed_vals else 0.0
        final = statistics.mean(metric_scores.values())
        records.append({
            "repo": repo, "git": git, "gh": gh, "manual": manual, "raw": raw,
            "scores": metric_scores, "observed": observed, "available_count": available_count,
            "coverage": coverage, "observed_mean": observed_mean, "final": final,
        })

    finals = [r["final"] for r in records]
    base_ranks = avg_rank(finals)
    scenarios = {"Baseline: 13 equal metrics, zero-filled": finals}
    observed_only = [r["observed_mean"] for r in records]
    scenarios["Sensitivity: observed-only mean"] = observed_only
    for key in METRIC_KEYS:
        scenarios[f"Sensitivity: omit {key}"] = [sum(v for code, v in r["scores"].items() if code != key) / 12 for r in records]
    scenario_ranks = {name: avg_rank(values) for name, values in scenarios.items()}

    wb = Workbook()
    wb.remove(wb.active)
    names = ["Summary", "Scoring", "Metric Rules", "Methodology", "Parameters", "Diagnostics", "Sensitivity", "Sensitivity Ranks", "Correlations", "Coverage", "Verification", "Calc", "Data", "Repository Register", "Extraction Rules", "Data Dictionary"]
    sheets = {name: wb.create_sheet(name) for name in names}
    tab_colors = [COLORS["green"], COLORS["green"], COLORS["red"], COLORS["blue"], COLORS["red"], COLORS["gray"], COLORS["gray"], COLORS["gray"], COLORS["gray"], COLORS["green_light"], COLORS["gold"], COLORS["gray"], COLORS["gray"], COLORS["gray"], COLORS["blue"], COLORS["gray"]]
    for name, color in zip(names, tab_colors):
        sheets[name].sheet_properties.tabColor = color

    scoring = sheets["Scoring"]
    scoring_headers = ["OISI ID", "Infrastructure", "Classification", "Repository status", "Primary repo(s)", "Git status", "GitHub/API status", *METRIC_KEYS, "Final score / 100", "Rank (average ties)", "Available metrics (of 13)", "Observed coverage", "Observed-only mean", "Assessment interpretation", "Data-quality notes"]
    scoring.append(scoring_headers)
    calc = sheets["Calc"]
    calc.append(["OISI ID", *METRIC_KEYS, "Python-checked final score", "Available metrics", "Observed coverage", "Observed-only mean"])
    coverage_ws = sheets["Coverage"]
    coverage_ws.append(["OISI ID", "Infrastructure", *[f"{key} observed" for key in METRIC_KEYS], "Available metrics (of 13)", "Observed coverage", "Observed-only mean", "Zero-filled final score", "Coverage interpretation"])

    for row_num, item in enumerate(records, start=2):
        repo, git, gh, manual = item["repo"], item["git"], item["gh"], item["manual"]
        scores = item["scores"]
        score_ref_start = 8
        calc.append([repo["oisi_id"], *[scores[k] for k in METRIC_KEYS], item["final"], item["available_count"], item["coverage"], item["observed_mean"]])
        flags = [int(item["observed"][k]) for k in METRIC_KEYS]
        coverage_ws.append([
            repo["oisi_id"], repo["infrastructure"], *flags,
            f"=SUM(C{row_num}:O{row_num})", f"=P{row_num}/13",
            f"=IFERROR(SUMPRODUCT(Calc!B{row_num}:N{row_num},C{row_num}:O{row_num})/P{row_num},0)",
            f"=Scoring!U{row_num}",
            "Diagnostic only; never an eligibility rule.",
        ])
        notes = quality_notes(repo, git, gh, manual, item["observed"])
        if repo.get("scored") == "yes":
            assessment = "Repository marked assessable in register"
        else:
            assessment = "NOT ASSESSABLE / EXCLUDED: numeric zero is a policy zero-fill, not measured inactivity"
        metric_formulas = [f"=Calc!{get_column_letter(score_ref_start+i)}{row_num}" for i in range(13)]
        scoring.append([
            repo["oisi_id"], repo["infrastructure"], repo.get("classification"), repo.get("decision"), repo.get("repo"),
            git.get("git_status"), gh.get("gh_status"), *metric_formulas,
            f"=AVERAGE(H{row_num}:T{row_num})",
            f"=RANK(U{row_num},$U$2:$U${n+1},0)+(COUNTIF($U$2:$U${n+1},U{row_num})-1)/2",
            f"=Coverage!P{row_num}", f"=Coverage!Q{row_num}", f"=Coverage!R{row_num}", assessment, notes,
        ])

    rules_ws = sheets["Metric Rules"]
    rules_ws.append(["Code", "Metric", "Dimension label", "Source field", "Exact weight", "Displayed weight", "Definition", "Scoring rule", "Unavailable-data rule", "Why included / interpretation"])
    for key, title, dim, source, definition, formula, rationale in METRIC_INFO:
        unavailable = "Scores 0 in the fixed-denominator final score; not counted as observed coverage. A valid observed zero remains observed."
        if key in {"m8_issue_response", "m9_pr_response", "m10_issue_close90", "m11_pr_decision30"}:
            unavailable += f" Rate is unavailable below n >= {pipeline_rules.MIN_ITEMS}; that does not exclude the infrastructure from scoring."
        rules_ws.append([key, title, dim, source, WEIGHT, "7.69%", definition, formula, unavailable, rationale])
    rules_ws.append(["TOTAL", "", "", "", 1.0, "100.00%", "", "All 13 normalized metric scores averaged with a fixed denominator of 13.", "No overall coverage or eligibility threshold.", "Equal metric weighting, not equal dimension weighting."])

    method = sheets["Methodology"]
    method.append(["Methodology and interpretation", "Details"])
    method_rows = [
        ["Purpose", "Summarize recorded software-development activity and repository-level interaction/adoption signals for the OISI Development Activity component."],
        ["Metric set", "The 13 established indicators m1-m13 are retained; no new metric is introduced in this workbook."],
        ["Equal weighting", "Each metric contributes exactly 1/13 = 7.6923% (display 7.69%). This is an explicit transparent choice, not an empirical estimate of causal importance."],
        ["Dimension labels", "Dimensions are descriptive only. Equal metric weights imply Continuity=4/13 (30.77%), Contributors=3/13 (23.08%), Responsiveness=4/13 (30.77%), Adoption=2/13 (15.38%). They are not separately reweighted."],
        ["Final formula", "Development Activity Score = (m1 + m2 + ... + m13) / 13, where every metric is normalized to 0-100 before averaging."],
        ["Missing data", "Unknown, unavailable, unsupported-platform, failed extraction, or insufficient sample-size values score 0 in the final score. The denominator remains 13 for every entry."],
        ["No coverage gate", "No minimum observed-metric count, weight-coverage threshold, or dimension eligibility threshold is applied. Every one of the 112 register entries receives a numeric zero-filled result."],
        ["Metric-level sample rule", f"Responsiveness rates m8-m11 retain the pipeline's minimum of {pipeline_rules.MIN_ITEMS} eligible records to avoid highly unstable rates. Below that sample, the metric is unavailable and contributes 0; the infrastructure still receives a score."],
        ["Observed coverage", "Available metric count / 13. A valid measured zero counts as observed; unavailable values are not observed. Coverage is descriptive only and never gates the score."],
        ["Observed-only mean", "Mean of available metric scores for that entry. It is reported as a diagnostic with a variable denominator and is not the final score or ranking."],
        ["Assessability status", "Ten entries are marked scored=no in the current repository register (non-assessable, standards, or unresolved repository choice). They still receive numeric 0 under the requested fixed-denominator zero-fill rule, but remain visibly flagged and should not be interpreted as measured low activity or compared as ordinary software repositories."],
        ["Data provenance", "Scores are recalculated from current repos.csv, git_metrics.csv, gh_metrics.csv, and manual_metrics.csv, joined by oisi_id. The older weighted scores.csv and its 75%/dimension eligibility rule are not used for the new score."],
        ["Observation window", "2025-10-01 00:00 UTC inclusive to 2026-10-01 00:00 UTC exclusive."],
        ["Common Git rules", "Default branch unless overridden; author date; non-merge commits for authorship metrics; bots filtered using shared rules; identities resolved within repository/product set using the documented pipeline rules."],
        ["Responsiveness events", "Issue response counts qualifying non-author/non-bot comments, labels, assignments, or closes. PR response counts qualifying comments/reviews and non-author merges/closes. Full follow-up windows are required."],
        ["Correlation caution", "The current pairwise results show strong associations (e.g., commits-active weeks rho=.92; contributor count-effective contributors rho=.919; stars-forks rho=.852). Equal metric weights therefore do not imply independent constructs; the Correlations sheet is diagnostic."],
        ["Sensitivity", "The workbook compares the fixed-denominator baseline with observed-only averaging and 13 leave-one-metric-out variants. These are robustness diagnostics, not alternative primary scores."],
        ["Selection / measurement caveat", "GitHub-derived responsiveness, stars, and forks do not cover every tracker/platform equally. Manual metric fields are currently blank. Zero-filling therefore combines observed performance with data availability and platform coverage."],
        ["Interpretation", "The score summarizes these operational indicators only. It is not a direct measure of software quality, impact, security, user satisfaction, or long-term sustainability."],
    ]
    for row in method_rows:
        method.append(row)

    params = sheets["Parameters"]
    params.append(["Parameter", "Value", "Explanation"])
    param_rows = [
        ["Number of metrics", 13, "m1 through m13"],
        ["Exact metric weight", WEIGHT, "Each metric contributes exactly 1/13; exact weights sum to 100%."],
        ["Displayed metric weight", "7.69%", "Rounded presentation of 7.6923%."],
        ["Final score denominator", 13, "Fixed for every entry, including entries with missing fields."],
        ["Overall coverage eligibility threshold", "None", "No entry is excluded for low observed coverage."],
        ["Unavailable-data treatment", 0, "Unavailable or insufficient-sample metrics score zero in the final composite."],
        ["Observed coverage denominator", 13, "Available metric count / 13; diagnostic only."],
        ["Responsiveness minimum sample", pipeline_rules.MIN_ITEMS, "Minimum eligible n for m8-m11 rates. Under-threshold indicators score zero; entry remains scored."],
        ["Observation window start", pipeline_rules.WINDOW_START.isoformat(), "UTC inclusive."],
        ["Observation window end", pipeline_rules.WINDOW_END.isoformat(), "UTC exclusive."],
        ["Tie ranking", "Average rank", "Tied entries receive the average of their occupied ranks."],
    ]
    for row in param_rows:
        params.append(row)

    # Per-metric descriptive statistics over the full 112-entry scoring cohort.
    diagnostics = sheets["Diagnostics"]
    diagnostics.append(["Code", "Metric", "Dimension", "Source field", "Observed n", "Observed coverage", "Mean among observed", "Mean under zero-fill", "SD among observed", "% zero among observed", "Availability note"])
    for key, title, dim, source, definition, formula, rationale in METRIC_INFO:
        obs_vals = [r["scores"][key] for r in records if r["observed"][key]]
        zero_vals = [r["scores"][key] for r in records]
        availability_note = f"{len(obs_vals)} of {n} entries have a usable value."
        if key in {"m8_issue_response", "m9_pr_response", "m10_issue_close90", "m11_pr_decision30"}:
            availability_note += f" Requires at least {pipeline_rules.MIN_ITEMS} eligible items."
        if key in {"m12_stars", "m13_forks"}:
            availability_note += " GitHub-derived; no-source entries are unavailable."
        diagnostics.append([
            key, title, dim, source, len(obs_vals), len(obs_vals) / n,
            statistics.mean(obs_vals) if obs_vals else 0.0,
            statistics.mean(zero_vals), statistics.pstdev(obs_vals) if len(obs_vals) > 1 else 0.0,
            sum(v == 0 for v in obs_vals) / len(obs_vals) if obs_vals else 0.0,
            availability_note,
        ])
    diagnostics.append([])
    diagnostics.append(["Cohort statistic", "Value"])
    diagnostics.append(["Entries", n])
    diagnostics.append(["Mean zero-filled Development Activity score", statistics.mean(finals)])
    diagnostics.append(["Median zero-filled Development Activity score", statistics.median(finals)])
    diagnostics.append(["Mean observed metric coverage", statistics.mean(r["coverage"] for r in records)])
    diagnostics.append(["Entries with all 13 metrics observed", sum(r["available_count"] == 13 for r in records)])
    diagnostics.append(["Entries with zero observed metrics", sum(r["available_count"] == 0 for r in records)])
    diagnostics.append(["Register marked assessable", sum(r["repo"].get("scored") == "yes" for r in records)])
    diagnostics.append(["Register marked not assessable/excluded", sum(r["repo"].get("scored") != "yes" for r in records)])

    # Leave-one-out sensitivity and observed-only comparison.
    sens = sheets["Sensitivity"]
    sens.append(["Scenario", "Metrics averaged", "Entries", "Spearman rho vs baseline", "Mean absolute score change", "Maximum absolute score change", "Mean absolute rank shift", "Maximum absolute rank shift", "Top-10 overlap", "Interpretation"])
    top10_base = set(sorted(range(n), key=lambda i: (-finals[i], records[i]["repo"]["oisi_id"]))[:10])
    for scenario_name, values in scenarios.items():
        ranks = scenario_ranks[scenario_name]
        top10 = set(sorted(range(n), key=lambda i: (-values[i], records[i]["repo"]["oisi_id"]))[:10])
        if scenario_name.startswith("Baseline"):
            sens.append([scenario_name, 13, n, 1.0, 0.0, 0.0, 0.0, 0.0, 10, "Primary fixed-denominator result."])
        else:
            metric_count = 0 if scenario_name.endswith("observed-only mean") else 12
            changes = [abs(a-b) for a, b in zip(finals, values)]
            shifts = [abs(a-b) for a, b in zip(base_ranks, ranks)]
            label = "Variable-denominator diagnostic; not the primary score." if metric_count == 0 else "Leave-one-metric-out robustness check; not the primary score."
            sens.append([scenario_name, metric_count, n, spearman(finals, values), statistics.mean(changes), max(changes), statistics.mean(shifts), max(shifts), len(top10_base & top10), label])

    sr = sheets["Sensitivity Ranks"]
    sr.append(["OISI ID", "Infrastructure", "Baseline score", "Baseline rank", "Observed-only score", "Observed-only rank", *[f"Rank without {key}" for key in METRIC_KEYS]])
    for i, item in enumerate(records):
        sr.append([item["repo"]["oisi_id"], item["repo"]["infrastructure"], finals[i], base_ranks[i], observed_only[i], scenario_ranks["Sensitivity: observed-only mean"][i], *[scenario_ranks[f"Sensitivity: omit {key}"][i] for key in METRIC_KEYS]])

    # Pairwise correlations on jointly observed metric values; include pair n.
    corr = sheets["Correlations"]
    corr.append(["Metric A", "Metric B", "Pairwise observed n", "Spearman rho", "Interpretation"])
    for i, a in enumerate(METRIC_KEYS):
        for b in METRIC_KEYS[i+1:]:
            pairs = [(r["scores"][a], r["scores"][b]) for r in records if r["observed"][a] and r["observed"][b]]
            rho = spearman([x for x, _ in pairs], [y for _, y in pairs]) if len(pairs) > 1 else 0.0
            interp = "Strong association; interpret equal weights as indicator weights, not independent construct weights." if abs(rho) >= 0.7 else "Exploratory pairwise association; not evidence of causality."
            corr.append([a, b, len(pairs), rho, interp])

    # Verification and source-level provenance.
    verification = sheets["Verification"]
    verification.append(["OISI ID", "Infrastructure", "Scored in register", "Classification", "Decision", "Host", "Primary repo(s)", "Issue tracker", "Git status", "GitHub/API status", "Manual m8/m10 populated", "Available metrics", "Observed coverage", "Data-quality notes", "Reviewer status", "Reviewer notes"])
    for item in records:
        repo, git, gh, manual = item["repo"], item["git"], item["gh"], item["manual"]
        verification.append([
            repo["oisi_id"], repo["infrastructure"], repo.get("scored"), repo.get("classification"), repo.get("decision"), repo.get("host"), repo.get("repo"), repo.get("issue_tracker"), git.get("git_status"), gh.get("gh_status"),
            "Yes" if clean(manual.get("m8_issue_response_rate")) or clean(manual.get("m10_issue_close90_rate")) else "No",
            item["available_count"], item["coverage"], quality_notes(repo, git, gh, manual, item["observed"]), "", "",
        ])

    # Keep complete source joins; prefix namespaces to avoid silently overwriting fields.
    data = sheets["Data"]
    merged_rows = []
    for item in records:
        joined = {"oisi_id": item["repo"]["oisi_id"], "infrastructure": item["repo"]["infrastructure"]}
        for prefix, source_row in (("repo", item["repo"]), ("git", item["git"]), ("github", item["gh"]), ("manual", item["manual"])):
            for key, value in source_row.items():
                if key not in {"oisi_id", "infrastructure"}:
                    joined[f"{prefix}_{key}"] = value
        for key in METRIC_KEYS:
            joined[f"equal_{key}"] = item["scores"][key]
            joined[f"observed_{key}"] = int(item["observed"][key])
        joined["equal_zero_filled_final"] = item["final"]
        joined["available_metric_count"] = item["available_count"]
        joined["observed_coverage"] = item["coverage"]
        joined["observed_only_mean"] = item["observed_mean"]
        merged_rows.append(joined)
    data_headers = list(dict.fromkeys(k for row in merged_rows for k in row))
    data.append(data_headers)
    for row in merged_rows:
        data.append([row.get(header, "") for header in data_headers])

    register = sheets["Repository Register"]
    register_headers = list(repos[0].keys())
    register.append(register_headers)
    for repo in repos:
        register.append([repo.get(k, "") for k in register_headers])

    extraction = sheets["Extraction Rules"]
    extraction.append(["Rule", "Current pipeline setting", "Workbook treatment / caveat"])
    extraction_rows = [
        ["Window start", pipeline_rules.WINDOW_START.isoformat(), "UTC inclusive."],
        ["Window end", pipeline_rules.WINDOW_END.isoformat(), "UTC exclusive."],
        ["Commit date", pipeline_rules.DATE_FIELD, "Commit authorship metrics use author date."],
        ["Branch", "Default branch unless repos.csv specifies branch", "Branch selection can change measured activity."],
        ["Authorship", "Merge commits excluded from human authorship counts", "Merge count remains available as a diagnostic in raw input."],
        ["Bot filtering", "Shared rules.py explicit accounts/patterns and GitHub Bot type", "Bot detection is imperfect; use bot_share as a diagnostic where available."],
        ["Identity resolution", "Repository-local .mailmap/email/login/name/manual map rules R1-R6", "Name-based matches are flagged and identity resolution affects contributor metrics."],
        ["Multiple repositories", "Semicolon-separated product repo sets; commit hashes de-duplicated", "Repository selection remains a substantive measurement decision."],
        ["Issue response cohort", f"Full 14-day follow-up; minimum eligible n={pipeline_rules.MIN_ITEMS}", "A low-volume tracker yields unavailable m8 and zero in final score."],
        ["PR response cohort", f"Full 7-day follow-up; minimum eligible n={pipeline_rules.MIN_ITEMS}", "A low-volume tracker yields unavailable m9 and zero in final score."],
        ["Issue closure cohort", f"Full 90-day follow-up; minimum eligible n={pipeline_rules.MIN_ITEMS}", "A low-volume tracker yields unavailable m10 and zero in final score."],
        ["PR decision cohort", f"Full 30-day follow-up; minimum eligible n={pipeline_rules.MIN_ITEMS}", "A low-volume tracker yields unavailable m11 and zero in final score."],
        ["Manual tracker values", "manual_metrics.csv has 16 rows; m8 and m10 are blank in the current file", "Manual data do not currently fill external trackers."],
        ["API scope", "GitHub GraphQL/REST and supported star-history endpoint", "Stars/forks and several response metrics are not platform-neutral."],
    ]
    for row in extraction_rows:
        extraction.append(row)

    dictionary = sheets["Data Dictionary"]
    dictionary.append(["Field / group", "Meaning", "Use in workbook"])
    dict_rows = [
        ["oisi_id", "Stable identifier from the repository register.", "Join key across all inputs."],
        ["repos.csv", "Repository/platform decisions, classification, issue tracker, repo set, branch and upstream mapping.", "Defines which source is analyzed and why entries may be not assessable."],
        ["git_metrics.csv", "Git-history activity, contributor, release and extraction status fields.", "Source for m1-m7 and Git diagnostics."],
        ["gh_metrics.csv", "GitHub issue/PR response rates, sample sizes, stars/forks and API status.", "Source for m8-m13 and response diagnostics."],
        ["manual_metrics.csv", "Optional manual overrides for m8 and m10.", "Current rate fields are blank; included for provenance."],
        ["m1-m13", "Normalized metric scores on 0-100 scale.", "Each contributes 1/13 to the final score; unavailable values are zero-filled."],
        ["Observed flags", "1 where a metric has a usable value under its source/sample rules; otherwise 0.", "Used only for observed coverage and observed-only mean."],
        ["Not assessable", "Register decision indicates no suitable public primary repository, standard, or unresolved source.", "Still receives numeric zero per fixed denominator but must not be read as measured inactivity."],
    ]
    for row in dict_rows:
        dictionary.append(row)

    summary = sheets["Summary"]
    summary.append(["Development Activity: Equal-Metric Results", "Value"])
    summary_rows = [
        ["Main-list entries", n],
        ["Metrics", 13],
        ["Exact weight per metric", WEIGHT],
        ["Displayed weight per metric", "7.69%"],
        ["Final score denominator", 13],
        ["Overall coverage eligibility rule", "None"],
        ["Unavailable metric treatment", "0 in final score"],
        ["Mean zero-filled score", "=AVERAGE(Scoring!U2:U113)"],
        ["Median zero-filled score", "=MEDIAN(Scoring!U2:U113)"],
        ["Minimum score", "=MIN(Scoring!U2:U113)"],
        ["Maximum score", "=MAX(Scoring!U2:U113)"],
        ["Mean observed coverage", "=AVERAGE(Coverage!Q2:Q113)"],
        ["Entries with all 13 metrics observed", sum(r["available_count"] == 13 for r in records)],
        ["Register assessable=yes", sum(r["repo"].get("scored") == "yes" for r in records)],
        ["Register assessable=no", sum(r["repo"].get("scored") != "yes" for r in records)],
        ["Mean final score (Python check)", statistics.mean(finals)],
        ["Repository register input", INPUTS["Repository register"].name],
        ["Git metrics input", INPUTS["Git metrics"].name],
        ["GitHub/API input", INPUTS["GitHub/API metrics"].name],
        ["Manual metrics input", INPUTS["Manual metrics"].name],
        ["Interpretation", "Zero-filled results include missingness/platform availability; review status and metric-level diagnostics before comparing."],
    ]
    for row in summary_rows:
        summary.append(row)

    # Set formatting, print-ready widths, and number formats.
    for name in names:
        ws = sheets[name]
        widths = None
        if name == "Summary": widths = {1: 42, 2: 88}
        elif name == "Scoring": widths = {1: 14, 2: 36, 3: 18, 4: 32, 5: 60, 6: 18, 7: 20, **{c: 17 for c in range(8, 21)}, 21: 19, 22: 18, 23: 24, 24: 18, 25: 20, 26: 72, 27: 100}
        elif name == "Metric Rules": widths = {1: 18, 2: 34, 3: 18, 4: 52, 5: 16, 6: 18, 7: 75, 8: 90, 9: 95, 10: 85}
        elif name == "Methodology": widths = {1: 34, 2: 115}
        elif name == "Parameters": widths = {1: 40, 2: 42, 3: 92}
        elif name == "Diagnostics": widths = {1: 19, 2: 36, 3: 20, 4: 56, 5: 15, 6: 20, 7: 23, 8: 22, 9: 23, 10: 24, 11: 80}
        elif name == "Sensitivity": widths = {1: 46, 2: 18, 3: 15, 4: 25, 5: 27, 6: 29, 7: 26, 8: 27, 9: 18, 10: 75}
        elif name == "Sensitivity Ranks": widths = {1: 14, 2: 38, **{c: 22 for c in range(3, 20)}}
        elif name == "Correlations": widths = {1: 22, 2: 22, 3: 21, 4: 18, 5: 90}
        elif name == "Coverage": widths = {1: 14, 2: 34, **{c: 16 for c in range(3, 16)}, 16: 22, 17: 19, 18: 22, 19: 22, 20: 65}
        elif name == "Verification": widths = {1: 14, 2: 34, 3: 18, 4: 20, 5: 34, 6: 16, 7: 66, 8: 40, 9: 18, 10: 20, 11: 23, 12: 18, 13: 18, 14: 100, 15: 22, 16: 80}
        elif name == "Calc": widths = {1: 14, **{c: 18 for c in range(2, 19)}}
        elif name in {"Data", "Repository Register"}: widths = {c: 32 for c in range(1, ws.max_column+1)}
        elif name == "Extraction Rules": widths = {1: 30, 2: 68, 3: 95}
        elif name == "Data Dictionary": widths = {1: 30, 2: 83, 3: 90}
        style_sheet(ws, widths=widths, freeze="A2")

    # Formats and input notes.
    for r in range(2, n+2):
        for c in range(8, 22):
            scoring.cell(r, c).number_format = "0.00"
        scoring.cell(r, 22).number_format = "0.00"
        scoring.cell(r, 24).number_format = "0.00%"
        scoring.cell(r, 25).number_format = "0.00"
        for c in range(3, 16):
            coverage_ws.cell(r, c).number_format = "0"
        coverage_ws.cell(r, 16).number_format = "0"
        coverage_ws.cell(r, 17).number_format = "0.00%"
        coverage_ws.cell(r, 18).number_format = "0.00"
        coverage_ws.cell(r, 19).number_format = "0.00"
        verification.cell(r, 13).number_format = "0.00%"
    for r in range(2, 15):
        diagnostics.cell(r, 6).number_format = "0.00%"
        for c in (7, 8, 9):
            diagnostics.cell(r, c).number_format = "0.00"
        diagnostics.cell(r, 10).number_format = "0.00%"
    rules_ws.column_dimensions["E"].number_format = "0.0000%"
    for r in range(2, 15):
        rules_ws.cell(r, 5).number_format = "0.0000%"
    for r in range(2, corr.max_row + 1):
        corr.cell(r, 4).number_format = "0.000"
    for r in range(2, sens.max_row+1):
        for c in (4,): sens.cell(r, c).number_format = "0.000"
        for c in (5, 6, 7, 8): sens.cell(r, c).number_format = "0.00"
    summary["B4"].number_format = "0.0000%"
    summary["B9"].number_format = "0.00"
    summary["B10"].number_format = "0.00"
    summary["B11"].number_format = "0.00"
    summary["B12"].number_format = "0.00"
    summary["B13"].number_format = "0.00%"

    # Freeze and formatting on high-level sheets; flag not-assessable rows visually.
    for row_num, item in enumerate(records, start=2):
        if item["repo"].get("scored") != "yes":
            for cell in scoring[row_num]:
                cell.fill = PatternFill("solid", fgColor="FCE9E7")
            for cell in verification[row_num]:
                cell.fill = PatternFill("solid", fgColor="FCE9E7")
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcMode = "auto"
    wb.save(OUTPUT)

    # Cached values make previewers show formulas without a full spreadsheet recalculation.
    cache = {"Scoring": {}, "Coverage": {}, "Summary": {}}
    for i, item in enumerate(records, start=2):
        for j, key in enumerate(METRIC_KEYS, start=8):
            cache["Scoring"][f"{get_column_letter(j)}{i}"] = item["scores"][key]
        cache["Scoring"][f"U{i}"] = item["final"]
        cache["Scoring"][f"V{i}"] = base_ranks[i-2]
        cache["Scoring"][f"W{i}"] = item["available_count"]
        cache["Scoring"][f"X{i}"] = item["coverage"]
        cache["Scoring"][f"Y{i}"] = item["observed_mean"]
        cache["Coverage"][f"P{i}"] = item["available_count"]
        cache["Coverage"][f"Q{i}"] = item["coverage"]
        cache["Coverage"][f"R{i}"] = item["observed_mean"]
        cache["Coverage"][f"S{i}"] = item["final"]
    summary_formulas = {
        "B9": statistics.mean(finals), "B10": statistics.median(finals), "B11": min(finals), "B12": max(finals),
        "B13": statistics.mean(r["coverage"] for r in records),
    }
    cache["Summary"].update(summary_formulas)
    write_formula_caches(OUTPUT, cache)
    print(f"Created {OUTPUT} with {n} entries, 13 equal-weight metrics, no coverage gate.")
    print(f"Mean zero-filled score={statistics.mean(finals):.4f}; mean observed coverage={statistics.mean(r['coverage'] for r in records):.4%}.")
    print(f"Assessable in register={sum(r['repo'].get('scored') == 'yes' for r in records)}; not assessable/excluded={sum(r['repo'].get('scored') != 'yes' for r in records)}.")


if __name__ == "__main__":
    main()
