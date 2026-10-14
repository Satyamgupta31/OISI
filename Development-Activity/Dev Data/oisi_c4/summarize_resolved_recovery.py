"""Summarize recovery after applying resolved URL markers."""
import csv


def load(path):
    with open(path, encoding="utf-8-sig", newline="") as f:
        return {r["oisi_id"]: r for r in csv.DictReader(f)}


def present(v):
    return str(v or "").strip() not in {"", "NA", "None", "nan"}


before_scores = load("scores_2026_10_14_resolved_urls_before.csv")
scores = load("scores.csv")
before_git = load("git_metrics_2026_10_14_resolved_urls_before.csv")
git = load("git_metrics.csv")
repos = load("repos.csv")
markers = load("resolved_url_markers.csv")

rows = []
for oid, marker in markers.items():
    r = repos.get(oid, {})
    g0 = before_git.get(oid, {})
    g1 = git.get(oid, {})
    s0 = before_scores.get(oid, {})
    s1 = scores.get(oid, {})
    rows.append({
        "oisi_id": oid,
        "infrastructure": marker["infrastructure"],
        "resolved_action": marker["action"],
        "candidate_repo": marker["candidate_repo"],
        "current_repo": r.get("repo", ""),
        "scored_flag": r.get("scored", ""),
        "git_status_before": g0.get("git_status", "not in git_metrics"),
        "git_status_after": g1.get("git_status", "not in git_metrics"),
        "human_commits_after": g1.get("human_commits", ""),
        "active_weeks_after": g1.get("active_weeks", ""),
        "contributors_after": g1.get("contributors", ""),
        "score_status_before": s0.get("score_status", "not in scores"),
        "score_status_after": s1.get("score_status", "not in scores"),
        "available_weight_before": s0.get("available_weight_pct", ""),
        "available_weight_after": s1.get("available_weight_pct", ""),
        "component4_score_after": s1.get("component4_score", ""),
        "remaining_gap": "needs GitHub/API or manual issue-PR/adoption metrics" if s1.get("score_status") == "insufficient data" else "",
    })

with open("resolved_url_recovery_report.csv", "w", encoding="utf-8", newline="") as f:
    writer = csv.DictWriter(f, fieldnames=list(rows[0]))
    writer.writeheader()
    writer.writerows(rows)

print("wrote resolved_url_recovery_report.csv")
print("markers", len(rows))
print("applied", sum(1 for r in rows if r["resolved_action"] == "applied"))
print("git complete after", sum(1 for r in rows if r["git_status_after"] == "complete"))
print("newly present in git", sum(1 for r in rows if r["git_status_before"] == "not in git_metrics" and r["git_status_after"] == "complete"))
