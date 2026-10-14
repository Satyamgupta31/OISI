"""OISI C4 - normalise, weight and score.

    python score.py git_metrics.csv [gh_metrics.csv] [manual_metrics.csv] scores.csv

manual_metrics.csv (optional): oisi_id plus any raw metric column (e.g. m8_issue_response_rate)
for sources the extractor cannot reach (Jira, GitLab, Gitea). Manual values override.
"""
import csv
import math
import sys

# metric key -> (dimension, weight %, raw column, scorer)
def _log(v, top):
    return 100 * min(1.0, math.log10(1 + v) / math.log10(1 + top))


def s_recency(d):
    if d is None:
        return 0.0
    return 100.0 if d <= 30 else max(0.0, (365 - d) / 335 * 100)


def s_tags(t):
    return 0 if t == 0 else 50 if t == 1 else 75 if t <= 3 else 100


def s_contrib(n):
    return 0.0 if n <= 1 else 100 * min(1.0, math.log2(n) / 5)


def s_eff(e):
    return 0.0 if not e or e <= 1 else 100 * min(1.0, math.log2(e) / 3)


def s_ret(r):
    return 0 if r == 0 else 25 if r == 1 else 50 if r == 2 else 75 if r <= 4 else 100


METRICS = [
    ("m1_commits", "Continuity", 10, "human_commits", lambda v: _log(v, 1000)),
    ("m2_active_weeks", "Continuity", 15, "active_weeks", lambda v: v / 52 * 100),
    ("m3_recency", "Continuity", 5, "days_since_last", s_recency),
    ("m4_releases", "Continuity", 5, "tags_in_window", s_tags),
    ("m5_contributors", "Contributors", 10, "contributors", s_contrib),
    ("m6_effective_contributors", "Contributors", 10, "eff_contributors", s_eff),
    ("m7_retention", "Contributors", 5, "retained", s_ret),
    ("m8_issue_response", "Responsiveness", 10, "m8_issue_response_rate", lambda v: v * 100),
    ("m9_pr_response", "Responsiveness", 10, "m9_pr_response_rate", lambda v: v * 100),
    ("m10_issue_close90", "Responsiveness", 7.5, "m10_issue_close90_rate", lambda v: v * 100),
    ("m11_pr_decision30", "Responsiveness", 7.5, "m11_pr_decision30_rate", lambda v: v * 100),
    ("m12_stars", "Adoption", 2.5, "m12_stars_gained", lambda v: _log(v, 1000)),
    ("m13_forks", "Adoption", 2.5, "m13_forks_created", lambda v: _log(v, 100)),
]
GIT_DORMANT_ZERO = {"days_since_last", "eff_contributors"}  # None means "no activity" -> scores 0, not missing
MIN_TOTAL = 75.0      # % of total weight that must be available
MIN_DIMENSION = 0.5   # share of each dimension's weight that must be available


def num(v):
    if v in (None, "", "None"):
        return None
    try:
        return float(v)
    except ValueError:
        return None


def load(path):
    return {r["oisi_id"]: r for r in csv.DictReader(open(path, encoding="utf-8"))} if path else {}


def score_row(raw, weights=None):
    weights = weights or {k: w for k, _, w, _, _ in METRICS}
    dims, res = {}, {}
    git_ok = raw.get("git_status") == "complete"
    for key, dim, _, col, fn in METRICS:
        w = weights[key]
        v = num(raw.get(col))
        if v is None and git_ok and col in GIT_DORMANT_ZERO:
            s = 0.0
        elif v is None:
            s = None
        else:
            s = round(fn(v), 2)
        res[key] = s
        d = dims.setdefault(dim, [0.0, 0.0])
        d[0] += w
        if s is not None:
            d[1] += w
    avail = sum(weights[k] for k, *_ in METRICS if res[k] is not None)
    total = sum(weights.values())
    contrib = sum(res[k] * weights[k] / 100 for k, *_ in METRICS if res[k] is not None)
    ok = avail >= MIN_TOTAL / 100 * total and all(a >= MIN_DIMENSION * t for t, a in dims.values() if t > 0)
    res["available_weight_pct"] = round(avail / total * 100, 1)
    res["weighted_points"] = round(contrib, 2)
    res["score_on_available"] = round(contrib / avail * 100, 1) if avail else None
    res["component4_score"] = res["score_on_available"] if ok else None
    res["score_status"] = "scored" if ok else "insufficient data"
    return res


def main(git_csv, gh_csv=None, manual_csv=None, out_csv="scores.csv"):
    git, gh, man = load(git_csv), load(gh_csv), load(manual_csv)
    out = []
    for oid, g in git.items():
        raw = {**g, **gh.get(oid, {}), **{k: v for k, v in man.get(oid, {}).items() if v not in ("", None)}}
        out.append({"oisi_id": oid, "infrastructure": g["infrastructure"], **score_row(raw)})
    with open(out_csv, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(out[0]))
        w.writeheader()
        w.writerows(out)


if __name__ == "__main__":
    args = sys.argv[1:]
    main(args[0], *(a or None for a in args[1:-1]), out_csv=args[-1])
