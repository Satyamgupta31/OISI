"""OISI C4 - GitHub extraction for metrics 8-13 (+ responsiveness diagnostics).

Run on your own machine (needs internet + a GitHub token with public read access):
    export GITHUB_TOKEN=ghp_...            # fine-grained token, "Public repositories (read-only)"
    python gh_extract.py repos.csv gh_metrics.csv
Raw API records are cached in raw/<owner>__<repo>/*.json so a re-run is cheap and
every number can be audited. Uses the GitHub GraphQL API for issues/PRs/forks and
GitHub's REST star-history endpoint for stars.
"""
import csv
import json
import os
import statistics
import sys
import time
from datetime import datetime, timedelta, timezone

import urllib.request

import rules as R

API = "https://api.github.com/graphql"
REST = "https://api.github.com"
RAW = "raw"

ACTOR = "login __typename"
ISSUES_Q = """
query($o:String!,$n:String!,$c:String){ repository(owner:$o,name:$n){
  issues(first:50, after:$c, orderBy:{field:CREATED_AT,direction:DESC}){
    pageInfo{hasNextPage endCursor}
    nodes{ number createdAt closedAt authorAssociation author{%(a)s}
      comments(first:20){nodes{createdAt author{%(a)s}}}
      timelineItems(first:30, itemTypes:[CLOSED_EVENT,LABELED_EVENT,ASSIGNED_EVENT]){nodes{
        __typename
        ... on ClosedEvent{createdAt actor{%(a)s}}
        ... on LabeledEvent{createdAt actor{%(a)s}}
        ... on AssignedEvent{createdAt actor{%(a)s}} }}}}}}
""" % {"a": ACTOR}
PRS_Q = """
query($o:String!,$n:String!,$c:String){ repository(owner:$o,name:$n){
  pullRequests(first:50, after:$c, orderBy:{field:CREATED_AT,direction:DESC}){
    pageInfo{hasNextPage endCursor}
    nodes{ number createdAt closedAt mergedAt state authorAssociation author{%(a)s} mergedBy{%(a)s}
      comments(first:20){nodes{createdAt author{%(a)s}}}
      reviews(first:20){nodes{submittedAt author{%(a)s}}}
      timelineItems(first:10, itemTypes:[CLOSED_EVENT]){nodes{... on ClosedEvent{createdAt actor{%(a)s}}}}}}}}
""" % {"a": ACTOR}
FORKS_Q = """
query($o:String!,$n:String!,$c:String){ repository(owner:$o,name:$n){
  forks(first:100, after:$c, orderBy:{field:CREATED_AT,direction:DESC}){
    pageInfo{hasNextPage endCursor} nodes{createdAt}}}}
"""
COUNT_Q = """query($q:String!){ search(query:$q, type:ISSUE, first:1){ issueCount } }"""


def gql(query, variables, token, tries=5):
    body = json.dumps({"query": query, "variables": variables}).encode()
    for i in range(tries):
        req = urllib.request.Request(API, body, {"Authorization": f"Bearer {token}",
                                                 "Content-Type": "application/json"})
        try:
            with urllib.request.urlopen(req, timeout=120) as r:
                d = json.loads(r.read())
            if "errors" in d and not d.get("data"):
                raise RuntimeError(d["errors"])
            return d["data"]
        except Exception as ex:  # secondary rate limits, timeouts
            if i == tries - 1:
                raise
            time.sleep(20 * (i + 1))


def rest_json(path, token, tries=5):
    headers = {
        "Accept": "application/vnd.github+json",
        "Authorization": f"Bearer {token}",
        "X-GitHub-Api-Version": "2026-03-10",
    }
    for i in range(tries):
        try:
            req = urllib.request.Request(f"{REST}{path}", headers=headers)
            with urllib.request.urlopen(req, timeout=120) as r:
                return json.loads(r.read())
        except Exception:
            if i == tries - 1:
                raise
            time.sleep(20 * (i + 1))


def ts(s):
    return datetime.fromisoformat(s.replace("Z", "+00:00")).timestamp() if s else None


def paged(query, field, o, n, token, stop_before, created_key="createdAt"):
    """Newest-first pages; stop once items are older than stop_before."""
    out, cur = [], None
    while True:
        d = gql(query, {"o": o, "n": n, "c": cur}, token)["repository"][field]
        items = d.get("nodes") or d.get("edges") or []
        out.extend(items)
        oldest = ts(items[-1][created_key]) if items else None
        if not d["pageInfo"]["hasNextPage"] or (oldest is not None and oldest < stop_before):
            return out
        cur = d["pageInfo"]["endCursor"]


def bot(actor):
    if not actor:
        return True  # deleted ("ghost") accounts cannot be attributed; treat as non-human
    return R.is_bot(login=actor.get("login", ""), user_type=actor.get("__typename", ""))


def login(actor):
    return (actor or {}).get("login", "").lower()


# ---------------------------------------------------------------- metric logic
def first_response(item, kind):
    """Earliest qualifying response time (epoch) by someone other than the author, or None."""
    author = login(item.get("author"))
    ev = []
    for c in (item.get("comments") or {}).get("nodes", []):
        if not bot(c["author"]) and login(c["author"]) != author:
            ev.append(ts(c["createdAt"]))
    if kind == "pr":
        for rv in (item.get("reviews") or {}).get("nodes", []):
            if rv.get("submittedAt") and not bot(rv["author"]) and login(rv["author"]) != author:
                ev.append(ts(rv["submittedAt"]))
        mb = item.get("mergedBy")
        if item.get("mergedAt") and mb and not bot(mb) and login(mb) != author:
            ev.append(ts(item["mergedAt"]))
    for t in (item.get("timelineItems") or {}).get("nodes", []):
        a = t.get("actor")
        if t.get("createdAt") and a and not bot(a) and login(a) != author:
            ev.append(ts(t["createdAt"]))
    return min(ev) if ev else None


def rate(items, ok):
    n = len(items)
    k = sum(1 for i in items if ok(i))
    return (round(k / n, 4) if n >= R.MIN_ITEMS else None), n, k


MAINT = {"OWNER", "MEMBER", "COLLABORATOR"}


def issue_metrics(issues):
    human = [i for i in issues if not bot(i.get("author"))]
    cut8 = R.E - R.ISSUE_RESPONSE_DAYS * 86400
    c8 = [i for i in human if R.S <= ts(i["createdAt"]) < cut8]
    ext = [i for i in c8 if i.get("authorAssociation") not in MAINT]
    cohort, basis = (ext, "non-maintainer issues") if len(ext) >= R.MIN_ITEMS else (c8, "all issues (too few non-maintainer)")
    lim = R.ISSUE_RESPONSE_DAYS * 86400
    r8, n8, k8 = rate(cohort, lambda i: (fr := first_response(i, "issue")) is not None and fr - ts(i["createdAt"]) <= lim)
    ttfr = [(fr - ts(i["createdAt"])) / 3600 for i in cohort if (fr := first_response(i, "issue")) is not None]
    cut10 = R.E - R.ISSUE_CLOSE_DAYS * 86400
    c10 = [i for i in human if R.S <= ts(i["createdAt"]) < cut10]
    r10, n10, k10 = rate(c10, lambda i: i["closedAt"] and ts(i["closedAt"]) - ts(i["createdAt"]) <= R.ISSUE_CLOSE_DAYS * 86400)
    created = [i for i in human if R.S <= ts(i["createdAt"]) < R.E]
    return {
        "m8_issue_response_rate": r8, "m8_n": n8, "m8_responded": k8, "m8_basis": basis,
        "diag_issue_ttfr_median_h": round(statistics.median(ttfr), 2) if ttfr else None, "diag_issue_ttfr_n": len(ttfr),
        "m10_issue_close90_rate": r10, "m10_n": n10, "m10_closed": k10,
        "diag_issues_created": len(created),
        "diag_issues_by_maintainers": sum(1 for i in created if i.get("authorAssociation") in MAINT),
    }


def pr_metrics(prs):
    human = [p for p in prs if not bot(p.get("author"))]
    cut9 = R.E - R.PR_RESPONSE_DAYS * 86400
    c9 = [p for p in human if R.S <= ts(p["createdAt"]) < cut9]
    lim = R.PR_RESPONSE_DAYS * 86400
    r9, n9, k9 = rate(c9, lambda p: (fr := first_response(p, "pr")) is not None and fr - ts(p["createdAt"]) <= lim)
    ttfr = [(fr - ts(p["createdAt"])) / 3600 for p in c9 if (fr := first_response(p, "pr")) is not None]
    cut11 = R.E - R.PR_DECISION_DAYS * 86400
    c11 = [p for p in human if R.S <= ts(p["createdAt"]) < cut11]

    def decided(p):
        t = p.get("mergedAt") or p.get("closedAt")
        return t and ts(t) - ts(p["createdAt"]) <= R.PR_DECISION_DAYS * 86400
    r11, n11, k11 = rate(c11, decided)
    win = [p for p in human if R.S <= ts(p["createdAt"]) < R.E]
    merged = sum(1 for p in win if p.get("mergedAt") and ts(p["mergedAt"]) < R.E)
    closed_unmerged = sum(1 for p in win if not p.get("mergedAt") and p.get("closedAt") and ts(p["closedAt"]) < R.E)
    self_merged = sum(1 for p in win if p.get("mergedAt") and login(p.get("mergedBy")) == login(p.get("author")))
    return {
        "m9_pr_response_rate": r9, "m9_n": n9, "m9_responded": k9,
        "diag_pr_ttfr_median_h": round(statistics.median(ttfr), 2) if ttfr else None, "diag_pr_ttfr_n": len(ttfr),
        "m11_pr_decision30_rate": r11, "m11_n": n11, "m11_decided": k11,
        "diag_prs_created": len(win), "diag_prs_merged": merged, "diag_prs_closed_unmerged": closed_unmerged,
        "diag_prs_open_at_cutoff": len(win) - merged - closed_unmerged, "diag_prs_self_merged": self_merged,
        "diag_acceptance_ratio": round(merged / (merged + closed_unmerged) * 100, 2) if merged + closed_unmerged else None,
    }


def count_in_window(items, key):
    return sum(1 for i in items if i.get(key) and R.S <= ts(i[key]) < R.E)


def stars_from_history(history):
    """Count daily star additions from GitHub's privacy-safe star history API."""
    total = 0
    for week in history:
        start = datetime.fromtimestamp(week["week"], tz=timezone.utc)
        for offset, count in enumerate(week.get("days", [])):
            day = start + timedelta(days=offset)
            if R.WINDOW_START <= day < R.WINDOW_END:
                total += int(count or 0)
    return total


def fetch_star_history(owner, repo, token):
    out, page = [], 1
    while True:
        batch = rest_json(f"/repos/{owner}/{repo}/stargazers/history?per_page=30&page={page}", token)
        if not batch:
            return out
        out.extend(batch)
        if len(batch) < 30:
            return out
        page += 1


# ---------------------------------------------------------------- driver
def cache(path, fn):
    if os.path.exists(path):
        return json.load(open(path, encoding="utf-8"))
    data = fn()
    os.makedirs(os.path.dirname(path), exist_ok=True)
    json.dump(data, open(path, "w", encoding="utf-8"))
    return data


def gh_parts(s):
    return [p.strip() for p in (s or "").split(";") if p.strip() and not p.startswith("http")]


def main(repos_csv, out_csv):
    token = os.environ.get("GITHUB_TOKEN") or sys.exit("Set GITHUB_TOKEN first")
    rows = [r for r in csv.DictReader(open(repos_csv, encoding="utf-8")) if r["scored"] == "yes"]
    out = []
    for r in rows:
        rec = {"oisi_id": r["oisi_id"], "infrastructure": r["infrastructure"]}
        try:
            code = gh_parts(r["repo"]) if r["host"] == "github" else []
            prs, stars, forks = [], 0, 0
            stars_ok = True
            for full in code:
                o, n = full.split("/")
                base = f"{RAW}/{o}__{n}"
                prs += cache(f"{base}/prs.json", lambda: paged(PRS_Q, "pullRequests", o, n, token, R.S))
                try:
                    st = cache(f"{base}/stars_history.json", lambda: fetch_star_history(o, n, token))
                    stars += stars_from_history(st)
                except Exception:
                    stars_ok = False
                fk = cache(f"{base}/forks.json", lambda: paged(FORKS_Q, "forks", o, n, token, R.S))
                forks += count_in_window(fk, "createdAt")
            if code:
                rec.update(pr_metrics(prs), m13_forks_created=forks)
                rec["m12_stars_gained"] = stars if stars_ok else None
            ir = r.get("issue_repo", "")
            if ir:
                o, n = ir.split("/")
                iss = cache(f"{RAW}/{o}__{n}/issues.json", lambda: paged(ISSUES_Q, "issues", o, n, token, R.S))
                rec.update(issue_metrics(iss))
                old = (R.WINDOW_END - timedelta(days=180)).date().isoformat()
                q1 = gql(COUNT_Q, {"q": f"repo:{ir} is:issue is:open created:<{old}"}, token)["search"]["issueCount"]
                q2 = gql(COUNT_Q, {"q": f"repo:{ir} is:issue created:<{old} closed:>={R.WINDOW_END.date()}"}, token)["search"]["issueCount"]
                rec["diag_open_issues_over_180d_at_cutoff"] = q1 + q2
            rec["gh_status"] = "complete" if code or ir else "no GitHub source"
        except Exception as ex:
            rec["gh_status"] = f"failed: {str(ex)[:150]}"
        print(r["infrastructure"], rec["gh_status"], file=sys.stderr, flush=True)
        out.append(rec)
    cols = sorted({k for x in out for k in x}, key=lambda k: (k not in ("oisi_id", "infrastructure"), k))
    with open(out_csv, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        w.writerows(out)


if __name__ == "__main__":
    main(*sys.argv[1:3])
