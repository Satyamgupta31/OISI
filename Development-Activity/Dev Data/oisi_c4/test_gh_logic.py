"""Offline tests for gh_extract metric logic (no network)."""
import gh_extract as G
U = lambda l, t="User": {"login": l, "__typename": t}

def iso(day, h=0):  # day offset from 2025-10-01
    from datetime import datetime, timedelta, timezone
    return (datetime(2025, 10, 1, tzinfo=timezone.utc) + timedelta(days=day, hours=h)).isoformat().replace("+00:00", "Z")

def issue(d, author="ext", assoc="NONE", comments=(), closed=None, closer=None, label=None):
    tl = []
    if closed is not None: tl.append({"__typename": "ClosedEvent", "createdAt": iso(closed), "actor": U(closer or "maint")})
    if label is not None: tl.append({"__typename": "LabeledEvent", "createdAt": iso(label), "actor": U("maint")})
    return {"createdAt": iso(d), "closedAt": iso(closed) if closed is not None else None, "authorAssociation": assoc,
            "author": U(author), "comments": {"nodes": [{"createdAt": iso(c), "author": U(a, t)} for c, a, t in comments]},
            "timelineItems": {"nodes": tl}}

iss = []
iss += [issue(i, comments=[(i + 1, "maint", "User")]) for i in range(8)]                 # answered in 1 day
iss += [issue(10 + i, comments=[(10 + i + 1, "dependabot", "Bot")]) for i in range(3)]    # only a bot replied -> unanswered
iss += [issue(20, comments=[(21, "ext", "User")])]                                          # only the author replied -> unanswered
iss += [issue(30, closed=50, closer="maint")]                                              # closed by maintainer within 20 d -> response
iss += [issue(40, label=60)]                                                                # labelled after 20 d -> too late
iss += [issue(5, author="maint", assoc="MEMBER") for _ in range(3)]                        # maintainer tickets
m = G.issue_metrics(iss)
assert m["m8_n"] == 14 and m["m8_responded"] == 8, m          # 8 of 14 non-maintainer issues answered within 14 d
assert m["m8_basis"] == "non-maintainer issues"
assert m["diag_issues_by_maintainers"] == 3
assert m["m10_n"] == 17 and m["m10_closed"] == 1, m            # one closure within 90 d

pr = lambda d, author="ext", merged=None, by="maint", reviews=(): {
    "createdAt": iso(d), "closedAt": iso(merged) if merged is not None else None,
    "mergedAt": iso(merged) if merged is not None else None, "mergedBy": U(by) if merged is not None else None,
    "authorAssociation": "NONE", "author": U(author), "comments": {"nodes": []},
    "reviews": {"nodes": [{"submittedAt": iso(r), "author": U("maint")} for r in reviews]}, "timelineItems": {"nodes": []}}
prs = [pr(i, merged=i + 2) for i in range(6)]                   # merged by maintainer in 2 d -> responded + decided
prs += [pr(10 + i, author="maint", merged=10 + i, by="maint") for i in range(4)]  # self-merged, no review -> not responded
prs += [pr(20, reviews=[25])]                                   # reviewed in 5 d, still open
prs += [pr(30, author="renovate[bot]", merged=31)]              # bot PR excluded
m = G.pr_metrics(prs)
assert m["m9_n"] == 11 and m["m9_responded"] == 7, m
assert m["m11_n"] == 11 and m["m11_decided"] == 10, m
assert m["diag_prs_self_merged"] == 4 and m["diag_prs_open_at_cutoff"] == 1 and m["diag_acceptance_ratio"] == 100.0

history = [
    {"week": G.ts("2025-09-28T00:00:00Z"), "total": 9, "days": [1, 1, 1, 1, 1, 1, 1]},  # first 3 days outside window
    {"week": G.ts("2025-10-05T00:00:00Z"), "total": 16, "days": [2, 2, 2, 2, 2, 2, 2]},
    {"week": G.ts("2026-09-27T00:00:00Z"), "total": 23, "days": [3, 3, 3, 3, 3, 3, 3]},  # last 3 days outside window
]
assert G.stars_from_history(history) == 4 + 14 + 12
print("all extractor logic tests passed")
