"""OISI C4 - commit-based metrics (1-7 + diagnostics) from a local git clone.

Usage:  python git_metrics.py repos.csv clones_dir out.csv
repos.csv needs columns: oisi_id, infrastructure, host, repo (owner/name or full git URL; several separated by ';' for a
multi-repo product), branch (optional), upstream (optional: owner/name of the project this repo is a fork of), scored (yes/no)
Clones are made with `git clone --bare --filter=tree:0` (history only, small).
Works for any git host (GitHub, GitLab, Gitea, Bitbucket).
"""
import collections
import csv
import os
import subprocess
import sys
from datetime import datetime, timezone

import rules as R

FMT = "%H%x1f%P%x1f%at%x1f%ct%x1f%aN%x1f%aE"
GIT_ENV = {**os.environ, "GIT_TERMINAL_PROMPT": "0"}


def git_run(args, **kwargs):
    return subprocess.run(args, env=GIT_ENV, **kwargs)


def git_output(args):
    return subprocess.check_output(args, env=GIT_ENV, text=True, encoding="utf-8", errors="replace")


def git_url(row):
    r = row["repo"].strip()
    if r.startswith("http"):
        return r if r.endswith(".git") else r + ".git"
    host = {"github": "github.com", "gitlab": "gitlab.com"}.get(row.get("host", "github"), row.get("host"))
    return f"https://{host}/{r}.git"


def ensure_clone(row, clones_dir):
    d = os.path.join(clones_dir, row["repo"].rstrip("/").split("://")[-1].replace("/", "_") + ".git")
    if not os.path.isdir(d):
        git_run(["git", "clone", "-q", "--bare", "--filter=tree:0", git_url(row), d],
                check=True, timeout=600)
    # Ensure Windows bare clone refs directory structure exists
    os.makedirs(os.path.join(d, "refs", "heads"), exist_ok=True)
    os.makedirs(os.path.join(d, "refs", "tags"), exist_ok=True)
    for f in ("HEAD", "packed-refs"):
        p = os.path.join(d, f)
        if os.path.isfile(p):
            with open(p, "rb") as fp:
                c = fp.read()
            if b"\r\n" in c:
                with open(p, "wb") as fp:
                    fp.write(c.replace(b"\r\n", b"\n"))
    return d


def load_manual(path="identity_map.csv"):
    m = collections.defaultdict(dict)
    if os.path.exists(path):
        for r in csv.DictReader(open(path)):
            m[r["repo"].lower()][r["email"].lower()] = r["canonical_id"]
    return m


def default_ref(clone, branch=""):
    if branch:
        return "refs/heads/" + branch
    return git_output(["git", "-C", clone, "symbolic-ref", "HEAD"]).strip()


def log_rows(clone, branch="", exclude_upstream=False):
    args = ["git", "-C", clone, "log", default_ref(clone, branch)]
    if exclude_upstream:
        args.append("^refs/upstream/head")  # commits inherited from the upstream project are not this fork's work
    out = git_output(args + [f"--format={FMT}"])
    return [l.split("\x1f") for l in out.splitlines() if l]


def commit_metrics(clones, branch="", manual=None, exclude_upstream=False):
    """clones: one path or a list (multi-repo product; commits de-duplicated by hash)."""
    if isinstance(clones, str):
        clones = [clones]
    seen, rows = set(), []
    for c in clones:
        for r in log_rows(c, branch if len(clones) == 1 else "", exclude_upstream):
            if r[0] not in seen:
                seen.add(r[0])
                rows.append(r)
    ir = R.IdentityResolver(manual)
    win, merges, bots = [], 0, 0
    last_any = 0
    for h, parents, at, ct, name, email in rows:
        t = int(at)
        last_any = max(last_any, t)
        if not (R.S <= t < R.E):
            continue
        if len(parents.split()) > 1:
            merges += 1
            continue
        if R.is_bot(name, email):
            bots += 1
            continue
        win.append((t, ir.canonical(name, email), email.lower()))

    n = len(win)
    res = dict(human_commits=n, merge_commits=merges, bot_commits=bots,
               bot_share=round(bots / (n + bots), 3) if n + bots else None,
               last_commit_any=datetime.fromtimestamp(last_any, timezone.utc).date().isoformat() if last_any else "")
    if not n:
        res.update(active_weeks=0, active_months=0, days_since_last=None, contributors=0,
                   eff_contributors=None, top_share=None, bf50=None, bf80=None, retained=0,
                   merged_identity_groups=0)
        return res
    c = collections.Counter(cid for _, cid, _ in win)
    emails_per = collections.defaultdict(set)
    for _, cid, e in win:
        emails_per[cid].add(e)
    shares = [v / n for v in c.values()]
    cum, bf = 0, {}
    for i, (_, v) in enumerate(c.most_common(), 1):
        cum += v
        for th in (0.5, 0.8):
            if th not in bf and cum >= th * n:
                bf[th] = i
    h1 = {cid for t, cid, _ in win if t < R.MID}
    h2 = {cid for t, cid, _ in win if t >= R.MID}
    weeks = {min(51, int((t - R.S) // (7 * 86400))) for t, _, _ in win}
    months = {datetime.fromtimestamp(t, timezone.utc).strftime("%Y-%m") for t, _, _ in win}
    res.update(
        active_weeks=len(weeks), active_months=len(months),
        days_since_last=round((R.E - max(t for t, _, _ in win)) / 86400, 2),
        contributors=len(c), eff_contributors=round(1 / sum(s * s for s in shares), 2),
        top_share=round(max(shares) * 100, 2), bf50=bf[0.5], bf80=bf[0.8],
        retained=len(h1 & h2),
        merged_identity_groups=sum(1 for v in emails_per.values() if len(v) > 1),
    )
    return res


def tags_in_window(clones):
    if isinstance(clones, str):
        clones = [clones]
    n = 0
    for c in clones:
        out = git_output(["git", "-C", c, "for-each-ref", "refs/tags", "--format=%(creatordate:unix)"])
        n += sum(1 for t in out.split() if t.isdigit() and R.S <= int(t) < R.E)
    return n


def add_upstream(clone, upstream):
    """Fetch the upstream project's default branch into refs/upstream/head."""
    url = upstream if upstream.startswith("http") else f"https://github.com/{upstream}.git"
    git_run(["git", "-C", clone, "fetch", "-q", "--filter=tree:0", url, "+HEAD:refs/upstream/head"],
            check=True, timeout=600)


def main(repos_csv, clones_dir, out_csv):
    manual = load_manual()
    rows = [r for r in csv.DictReader(open(repos_csv)) if r.get("repo") and r.get("scored", "yes") == "yes"]
    os.makedirs(clones_dir, exist_ok=True)
    results = []
    for r in rows:
        rec = {"oisi_id": r["oisi_id"], "infrastructure": r["infrastructure"], "repo": r["repo"]}
        try:
            parts = [p.strip() for p in r["repo"].split(";") if p.strip()]
            ds = [ensure_clone({**r, "repo": p}, clones_dir) for p in parts]
            up = r.get("upstream", "").strip()
            if up:
                add_upstream(ds[0], up)
            rec.update(commit_metrics(ds, r.get("branch", ""), manual.get(r["repo"].lower()), bool(up)))
            rec["tags_in_window"] = tags_in_window(ds)
            rec["git_status"] = "complete"
        except Exception as ex:  # keep going; record why
            rec["git_status"] = f"failed: {str(ex)[:120]}"
        results.append(rec)
        print(rec["repo"], rec.get("git_status"), file=sys.stderr)
    cols = sorted({k for r in results for k in r}, key=lambda k: (k not in ("oisi_id", "infrastructure", "repo"), k))
    with open(out_csv, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        w.writerows(results)


if __name__ == "__main__":
    main(*sys.argv[1:4])
