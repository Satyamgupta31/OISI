"""OISI C4 - weight sensitivity and metric correlation.

    python analysis.py git_metrics.csv [gh_metrics.csv] out_prefix
Uses every metric that is available for >= 80% of repos. Writes:
  <prefix>_correlation.csv   Spearman rho between metric scores
  <prefix>_sensitivity.csv   rank of each repo under each weight scheme
  <prefix>_summary.txt       headline numbers
"""
import csv
import itertools
import statistics
import sys

import score as S


def ranks(vals):
    order = sorted(range(len(vals)), key=lambda i: vals[i])
    r = [0.0] * len(vals)
    i = 0
    while i < len(order):
        j = i
        while j + 1 < len(order) and vals[order[j + 1]] == vals[order[i]]:
            j += 1
        for k in range(i, j + 1):
            r[order[k]] = (i + j) / 2 + 1
        i = j + 1
    return r


def spearman(a, b):
    ra, rb = ranks(a), ranks(b)
    ma, mb = statistics.mean(ra), statistics.mean(rb)
    num = sum((x - ma) * (y - mb) for x, y in zip(ra, rb))
    den = (sum((x - ma) ** 2 for x in ra) * sum((y - mb) ** 2 for y in rb)) ** 0.5
    return num / den if den else 0.0


def main(git_csv, gh_csv=None, prefix="analysis"):
    git, gh = S.load(git_csv), S.load(gh_csv)
    rows = []
    for oid, g in git.items():
        if g.get("git_status") != "complete":
            continue
        raw = {**g, **gh.get(oid, {})}
        rows.append((oid, g["infrastructure"], S.score_row(raw)))
    keys = [k for k, *_ in S.METRICS if sum(r[2][k] is not None for r in rows) >= 0.8 * len(rows)]
    rows = [r for r in rows if all(r[2][k] is not None for k in keys)]
    n = len(rows)
    meta = {k: (d, w) for k, d, w, *_ in S.METRICS}

    # correlation
    corr = []
    for a, b in itertools.combinations(keys, 2):
        corr.append((a, b, round(spearman([r[2][a] for r in rows], [r[2][b] for r in rows]), 3)))
    with open(f"{prefix}_correlation.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["metric_a", "metric_b", "spearman_rho"])
        w.writerows(sorted(corr, key=lambda x: -abs(x[2])))

    # weight schemes (restricted to available metrics, renormalised)
    base = {k: meta[k][1] for k in keys}
    dims = sorted({meta[k][0] for k in keys})
    schemes = {"proposed": base, "equal_metrics": {k: 1 for k in keys}}
    schemes["equal_dimensions"] = {k: 1 / sum(meta[j][0] == meta[k][0] for j in keys) for k in keys}
    for d in dims:
        schemes[f"drop_{d}"] = {k: (0 if meta[k][0] == d else base[k]) for k in keys}
    for k in keys:
        schemes[f"drop_{k}"] = {j: (0 if j == k else base[j]) for j in keys}

    def comp(r, ws):
        t = sum(ws.values())
        return sum(r[2][k] * ws[k] for k in keys) / t

    results = {name: [comp(r, ws) for r in rows] for name, ws in schemes.items()}
    rk = {name: ranks([-v for v in vals]) for name, vals in results.items()}
    with open(f"{prefix}_sensitivity.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["oisi_id", "infrastructure", "proposed_score"] + [f"rank_{s}" for s in schemes])
        order = sorted(range(n), key=lambda i: rk["proposed"][i])
        for i in order:
            w.writerow([rows[i][0], rows[i][1], round(results["proposed"][i], 1)] + [rk[s][i] for s in schemes])

    lines = [f"Repos analysed: {n}; metrics used: {', '.join(keys)}", "",
             "Rank stability vs proposed weights (Spearman rho; median / max absolute rank shift):"]
    for s in schemes:
        if s == "proposed":
            continue
        rho = spearman(results["proposed"], results[s])
        shifts = [abs(a - b) for a, b in zip(rk["proposed"], rk[s])]
        lines.append(f"  {s:32} rho={rho:.3f}  median shift={statistics.median(shifts):.1f}  max shift={max(shifts):.1f}")
    top10 = set(sorted(range(n), key=lambda i: rk["proposed"][i])[:10])
    for s in ("equal_metrics", "equal_dimensions"):
        t = set(sorted(range(n), key=lambda i: rk[s][i])[:10])
        lines.append(f"  top-10 overlap proposed vs {s}: {len(top10 & t)}/10")
    lines += ["", "Highly correlated metric pairs (|rho| >= 0.7):"]
    lines += [f"  {a} ~ {b}: {r}" for a, b, r in sorted(corr, key=lambda x: -abs(x[2])) if abs(r) >= 0.7] or ["  none"]
    with open(f"{prefix}_summary.txt", "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    print("\n".join(lines))


if __name__ == "__main__":
    a = sys.argv[1:]
    main(a[0], a[1] if len(a) > 2 else None, a[-1])
