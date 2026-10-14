"""Apply user-supplied resolved URL markers to repos.csv.

This keeps an audit trail in repos.csv and writes resolved_url_markers.csv.
Rows that were previously excluded or had a non-GitHub cloneable repo are
restored for scoring. Rows where the existing Component 4 decision intentionally
uses a product set are marked but not overwritten.
"""
import csv


RESOLVED = {
    "C4-005": ("esmero/strawberryfield", "GitHub", "marker_only", "Current scoring uses the Archipelago core module set; single repo recorded as candidate."),
    "C4-006": ("arks-org/arks.github.io", "GitHub", "applied", "Restored as scoreable documentation/spec repository per resolved list; still marked as standard."),
    "C4-010": ("arXiv/arxiv-browse", "GitHub", "already_current", "Current scoring target already matches resolved list."),
    "C4-016": ("projectblacklight/blacklight", "GitHub", "already_current", "Current scoring target already matches resolved list."),
    "C4-017": ("brcris-ibict/brcris-interface-busca", "GitHub", "applied", "Restored from not-assessable using resolved GitHub repository."),
    "C4-019": ("lockss/lockss-daemon", "GitHub", "already_current", "Current scoring target already matches resolved list."),
    "C4-022": ("collectiveaccess/providence", "GitHub", "already_current", "Current scoring target already matches resolved list."),
    "C4-023": ("oacore/dashboard", "GitHub", "applied", "Restored from not-assessable using resolved GitHub repository."),
    "C4-024": ("Project-Counter/cop5", "GitHub", "applied", "Restored as scoreable documentation/spec repository per resolved list; still marked as standard."),
    "C4-025": ("creativecommons/cc-legal-tools", "GitHub", "marker_only", "Current scoring uses cc-legal-tools-app; resolved repo recorded as candidate."),
    "C4-026": ("https://gitlab.com/crossref/rest_api", "GitLab", "already_current", "Current scoring target already matches resolved list."),
    "C4-027": ("code-repo.d4science.org", "Gitea Instance", "marker_only", "Resolved target is the host root, not a cloneable repository."),
    "C4-031": ("datacite/lupo", "GitHub", "already_current", "Current scoring target already matches resolved list."),
    "C4-032": ("ds-wizard/wizard-server", "GitHub", "marker_only", "Current scoring uses backend + frontend product set; single server repo recorded as candidate."),
    "C4-034": ("https://gitlab.unige.ch/dlcm/community/dlcm-backend", "GitLab (UNIGE)", "already_current", "Current scoring target already matches resolved list."),
    "C4-035": ("https://gitlab.unige.ch/dlcm/ui/dlcm-portal", "GitLab (UNIGE)", "already_current", "Current scoring target already matches resolved list."),
    "C4-037": ("DOAJ/doaj", "GitHub", "already_current", "Current scoring target already matches resolved list."),
    "C4-046": ("fair-wizard/fair-wizard-integration-sdk", "GitHub", "applied", "Restored from duplicate exclusion using resolved GitHub repository."),
    "C4-065": ("mets/METS-schema", "GitHub", "applied", "Restored as scoreable schema repository per resolved list; still marked as standard."),
    "C4-071": ("https://bitbucket.org/oaswitchboard/api", "Bitbucket", "already_current", "Current scoring target already matches resolved list."),
    "C4-073": ("https://code-repo.d4science.org/D-Net/dnet-hadoop", "Gitea (D4Science)", "already_current", "Current scoring target already matches resolved list."),
    "C4-079": ("https://gitlab.openedition.org/lodel2/lodel-skeleton", "GitLab (OpenEdition)", "marker_only", "Current scoring uses OpenEdition/lodel GitHub mirror; GitLab repo recorded as candidate."),
}


APPLY_FIELDS = {
    "C4-006": {"host": "github", "repo": "arks-org/arks.github.io", "classification": "standard", "decision": "Resolved URL -> scoreable standard repo", "scored": "yes", "issue_tracker": "GitHub (same repo)", "issue_repo": "arks-org/arks.github.io"},
    "C4-017": {"host": "github", "repo": "brcris-ibict/brcris-interface-busca", "classification": "software", "decision": "Resolved URL", "scored": "yes", "issue_tracker": "GitHub (same repo)", "issue_repo": "brcris-ibict/brcris-interface-busca"},
    "C4-023": {"host": "github", "repo": "oacore/dashboard", "classification": "software", "decision": "Resolved URL", "scored": "yes", "issue_tracker": "GitHub (same repo)", "issue_repo": "oacore/dashboard"},
    "C4-024": {"host": "github", "repo": "Project-Counter/cop5", "classification": "standard", "decision": "Resolved URL -> scoreable standard repo", "scored": "yes", "issue_tracker": "GitHub (same repo)", "issue_repo": "Project-Counter/cop5"},
    "C4-046": {"host": "github", "repo": "fair-wizard/fair-wizard-integration-sdk", "classification": "software", "decision": "Resolved URL", "scored": "yes", "issue_tracker": "GitHub (same repo)", "issue_repo": "fair-wizard/fair-wizard-integration-sdk"},
    "C4-065": {"host": "github", "repo": "mets/METS-schema", "classification": "standard", "decision": "Resolved URL -> scoreable standard repo", "scored": "yes", "issue_tracker": "GitHub (same repo)", "issue_repo": "mets/METS-schema"},
}

MARKER_COLUMNS = [
    "resolved_url_marker",
    "resolved_url_candidate_repo",
    "resolved_url_platform",
    "resolved_url_action",
    "resolved_url_note",
]


def main():
    with open("repos.csv", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        rows = list(reader)
        fields = list(reader.fieldnames or [])

    for col in MARKER_COLUMNS:
        if col not in fields:
            fields.append(col)

    audit_rows = []
    for row in rows:
        oid = row["oisi_id"]
        if oid not in RESOLVED:
            for col in MARKER_COLUMNS:
                row.setdefault(col, "")
            continue
        candidate, platform, action, note = RESOLVED[oid]
        row.update({
            "resolved_url_marker": "yes",
            "resolved_url_candidate_repo": candidate,
            "resolved_url_platform": platform,
            "resolved_url_action": action,
            "resolved_url_note": note,
        })
        if action == "applied":
            previous = row.get("notes", "")
            row.update(APPLY_FIELDS[oid])
            row["notes"] = (previous + " | " if previous else "") + note
        audit_rows.append({
            "oisi_id": oid,
            "infrastructure": row["infrastructure"],
            "candidate_repo": candidate,
            "platform": platform,
            "action": action,
            "current_repo_after_update": row.get("repo", ""),
            "scored_after_update": row.get("scored", ""),
            "note": note,
        })

    with open("repos.csv", "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)

    with open("resolved_url_markers.csv", "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(audit_rows[0]))
        writer.writeheader()
        writer.writerows(audit_rows)


if __name__ == "__main__":
    main()
