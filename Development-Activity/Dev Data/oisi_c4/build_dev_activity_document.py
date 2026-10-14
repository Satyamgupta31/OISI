from __future__ import annotations

from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
from xml.sax.saxutils import escape


OUTPUT = Path("OISI_Development_Activity_v3_Current_Metrics_Methodology.docx")


def p(text: str = "", style: str | None = None) -> str:
    style_xml = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
    runs = []
    for part in text.split("\n"):
        runs.append(f'<w:r><w:t xml:space="preserve">{escape(part)}</w:t></w:r>')
        runs.append("<w:r><w:br/></w:r>")
    if runs:
        runs.pop()
    return f"<w:p>{style_xml}{''.join(runs)}</w:p>"


def bullet(text: str) -> str:
    return (
        '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>'
        f'<w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>'
    )


def table(rows: list[list[str]]) -> str:
    out = ['<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr>']
    for row in rows:
        out.append("<w:tr>")
        for cell in row:
            out.append(f'<w:tc><w:tcPr><w:tcW w:w="2500" w:type="dxa"/></w:tcPr>{p(cell)}</w:tc>')
        out.append("</w:tr>")
    out.append("</w:tbl>")
    return "".join(out)


def build_document() -> str:
    body: list[str] = []
    body.append(p("OISI Development Activity Metrics v3", "Title"))
    body.append(p("Current Metrics Methodology, Scoring Rules, and Reviewer Defense", "Subtitle"))
    body.append(
        p(
            "This document explains the current Development Activity metric structure used in "
            "OISI_Component4_Final_Master_With_Provenance.xlsx. It focuses on the selected metrics, "
            "their weights, their calculation rules, the data-availability threshold, and the reviewer-facing "
            "defense for each percentage."
        )
    )

    body.append(p("1. Purpose of the Framework", "Heading1"))
    body.append(
        p(
            "The Development Activity component measures whether an open infrastructure shows recent, sustained, "
            "human software-development activity, distributed contributor participation, responsiveness to issues "
            "and pull requests, and external adoption or interest. The component is intentionally code-centric: it is "
            "used only where a meaningful software repository can be identified."
        )
    )
    body.append(
        p(
            "The framework combines four dimensions. Continuity checks whether development is recent and sustained. "
            "Contributors checks whether activity depends on one person or is distributed across contributors. "
            "Responsiveness checks whether maintainers respond to issues and pull requests in a timely way. Adoption "
            "checks whether there is external attention through stars and forks."
        )
    )

    body.append(p("2. Empirical Scoring Cohort", "Heading1"))
    for item in [
        "The master set contains 112 infrastructure entries.",
        "102 entries are marked as scoreable in repos.csv and enter the scoring pipeline.",
        "72 repositories meet the methodological validity threshold of at least 75% available metric weight and receive final scores.",
        "45 of the 72 scored repositories have 100% completeness across all 13 metrics.",
        "40 master entries do not receive a final Component 4 score: 10 are excluded/not score-attempted and 30 are score-attempted but insufficient-data rows.",
        "The 10 excluded/not score-attempted entries consist of not-assessable services, standards/specifications, and one needs-decision case.",
        "Among the 30 insufficient-data rows, most fail because issue/PR volume is below the n >= 10 minimum for responsiveness metrics or because the project uses non-GitHub trackers such as Jira, GitLab, Gitea, or Bitbucket.",
    ]:
        body.append(bullet(item))

    body.append(p("3. Current Metric Structure", "Heading1"))
    body.append(
        table(
            [
                ["Dimension", "Metric", "Code", "Weight", "What it measures"],
                ["Continuity", "Human commits", "m1", "10%", "Volume of non-bot human development work."],
                ["Continuity", "Active weeks", "m2", "15%", "Consistency of development across the year."],
                ["Continuity", "Recency", "m3", "5%", "Whether human development happened near the cutoff date."],
                ["Continuity", "Release activity", "m4", "5%", "Whether release/tag activity occurred in the window."],
                ["Contributors", "Active contributors", "m5", "10%", "Number of human contributors."],
                ["Contributors", "Effective contributors", "m6", "10%", "Distribution of work across contributors."],
                ["Contributors", "Contributor retention", "m7", "5%", "Whether contributors remain active across both half-years."],
                ["Responsiveness", "Issue response rate", "m8", "10%", "Share of eligible issues receiving response within 14 days."],
                ["Responsiveness", "PR response rate", "m9", "10%", "Share of eligible PRs receiving response within 7 days."],
                ["Responsiveness", "Issue closure within 90 days", "m10", "7.5%", "Share of eligible issues closed within 90 days."],
                ["Responsiveness", "PR decision within 30 days", "m11", "7.5%", "Share of eligible PRs merged or closed within 30 days."],
                ["Adoption", "Stars gained", "m12", "2.5%", "External attention and discovery during the window."],
                ["Adoption", "Forks created", "m13", "2.5%", "Technical copying, experimentation, and contribution staging."],
            ]
        )
    )

    body.append(p("4. Dimension Weighting Logic", "Heading1"))
    body.append(
        table(
            [
                ["Dimension", "Weight", "Reason"],
                ["Continuity", "35%", "Sustained recent development is the foundation of development activity."],
                ["Contributors", "25%", "Contributor distribution matters, but it should not outweigh actual continuity and responsiveness."],
                ["Responsiveness", "35%", "Maintainer responsiveness is a strong sign of an active, governed development process."],
                ["Adoption", "5%", "Stars and forks are useful but indirect external-interest signals, so they remain low-weight."],
            ]
        )
    )

    body.append(p("5. Percentage Weight Justification and Reviewer Defense", "Heading1"))
    body.append(
        p(
            "The weights are designed to privilege direct development and maintenance behavior over indirect popularity signals. "
            "Continuity and responsiveness receive the largest dimension weights because they are closest to actual sustainability: "
            "a project must both continue producing work and respond to community development activity. Contributors receive substantial "
            "but slightly lower weight because contributor breadth is important, while adoption receives only 5% because stars and forks "
            "are indirect proxies."
        )
    )
    body.append(
        table(
            [
                ["Metric", "Weight", "Why this percentage", "Reviewer-facing defense", "If weighted higher / lower"],
                ["m1 Human commits", "10%", "Commit volume captures development intensity but can be inflated by project style.", "Use a moderate weight because volume matters, but raw quantity should not dominate over consistency or responsiveness.", "Higher: high-churn repositories dominate. Lower: real engineering activity is undercounted."],
                ["m2 Active weeks", "15%", "Continuity over time is stronger than one burst of commits.", "The highest continuity metric rewards sustained development, which is more relevant to infrastructure maintenance than raw commit count alone.", "Higher: steady low-volume repos may be over-rewarded. Lower: bursty projects could appear too strong."],
                ["m3 Recency", "5%", "Recency is a small freshness check, not a full development-quality measure.", "A recent commit indicates life, but not necessarily sustained capacity, so it stays low-weight.", "Higher: one recent commit could distort results. Lower/removal: dormant projects with old bursts may be overvalued."],
                ["m4 Release activity", "5%", "Release tags are useful but inconsistent across projects.", "Keep low because not all mature projects tag releases consistently, and some release through channels other than git tags.", "Higher: projects with tag-heavy workflows gain unfairly. Lower/removal: release discipline is ignored."],
                ["m5 Active contributors", "10%", "Contributor count captures community/developer participation breadth.", "A moderate weight recognizes contributor base size without assuming all contributors contribute equally.", "Higher: large drive-by contributor projects dominate. Lower: contributor breadth is undervalued."],
                ["m6 Effective contributors", "10%", "Effective contributors correct for concentration of work.", "Equal weight with m5 is defensible because one measures headcount and the other measures distribution.", "Higher: concentration effects dominate. Lower: single-maintainer risk is undercounted."],
                ["m7 Contributor retention", "5%", "Retention is useful but often unavailable or sparse for small projects.", "A low weight rewards continuity of contributor participation without over-penalizing smaller but active teams.", "Higher: small teams are penalized heavily. Lower/removal: contributor continuity is ignored."],
                ["m8 Issue response rate", "10%", "Timely issue response is a direct maintainer-responsiveness signal.", "10% is high enough to matter but requires n >= 10 to avoid noisy small-sample scoring.", "Higher: issue-tracker-heavy projects dominate. Lower: user/developer support is undervalued."],
                ["m9 PR response rate", "10%", "PR response measures interaction with code contributions.", "Equal weight with issue response reflects both user support and contributor support.", "Higher: PR-heavy workflows dominate. Lower: contributor responsiveness is underweighted."],
                ["m10 Issue closure within 90 days", "7.5%", "Closure captures resolution, but closure speed depends on issue complexity.", "Slightly lower than response because response is more broadly comparable than closure.", "Higher: complex projects are penalized. Lower: unresolved issue load is undercounted."],
                ["m11 PR decision within 30 days", "7.5%", "PR decision captures maintainer throughput for code contributions.", "Slightly lower than PR response because decision time depends on review complexity.", "Higher: complex review workflows are penalized. Lower: maintainers can respond but not decide without penalty."],
                ["m12 Stars gained", "2.5%", "Stars capture discovery and external interest, but not direct maintenance.", "Equal with forks under the principle of indifference: both are indirect adoption proxies.", "Higher: popularity influences sustainability too much. Lower/removal: discovery and awareness are ignored."],
                ["m13 Forks created", "2.5%", "Forks capture technical copying or experimentation, but many forks are inactive.", "Equal with stars because neither proxy proves production use; sensitivity testing shows asymmetric reweighting has negligible impact.", "Higher: dormant forks are overvalued. Lower/removal: contribution staging and technical experimentation are ignored."],
            ]
        )
    )

    body.append(p("6. Scoring Rules", "Heading1"))
    for item in [
        "All time-window metrics use 2025-10-01 00:00 UTC inclusive to 2026-10-01 00:00 UTC exclusive.",
        "Bots are excluded from human commit, contributor, issue, and PR response metrics.",
        "Merge commits are excluded from human-authorship metrics.",
        "Issue and PR responsiveness metrics require at least 10 eligible items; otherwise the metric is unavailable.",
        "Final score is shown only if at least 75% of total metric weight is available and at least half of each dimension's weight is available.",
        "If the threshold is met, component4_score equals score_on_available; otherwise score_status is insufficient data.",
    ]:
        body.append(bullet(item))

    body.append(p("7. Adoption Weighting Defense", "Heading1"))
    body.append(
        p(
            "The Adoption dimension receives only 5% of total weight and is split equally between stars and forks. "
            "This is methodologically defensible because both are indirect signals: stars reflect discovery, awareness, "
            "and screening behavior, while forks reflect code copying, experimentation, and contribution staging. Neither "
            "metric alone proves production use or maintenance."
        )
    )
    body.append(
        p(
            "Empirically, the normalized star and fork scores are highly correlated in this cohort (Spearman rho = 0.852). "
            "Sensitivity testing shows that shifting weights from 2.5%/2.5% to 1.5%/3.5% changes median score by only about "
            "0.10 points and leaves over 80% of repository ranks unchanged. Therefore equal weighting is the most transparent "
            "and least arbitrary specification."
        )
    )

    body.append(p("8. Correct Metric List for Use", "Heading1"))
    body.append(
        p(
            "The correct current metric list is m1 through m13 exactly as shown above. I do not recommend adding further metrics "
            "at this stage. The current design already covers development continuity, contributor distribution, maintainer responsiveness, "
            "and adoption. Adding more metrics would likely increase missingness and reduce the number of valid scored repositories."
        )
    )

    body.append(p("9. Research and Practice References", "Heading1"))
    refs = [
        "Borges, H., & Valente, M. T. (2018). What's in a GitHub star? Understanding repository starring practices in a social coding platform. Journal of Systems and Software, 146, 112-129. https://doi.org/10.1016/j.jss.2018.09.016",
        "Jiang, J., Lo, D., He, J., Xia, X., Kochhar, P. S., & Zhang, L. (2017). Why and how developers fork what from whom in GitHub. Empirical Software Engineering, 22, 547-578. https://doi.org/10.1007/s10664-016-9436-9",
        "Kalliamvakou, E., Gousios, G., Blincoe, K., Singer, L., German, D. M., & Damian, D. (2014). The promises and perils of mining GitHub. MSR '14, 92-101. https://doi.org/10.1145/2597073.2597074",
        "Kalliamvakou, E., Gousios, G., Blincoe, K., Singer, L., German, D. M., & Damian, D. (2016). An in-depth study of the promises and perils of mining GitHub. Empirical Software Engineering, 21, 2035-2071. https://doi.org/10.1007/s10664-015-9393-5",
        "OECD / European Commission Joint Research Centre. (2008). Handbook on Constructing Composite Indicators: Methodology and User Guide. OECD Publishing. https://doi.org/10.1787/9789264043466-en",
        "CHAOSS Project. CHAOSS Metrics. Community health metrics for open source projects, including contribution and responsiveness concepts. https://github.com/chaoss/metrics",
        "Gousios, G., Pinzger, M., & van Deursen, A. (2014). An exploratory study of the pull-based software development model. ICSE '14, 345-355. https://doi.org/10.1145/2568225.2568260",
    ]
    for ref in refs:
        body.append(bullet(ref))

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:body>{"".join(body)}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>'
    )


def styles() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:rPr><w:b/><w:sz w:val="36"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:rPr><w:i/><w:sz w:val="26"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:rPr><w:b/><w:sz w:val="30"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:rPr><w:b/><w:sz w:val="26"/></w:rPr></w:style>
  <w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr></w:style>
</w:styles>"""


def numbering() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/></w:lvl></w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>"""


def write_docx():
    with ZipFile(OUTPUT, "w", ZIP_DEFLATED) as z:
        z.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>""",
        )
        z.writestr(
            "_rels/.rels",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>""",
        )
        z.writestr(
            "word/_rels/document.xml.rels",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>""",
        )
        z.writestr("word/document.xml", build_document())
        z.writestr("word/styles.xml", styles())
        z.writestr("word/numbering.xml", numbering())


if __name__ == "__main__":
    write_docx()
