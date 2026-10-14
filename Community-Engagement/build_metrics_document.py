from __future__ import annotations

from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
from xml.sax.saxutils import escape


OUTPUT = Path("OISI_Community_Engagement_v3_Current_Metrics_Methodology.docx")


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
            out.append(f'<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>{p(cell)}</w:tc>')
        out.append("</w:tr>")
    out.append("</w:tbl>")
    return "".join(out)


def build_document() -> str:
    body: list[str] = []

    body.append(p("OISI Community Engagement Metrics v3", "Title"))
    body.append(p("Current Metrics Methodology and Calculation Justification", "Subtitle"))
    body.append(
        p(
            "This document explains the current Community Engagement metric structure used in "
            "OISI_Community_Engagement_v3_multi_sheet.xlsx. It focuses only on the selected metrics: "
            "why each metric is relevant, what role it plays, how it is calculated, why its percentage "
            "weight is appropriate, and how it affects the final result."
        )
    )

    body.append(p("1. Purpose of the Framework", "Heading1"))
    body.append(
        p(
            "The purpose of the Community Engagement score is to measure whether an open scholarly or "
            "open-source infrastructure gives its community practical, visible, and safe ways to participate. "
            "A strong community-engagement system should show three things: first, the infrastructure actively "
            "engages with its community; second, it offers concrete ways for external participants to contribute; "
            "and third, it maintains norms that make participation safer and more predictable."
        )
    )
    body.append(
        p(
            "This approach is consistent with community-engagement literature, which treats engagement as a "
            "combination of relationship-building, participation mechanisms, communication, shared capacity, "
            "and evaluation of community involvement. It also aligns with open-source community-health work, "
            "where contribution pathways, documentation, activity, safety norms, and transparent participation "
            "channels are treated as signals of project health and sustainability."
        )
    )

    body.append(p("2. Current Metric Structure", "Heading1"))
    body.append(
        table(
            [
                ["Code", "Metric", "Fields used", "Weight", "What it measures"],
                ["E1", "Engagement availability", "037", "10%", "Whether engagement exists as an identifiable practice."],
                ["E2", "Engagement activity breadth", "038 + 938", "25%", "How many different engagement channels are available."],
                ["E3", "Engagement evidence strength", "116 + 938", "5%", "Whether engagement claims have visible supporting evidence."],
                ["C1", "Contribution opportunity breadth", "040 + 940", "25%", "How many practical contribution routes are available."],
                ["C2", "Contribution guidance", "041", "15%", "Whether contributors receive actionable instructions."],
                ["C3", "Open contribution infrastructure", "019", "5%", "Whether there is a visible technical space for contribution."],
                ["N1", "Community norms / Code of Conduct", "042 + 043", "15%", "Whether participation norms and safety expectations are visible."],
            ]
        )
    )

    body.append(p("3. Why These Metrics Are Relevant", "Heading1"))
    body.append(
        p(
            "The selected metrics cover the full pathway from engagement to contribution. E1, E2, and E3 ask "
            "whether the community can see and access engagement mechanisms. C1, C2, and C3 ask whether interested "
            "participants can move from engagement into actual contribution. N1 asks whether the environment has "
            "basic behavioral norms that support inclusive participation."
        )
    )
    body.append(
        p(
            "This structure avoids relying on a single proxy. For example, a public GitHub repository is useful, "
            "but it does not prove community engagement by itself. A code of conduct is important, but it does not "
            "prove contribution pathways exist. A broad activity list is valuable, but without guidance and evidence "
            "it may be difficult for a new participant to act. The score therefore combines multiple complementary "
            "signals."
        )
    )

    body.append(p("4. Weighting Logic", "Heading1"))
    body.append(
        table(
            [
                ["Sub-component", "Metrics", "Total weight", "Reason"],
                [
                    "Engagement & Participation",
                    "E1, E2, E3",
                    "40%",
                    "Engagement is the foundation of the component: services need visible practices, multiple interaction channels, and evidence.",
                ],
                [
                    "Contribution Mechanisms",
                    "C1, C2, C3",
                    "45%",
                    "Contribution mechanisms receive the largest combined weight because community engagement should lead to practical participation.",
                ],
                [
                    "Community Norms",
                    "N1",
                    "15%",
                    "Norms are essential for safe participation, but they are enabling conditions rather than the whole engagement system.",
                ],
            ]
        )
    )
    body.append(
        p(
            "The highest individual weights are assigned to E2 and C1, each at 25%, because these are the strongest "
            "direct indicators of breadth. E2 measures the range of engagement activities; C1 measures the range of "
            "contribution opportunities. Together, they answer the most important practical question: can different "
            "types of community members find meaningful ways to participate?"
        )
    )

    body.append(p("5. Percentage Weight Justification and Reviewer Defense", "Heading1"))
    body.append(
        p(
            "The percentages are designed to balance conceptual importance, auditability, and dataset behavior. "
            "Community engagement should not be reduced to one visible signal, such as a code of conduct or a repository. "
            "The weighting therefore gives the greatest influence to the two strongest participation-breadth measures: "
            "engagement activity breadth and contribution opportunity breadth. Lower weights are used for supporting "
            "signals that matter but should not dominate the score by themselves."
        )
    )
    body.append(
        table(
            [
                ["Metric", "Weight", "Why this percentage", "Reviewer-facing defense", "If weighted higher / lower"],
                [
                    "E1 Engagement availability",
                    "10%",
                    "This is a baseline signal showing whether engagement exists as a stated practice, but it is weaker than actual activity evidence.",
                    "We reward the existence of engagement mechanisms, but avoid letting a status claim dominate the component. A service must still show activities, contribution routes, guidance, and norms.",
                    "Higher: services could score well based mainly on a self-reported status. Lower: the framework would under-recognize whether engagement exists at all.",
                ],
                [
                    "E2 Engagement activity breadth",
                    "25%",
                    "Breadth of engagement activities is one of the strongest direct indicators of community engagement capacity.",
                    "Community engagement is strongest when different participants can engage through different channels: calls, meetings, forums, training, user groups, advisory groups, blogs, and research.",
                    "Higher: engagement activity lists would dominate the score. Lower: a core measure of actual engagement practice would be underweighted.",
                ],
                [
                    "E3 Engagement evidence strength",
                    "5%",
                    "Evidence is important for verification, but it is a supporting auditability signal rather than a substantive engagement mechanism by itself.",
                    "The metric reduces reliance on unsupported claims, but it remains low-weight because a URL alone does not prove strong engagement.",
                    "Higher: projects with better documentation could outrank projects with stronger actual practices. Lower/removal: the score would be less auditable.",
                ],
                [
                    "C1 Contribution opportunity breadth",
                    "25%",
                    "Contribution opportunities are the clearest sign that the community can participate practically, not only observe.",
                    "This metric is weighted equally with E2 because engagement and contribution are the two central dimensions of community participation.",
                    "Higher: contribution routes would overshadow engagement activities and norms. Lower: practical participation would be undercounted.",
                ],
                [
                    "C2 Contribution guidance",
                    "15%",
                    "Guidance makes contribution opportunities usable. Without instructions, contribution routes may exist formally but remain difficult for outsiders.",
                    "The weight reflects the importance of onboarding and actionable documentation while keeping breadth of opportunities as the stronger signal.",
                    "Higher: documentation quality could outweigh the presence of actual opportunities. Lower: projects with vague or inaccessible contributor pathways would be over-rewarded.",
                ],
                [
                    "C3 Open contribution infrastructure",
                    "5%",
                    "Open technical infrastructure supports contribution, but repository availability alone does not prove community engagement.",
                    "The low weight recognizes infrastructure as an enabling condition. It avoids over-rewarding projects simply for having a public repository.",
                    "Higher: software-heavy projects would gain an unfair advantage. Lower/removal: the technical contribution pathway would be ignored.",
                ],
                [
                    "N1 Community norms / Code of Conduct",
                    "15%",
                    "Participation norms support safety, inclusion, and predictability, which are necessary conditions for healthy engagement.",
                    "The metric is important enough to materially affect the score, but it does not replace actual engagement activities or contribution pathways.",
                    "Higher: a code of conduct could overpower evidence of actual participation. Lower: safety and inclusion norms would be undervalued.",
                ],
            ]
        )
    )
    body.append(
        p(
            "If a reviewer asks why the component uses 40% Engagement & Participation, 45% Contribution Mechanisms, and 15% Community Norms, "
            "the answer is that the component measures engagement as a pathway from visibility to action. Engagement & Participation receives "
            "40% because services need visible ways to interact with their communities. Contribution Mechanisms receives 45% because engagement "
            "should translate into practical participation. Community Norms receives 15% because norms are essential enabling conditions, but "
            "they should not outweigh evidence of actual engagement and contribution."
        )
    )

    metric_sections = [
        (
            "E1. Engagement Availability - 10%",
            "E1 checks whether the infrastructure reports community engagement as an available or planned practice. It is a baseline signal: without an engagement status, later activity evidence is harder to interpret.",
            "037_community_engagement_status",
            "Implemented/available = 1.00; In progress = 0.67; Considering = 0.33; Not planning = 0.00; Unknown or blank = NA.",
            "10% is enough to reward the existence of an engagement mechanism, but not enough to let a simple status claim dominate the score. A service must still show activity, contribution routes, guidance, and norms.",
            "If a service reports Implemented/available, it receives full E1 credit. If it is only Considering engagement, it receives partial credit because the practice is not yet established.",
        ),
        (
            "E2. Engagement Activity Breadth - 25%",
            "E2 measures the range of community engagement activities. It captures whether engagement happens through multiple channels such as meetings, community calls, blogs, forums, advisory/user groups, training, sprints, user research, or ambassador networks.",
            "038_community_engagement_activities and 938_community_engagement_activities_other",
            "Distinct categories from 038 and coded 938 are combined. Score = min(number of categories, 8) / 8.",
            "25% reflects that breadth is one of the clearest indicators of real engagement capacity. A project with multiple channels can serve different participant types: users, contributors, maintainers, funders, trainers, and advisory participants.",
            "2i2c has a blank 038 field, but its 938 text describes community calls/calendar evidence and stakeholder/advisory-style roles. The current method gives credit for that evidence instead of treating E2 as missing.",
        ),
        (
            "E3. Engagement Evidence Strength - 5%",
            "E3 measures whether engagement claims are supported by visible evidence, such as a community engagement URL or meaningful free-text evidence.",
            "116_community_engagement_url and 938_community_engagement_activities_other",
            "URL present = 1.00; only free-text evidence = 0.50; no evidence = NA.",
            "5% keeps this as an auditability metric. Evidence matters because it helps reviewers verify claims, but it should not outweigh the substantive activity and contribution metrics.",
            "A service with a public community page receives full evidence credit. A service with only descriptive text receives partial evidence credit.",
        ),
        (
            "C1. Contribution Opportunity Breadth - 25%",
            "C1 measures the number of different contribution routes available to community members. Examples include code, documentation, education/training, funds, user research/testing, working groups, and other listed opportunities.",
            "040_contribution_opportunities and 940_contribution_opportunities_other",
            "Distinct contribution categories are counted. 940 is counted only when it contains actual other-text evidence. Score = min(number of categories, 6) / 6.",
            "25% is assigned because contribution opportunities are the practical bridge between community engagement and community participation. A service should not only communicate with its community; it should let people contribute in concrete ways.",
            "A service listing code, documentation, working groups, governance, user testing, and education/training receives full C1 credit. A service listing only code contribution receives limited credit.",
        ),
        (
            "C2. Contribution Guidance - 15%",
            "C2 measures whether a potential contributor can understand how to contribute. Contribution opportunities become more usable when they are supported by instructions, guidelines, or a clear participation page.",
            "041_contribution_guidelines",
            "Dedicated contribution guide = 1.00; related contribution or participation page = 0.50; blank or suspect URL = NA pending review.",
            "15% reflects that guidance reduces participation barriers. Research and practice in open source health emphasize that clear documentation and onboarding help people move from interest to action.",
            "A CONTRIBUTING.md file receives full credit. A broader community page that mentions participation but is not a step-by-step guide receives partial credit.",
        ),
        (
            "C3. Open Contribution Infrastructure - 5%",
            "C3 checks whether a technical contribution space is visible, such as a public repository. This is especially relevant for infrastructures with software, standards, or technical documentation.",
            "019_open_code_repository_url",
            "Public forge host such as GitHub, GitLab, Bitbucket, or Codeberg = 1.00; other repository host = 0.50; blank = NA.",
            "5% is appropriate because repository availability is useful but not sufficient. A public repo can support contribution, but engagement also requires guidance, opportunities, communication, and norms.",
            "2i2c receives full C3 credit because it has a GitHub repository URL.",
        ),
        (
            "N1. Community Norms / Code of Conduct - 15%",
            "N1 measures whether the community has visible behavioral expectations. This supports safer participation and gives community members clearer expectations about acceptable conduct.",
            "042_code_of_conduct_status and 043_code_of_conduct_url",
            "Implemented with URL = 1.00; implemented without URL = 0.75; in progress = 0.67; considering = 0.33; not planning = 0.00; unknown/blank = NA.",
            "15% reflects that norms are a major enabling condition for inclusive participation. They matter for trust and safety, but they do not replace the need for activities and contribution mechanisms.",
            "A service with an implemented code of conduct and public URL receives full N1 credit.",
        ),
    ]

    body.append(p("6. Metric-by-Metric Explanation", "Heading1"))
    for title, role, fields, calculation, weight, example in metric_sections:
        body.append(p(title, "Heading2"))
        body.append(p("Role: " + role))
        body.append(p("Fields used: " + fields))
        body.append(p("Calculation: " + calculation))
        body.append(p("Why this percentage: " + weight))
        body.append(p("Example: " + example))

    body.append(p("7. Final Score Calculation", "Heading1"))
    body.append(
        p(
            "Each metric is scored from 0 to 1. The final score is a weighted average over the scored metrics:\n"
            "Community Engagement Score = 100 x sum(metric score x metric weight) / sum(scored metric weights)."
        )
    )
    body.append(
        p(
            "Missing data is excluded from both numerator and denominator. This is important because a blank field is not always a verified absence. "
            "However, the score is shown only when there is enough information: at least 60% of the metric weight must be scored, and at least one metric "
            "must be available in each of the three sub-components."
        )
    )
    body.append(
        p(
            "Result status is based on completeness. A score with 85% or more scored weight is treated as Reportable. A score meeting the minimum "
            "threshold but below 85% is Provisional. This helps separate strong results from results that are useful but still affected by missing data."
        )
    )

    body.append(p("8. Worked Example: 2i2c", "Heading1"))
    body.append(
        table(
            [
                ["Metric", "Value", "Reason"],
                ["E1", "1.00", "Engagement status is Implemented/available."],
                ["E2", "0.25", "938 text supports two engagement categories: Community calls and Interest/user/advisory groups."],
                ["E3", "1.00", "Engagement evidence URLs are present."],
                ["C1", "NA", "No contribution opportunities are listed in the supplied 040/940 fields."],
                ["C2", "NA", "No contribution guidance URL is listed in 041."],
                ["C3", "1.00", "A GitHub repository URL is present."],
                ["N1", "1.00", "Code of Conduct is implemented and URL is present."],
                ["Final", "68.75", "The score is eligible but provisional because some contribution fields are missing."],
            ]
        )
    )

    body.append(p("9. How the Metrics Affect Interpretation", "Heading1"))
    for item in [
        "High E2 but low C1 means the service communicates with its community but may not provide many contribution routes.",
        "High C1 but low C2 means opportunities exist but contributor onboarding may be unclear.",
        "High C3 alone should not be read as strong engagement; repository availability is only one small part of participation capacity.",
        "High N1 supports trust and inclusion, but it does not substitute for engagement activities or contribution pathways.",
        "A provisional score should be interpreted cautiously because missing fields can affect the denominator and uncertainty range.",
    ]:
        body.append(bullet(item))

    body.append(p("10. Research and Practice References", "Heading1"))
    references = [
        "Principles of Community Engagement, 3rd Edition. This guide frames engagement as a practice involving sustained relationships, communication, participation, and evaluation of community involvement. https://health.ucdavis.edu/media-resources/crhd/documents/pdfs/pce-3rd-edition.pdf",
        "CHAOSS Metrics. CHAOSS develops implementation-agnostic metrics for open source community health, including contribution, activity, and community practices. https://github.com/chaoss/metrics",
        "CHAOSS Code of Conduct for a Project metric. CHAOSS treats code-of-conduct presence, clarity, enforcement procedures, and transparency as community-health signals. https://chaoss.community/kb/metric-code-of-conduct-for-a-project/",
        "CHAOSS Data Use Statement. CHAOSS emphasizes transparency and actionability in project health metrics so stakeholders can make better engagement decisions. https://github.com/chaoss/community/blob/main/data-use-statement.md",
        "Mozilla and Open Tech Strategies. Open Source Archetypes. This report explains that open-source participation models differ by project goals, and that contribution models should be evaluated according to the project’s intended community role. https://blog.mozilla.org/wp-content/uploads/2018/05/MZOTS_OS_Archetypes_report_ext_scr.pdf",
        "Mozilla Community Participation Guidelines. Mozilla presents participation guidelines as a mechanism for respectful and inclusive contribution spaces. https://www.mozilla.org/en-US/about/governance/policies/participation/",
        "Principles of Open Scholarly Infrastructure. POSI emphasizes community accountability, sustainability, and transparent stewardship for scholarly infrastructure. https://openscholarlyinfrastructure.org/",
        "Invest in Open Infrastructure, Infra Finder dataset. The score uses the community-engagement fields represented in the Infra Finder data model. https://zenodo.org/records/18225266",
        "Link, G. J. P. (2020). CHAOSS and open source health metrics. USENIX ;login:. This article discusses why shared practical metrics are useful for understanding open source community health. https://www.usenix.org/system/files/login/articles/login_fall20_07_link.pdf",
        "Wynn, D. E. (2022). How to characterize the health of an Open Source Software project? A snowball literature review of an emerging practice. The review supports treating project health as multi-dimensional rather than reducible to code activity alone. https://arxiv.org/abs/2208.01105",
    ]
    for ref in references:
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
