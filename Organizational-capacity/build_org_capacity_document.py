from __future__ import annotations

from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
from xml.sax.saxutils import escape


OUTPUT = Path("OISI_Organizational_Capacity_v3_Current_Metrics_Methodology.docx")


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
    body.append(p("OISI Organizational Capacity Metrics v3", "Title"))
    body.append(p("Current Metrics Methodology and Calculation Justification", "Subtitle"))
    body.append(
        p(
            "This document explains the current Organizational Capacity metric structure used in "
            "OISI_Organizational_Capacity_Scoring_v3.xlsx. It focuses on why each metric is relevant, "
            "what role it plays, how it is calculated, why the percentage weight is assigned, and how "
            "the metric affects the final result."
        )
    )

    body.append(p("1. Purpose of the Framework", "Heading1"))
    body.append(
        p(
            "The Organizational Capacity score measures whether an infrastructure has enough technical maturity, "
            "maintenance practice, institutional support, and operational delivery capacity to support continuity. "
            "For open scholarly infrastructure, capacity is not only about whether software exists. It is also about "
            "whether the service is maintained, institutionally supported, and practically deliverable over time."
        )
    )
    body.append(
        p(
            "The framework is built around three questions: Is the technology mature and maintained? Is there an "
            "organizational base that can hold responsibility for the service? Are there operational channels through "
            "which the service can be delivered or supported? These questions align with open infrastructure principles "
            "that emphasize sustainability, transparent operations, continuity, and risk management."
        )
    )

    body.append(p("2. Current Metric Structure", "Heading1"))
    body.append(
        table(
            [
                ["Code", "Metric", "Field used", "Weight", "What it measures"],
                ["M1", "Technology readiness", "079", "10%", "Whether the technology is proven enough for reliable use."],
                ["M2", "Maintenance status", "080", "20%", "Whether the infrastructure is actively maintained."],
                ["M3", "Infrastructure age", "056", "5%", "A bounded maturity signal based on time since launch."],
                ["I1", "Formal organizational host", "051", "15%", "Whether there is an organizational form behind the infrastructure."],
                ["I2", "Affiliations / institutional support", "058", "15%", "Whether named institutions or supporters are attached."],
                ["I3", "SCOSS participation", "099", "5%", "Whether the infrastructure has recognized collective-support participation."],
                ["O1", "Hosted/SaaS service availability", "076", "20%", "Whether users can access hosted or service-delivery options."],
                ["O2", "Registered service-provider ecosystem", "077", "10%", "Whether third-party delivery has a visible provider ecosystem."],
            ]
        )
    )

    body.append(p("3. Why These Metrics Are Relevant", "Heading1"))
    body.append(
        p(
            "The selected metrics combine technical, institutional, and operational evidence. Technology readiness and "
            "maintenance show whether the infrastructure is usable and cared for. Organizational host and affiliations show "
            "whether responsibility is embedded in an identifiable organization or support network. Hosted service availability "
            "and provider ecosystems show whether the infrastructure can be practically operated, adopted, or supported by users."
        )
    )
    body.append(
        p(
            "This avoids using a single proxy for capacity. A mature technology without maintenance is risky. A maintained "
            "project without institutional support may be fragile. A project with institutional support but no operational delivery "
            "channel may still be hard to adopt. The score therefore combines several complementary signals."
        )
    )

    body.append(p("4. Weighting Logic", "Heading1"))
    body.append(
        table(
            [
                ["Sub-component", "Metrics", "Total weight", "Reason"],
                [
                    "Maturity & Maintenance",
                    "M1, M2, M3",
                    "35%",
                    "Technical maturity and maintenance are essential, but the dataset shows readiness and maintenance are common, so the block is not allowed to dominate.",
                ],
                [
                    "Institutional Support",
                    "I1, I2, I3",
                    "35%",
                    "Organizational backing and institutional relationships are central to continuity and accountability.",
                ],
                [
                    "Operational Capacity",
                    "O1, O2",
                    "30%",
                    "Hosted delivery and service-provider ecosystems measure whether capacity can be exercised in practice.",
                ],
            ]
        )
    )
    body.append(
        p(
            "The largest individual weights are given to Maintenance status and Hosted/SaaS service availability, each at 20%. "
            "Maintenance is the clearest continuity signal: an infrastructure cannot be capacity-strong if it is not maintained. "
            "Hosted/SaaS availability is also important because many users need an operational service channel, not only code or documentation."
        )
    )

    body.append(p("5. Percentage Weight Justification and Reviewer Defense", "Heading1"))
    body.append(
        p(
            "The percentages are not arbitrary. They are designed to balance construct validity and empirical behavior in the dataset. "
            "Construct validity means the score should measure organizational capacity, not only technical readiness or organizational type. "
            "Empirical behavior means the weights should not over-reward fields that are nearly constant across the dataset. The final design "
            "therefore gives the highest weight to indicators that are both conceptually important and practically meaningful for continuity."
        )
    )
    body.append(
        table(
            [
                ["Metric", "Weight", "Why this percentage", "Reviewer-facing defense", "If weighted higher / lower"],
                [
                    "M1 Technology readiness",
                    "10%",
                    "Readiness is necessary, but the field is near-saturated in the dataset, so a high weight would inflate most scores without improving discrimination.",
                    "We keep technology readiness because an unproven system should not score as organizationally mature, but we cap its influence because most services already report high readiness.",
                    "Higher: rankings would mostly reward a field where nearly everyone scores well. Lower: technical maturity would become underrepresented.",
                ],
                [
                    "M2 Maintenance status",
                    "20%",
                    "Maintenance is the strongest direct continuity signal. It deserves the largest maturity weight because active maintenance is central to ongoing capacity.",
                    "A service cannot be considered capacity-strong if it is not maintained. This justifies a substantial weight even though many services are actively maintained.",
                    "Higher: the score would become too maintenance-dominated. Lower: weakly maintained services could score too well through institutional or service-delivery signals.",
                ],
                [
                    "M3 Infrastructure age",
                    "5%",
                    "Age is useful but weak evidence. Older services may be stable, but age alone does not prove quality, maintenance, governance, or support.",
                    "The small weight recognizes maturity over time while preventing old infrastructures from receiving automatic advantage.",
                    "Higher: old services would be over-rewarded. Lower/removal: early-stage risk would be ignored completely.",
                ],
                [
                    "I1 Formal organizational host",
                    "15%",
                    "A formal host supports responsibility, contracting, continuity, and accountability, but it should not dominate over actual maintenance or operations.",
                    "The metric captures whether there is an identifiable organizational base behind the infrastructure. It is weighted equally with affiliations because both represent institutional support from different angles.",
                    "Higher: legal/organizational form would overshadow practical capacity. Lower: the score would understate stewardship and accountability.",
                ],
                [
                    "I2 Affiliations / institutional support",
                    "15%",
                    "Affiliations indicate external support, legitimacy, and resilience, but the count of affiliations is not rewarded to avoid size bias.",
                    "The weight reflects that institutional relationships matter for sustainability, but presence is scored rather than volume, so large organizations are not automatically favored.",
                    "Higher: projects with institutional names would dominate. Lower: institutional resilience would be undercounted.",
                ],
                [
                    "I3 SCOSS participation",
                    "5%",
                    "SCOSS is a meaningful support signal but rare and not expected for all valid infrastructures.",
                    "A small weight treats SCOSS participation as a recognized sustainability signal without making SCOSS status a requirement for organizational capacity.",
                    "Higher: non-SCOSS services would be unfairly penalized. Lower/removal: recognized collective-support evidence would be ignored.",
                ],
                [
                    "O1 Hosted/SaaS service availability",
                    "20%",
                    "Operational delivery is a major capacity signal. Users often need a working service pathway, not only software or documentation.",
                    "The weight is high because organizational capacity includes the ability to deliver or support use. Structurally not-applicable cases are excluded from the applicable denominator.",
                    "Higher: service-delivery models would dominate the score. Lower: practical operational capacity would be understated.",
                ],
                [
                    "O2 Registered service-provider ecosystem",
                    "10%",
                    "Provider ecosystems indicate adoption and support capacity, but the field is sparse and applies only to certain service models.",
                    "The metric is weighted enough to matter where relevant, but not so much that services without a provider ecosystem are unfairly treated when the model is not applicable.",
                    "Higher: third-party-provider models would be favored. Lower/removal: external operational resilience would be undercounted.",
                ],
            ]
        )
    )
    body.append(
        p(
            "If a reviewer asks why the component uses 35% Maturity & Maintenance, 35% Institutional Support, and 30% Operational Capacity, "
            "the answer is that organizational capacity requires all three dimensions. Maturity and maintenance establish whether the infrastructure "
            "is technically dependable. Institutional support establishes whether responsibility and continuity are anchored. Operational capacity "
            "establishes whether the service can actually be delivered or supported. The split is intentionally balanced, with slightly less weight "
            "on operational capacity because some infrastructures are not designed around hosted/SaaS or third-party-provider models."
        )
    )

    body.append(p("6. Metric-by-Metric Explanation", "Heading1"))
    sections = [
        (
            "M1. Technology Readiness - 10%",
            "Technology readiness measures whether the system is technically mature enough to be relied upon.",
            "079_technology_readiness",
            "Actual system proven in operational environment = 1.00; System complete and qualified = 8/9; Technology validated in relevant environment = 5/9; otherwise NA.",
            "10% is appropriate because readiness matters, but in this dataset most services are already at the highest readiness level. A higher weight would inflate scores without adding much discrimination.",
            "A service marked 'Actual system proven in operational environment' receives full M1 credit.",
        ),
        (
            "M2. Maintenance Status - 20%",
            "Maintenance status measures whether the infrastructure is actively cared for. It is the strongest direct continuity metric.",
            "080_maintenance_status",
            "Actively Maintained = 1.00; Minimally Maintained = 0.50; inactive/not maintained = 0.00; unknown/blank = NA.",
            "20% reflects that maintenance is central to organizational capacity. A mature technology can still be risky if maintenance is weak.",
            "A project that is actively maintained receives full M2 credit; minimally maintained projects receive partial credit.",
        ),
        (
            "M3. Infrastructure Age - 5%",
            "Age is a small maturity indicator. A service that has existed for several years may have passed some early-stage risk, but age alone does not prove quality or sustainability.",
            "056_launch_date",
            "Age = 2026 minus launch year. Score = min(age, 15) / 15. Full age credit is reached at 15 years.",
            "5% keeps age bounded. Older infrastructures should not automatically outrank newer but better-supported services.",
            "A service launched in 2020 has age 6 in the 2026 scoring year, so M3 = 6/15 = 0.40.",
        ),
        (
            "I1. Formal Organizational Host - 15%",
            "This metric checks whether the infrastructure has a recognizable organizational form or host that can carry responsibility.",
            "051_business_form",
            "Non-profit, academic fiscal sponsorship, non-profit fiscal sponsorship, commercial organization, or similar formal organizational backing = 1.00; volunteer community = 0.50; Other/blank = NA.",
            "15% reflects that organizational form matters for continuity, contracting, accountability, and stewardship. Volunteer communities are not treated as zero because they can have capacity, but they receive partial credit because responsibility is less formalized.",
            "A fiscally sponsored infrastructure receives full I1 credit.",
        ),
        (
            "I2. Affiliations / Institutional Support - 15%",
            "Affiliations measure whether the infrastructure has named institutional relationships or support structures beyond its own project identity.",
            "058_affiliations",
            "Named affiliation or institutional support = 1.00; explicit No affiliations = 0.00; blank = NA.",
            "15% is used because institutional support can provide resilience, legitimacy, expertise, and continuity. The score rewards presence of support, not the number of affiliations.",
            "A service with a named university, library, foundation, or partner institution receives full I2 credit.",
        ),
        (
            "I3. SCOSS Participation - 5%",
            "SCOSS participation is treated as a small signal of recognized collective support for open science infrastructure.",
            "099_scoss",
            "TRUE = 1.00; FALSE = 0.00; unknown/blank = NA.",
            "5% keeps SCOSS as a recognition/support signal without making it a requirement. Most legitimate infrastructures are not SCOSS-supported, so this should not dominate the score.",
            "A SCOSS-supported infrastructure receives a small institutional-support boost.",
        ),
        (
            "O1. Hosted/SaaS Service Availability - 20%",
            "This metric measures whether users can access the infrastructure through a hosted, SaaS, solution-provider, or third-party service route.",
            "076_saas_options",
            "Through solution provider only, through third-party vendor only, or through solution provider/third-party vendor = 1.00; No services available = 0.00; Not applicable/Unknown/blank = NA.",
            "20% reflects the importance of operational delivery. A system may be technically mature, but organizational capacity is stronger when there are practical routes for use and support.",
            "If SaaS is not structurally relevant to a service, the metric is treated as not applicable and removed from the applicable denominator.",
        ),
        (
            "O2. Registered Service-Provider Ecosystem - 10%",
            "This metric measures whether third-party service delivery has a visible provider ecosystem.",
            "077_registered_service_providers",
            "Applicable only when 076 indicates third-party service delivery. Provider evidence present = 1.00; third-party delivery but no provider list = 0.00; otherwise NA.",
            "10% reflects that provider ecosystems can increase adoption capacity and resilience, but the field is sparse and not relevant to all infrastructure types.",
            "A project that lists certified or registered providers receives full O2 credit where third-party delivery applies.",
        ),
    ]
    for title, role, fields, calculation, weight, example in sections:
        body.append(p(title, "Heading2"))
        body.append(p("Role: " + role))
        body.append(p("Field used: " + fields))
        body.append(p("Calculation: " + calculation))
        body.append(p("Why this percentage: " + weight))
        body.append(p("Example: " + example))

    body.append(p("7. Final Score Calculation", "Heading1"))
    body.append(
        p(
            "Each metric is scored from 0 to 1. The final score is:\n"
            "Organizational Capacity Score = 100 x sum(metric score x metric weight) / sum(scored metric weights)."
        )
    )
    body.append(
        p(
            "Completeness is calculated against applicable metric weight. This matters because some operational metrics are structurally "
            "not relevant to every service. For example, if hosted/SaaS delivery is marked Not applicable, O1 is removed from the applicable "
            "denominator. If third-party delivery is not applicable, O2 is also removed."
        )
    )
    body.append(
        p(
            "A result is shown only when applicable completeness is at least 70% and at least two sub-components are present. "
            "Eligible results are Reportable when completeness is at least 85%; otherwise they are Provisional."
        )
    )

    body.append(p("8. Worked Example: 2i2c", "Heading1"))
    body.append(
        table(
            [
                ["Metric", "Score", "Reason"],
                ["M1", "1.00", "Technology is proven in an operational environment."],
                ["M2", "1.00", "Maintenance status is actively maintained."],
                ["M3", "0.40", "Launch year 2020; age in 2026 is 6 years; 6/15 = 0.40."],
                ["I1", "1.00", "Business form is fiscal sponsorship by a non-profit."],
                ["I2", "NA", "No affiliation value is supplied."],
                ["I3", "0.00", "SCOSS is FALSE."],
                ["O1", "NA", "SaaS/service option is Not applicable, so it is removed from applicable weight."],
                ["O2", "NA", "Third-party provider ecosystem is not applicable."],
                ["Final score", "85.45", "Eligible but Provisional because applicable completeness is 78.57%."],
            ]
        )
    )

    body.append(p("9. How to Interpret Results", "Heading1"))
    for item in [
        "High maturity but low institutional support suggests the technology works but organizational continuity may need review.",
        "High institutional support but low operational capacity suggests the organization exists, but service delivery may be limited.",
        "High operational capacity with low maintenance would be a warning sign because delivery exists but upkeep may be weak.",
        "A Provisional score is usable for analysis but should be interpreted with attention to missing applicable evidence.",
        "Lower and upper bounds show uncertainty by treating missing applicable metrics as 0 and 1 respectively.",
    ]:
        body.append(bullet(item))

    body.append(p("10. Research and Practice References", "Heading1"))
    references = [
        "NASA. Technology Readiness Levels. NASA describes TRL as a measurement system for assessing the maturity level of a technology, supporting the use of readiness as a technical maturity metric. https://www.nasa.gov/directorates/somd/space-communications-navigation-program/technology-readiness-levels/",
        "NASA. Technology Readiness Levels Demystified. NASA explains the nine-level TRL scale and gives TRL 9 as technology proven in operations. https://www.nasa.gov/aeronautics/technology-readiness-levels-demystified/",
        "Principles of Open Scholarly Infrastructure. POSI emphasizes sustainability, transparent operations, accountability, and risk/insurance mechanisms for scholarly infrastructure. https://openscholarlyinfrastructure.org/",
        "SCOSS - Global Sustainability Coalition for Open Science Services. SCOSS is a network focused on securing the future of open infrastructure and mainstreaming support for non-commercial open science infrastructure. https://scoss.org/",
        "Bilder, G., Lin, J., & Neylon, C. Principles for Open Scholarly Infrastructures. The original POSI framing highlights governance, sustainability, and insurance as core requirements for infrastructure reliability. https://scholcommlab.ca/wp-content/uploads/2018/11/PrinciplesforOpenScholarlyInfrastructures.pdf",
        "Invest in Open Infrastructure, Infra Finder dataset. The score is built from fields represented in the Infra Finder organizational-capacity data model. https://zenodo.org/records/18225266",
        "Marsolo et al. Ensuring Support for Research and Quality Improvement Networks: Four Pillars of Sustainability. The article emphasizes infrastructure support, maintenance, management structures, adaptive capacity, and institutionalization for long-term network sustainability. https://pmc.ncbi.nlm.nih.gov/articles/PMC4371504/",
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
