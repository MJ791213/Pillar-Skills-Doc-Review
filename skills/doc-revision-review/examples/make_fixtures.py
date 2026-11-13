#!/usr/bin/env python3
"""Generate the Rev 1 / Rev 2 fixture pair used to test this skill.

Usage:
    python3 make_fixtures.py --out <dir>

Writes DES-9999_Rev1.docx and DES-9999_Rev2.docx. Rev 2 carries one deliberate
defect per check; examples/expected-findings.md records what a correct review
reports for this pair.

These are synthetic documents written for testing. They are not Pillar
documents and describe no real product.
"""
import argparse
import sys
from pathlib import Path

try:
    import docx
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Pt
except ImportError:  # pragma: no cover
    sys.exit("python-docx is required: pip install python-docx")


def add_field(paragraph, instruction):
    """Append a Word field (e.g. PAGE, NUMPAGES) to a paragraph."""
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" {instruction} "
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.append(begin)
    run._r.append(instr)
    run._r.append(end)


def build_header_footer(document, doc_id, header_title, rev):
    section = document.sections[0]

    header = section.header
    hp = header.paragraphs[0]
    hp.text = f"{doc_id} - {header_title}"
    hp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for run in hp.runs:
        run.font.name = "Arial"
        run.font.size = Pt(9)
        run.font.bold = True
    hp2 = header.add_paragraph()
    r = hp2.add_run(f"Rev. {rev}")
    r.font.name = "Arial"
    r.font.size = Pt(9)

    footer = section.footer
    fp = footer.paragraphs[0]
    r = fp.add_run("Printed copies are uncontrolled.    Page ")
    r.font.name = "Arial"
    r.font.size = Pt(8)
    add_field(fp, "PAGE")
    r = fp.add_run(" of ")
    r.font.name = "Arial"
    r.font.size = Pt(8)
    add_field(fp, "NUMPAGES")
    for run in fp.runs:
        run.font.name = "Arial"
        run.font.size = Pt(8)


def add_revision_table(document, rows):
    table = document.add_table(rows=1, cols=4)
    table.style = "Table Grid"
    for i, label in enumerate(["Revision", "Date", "Author", "Description of Change"]):
        cell = table.rows[0].cells[i]
        cell.text = label
        for p in cell.paragraphs:
            for run in p.runs:
                run.font.bold = True
                run.font.name = "Arial"
                run.font.size = Pt(10)
    for entry in rows:
        cells = table.add_row().cells
        for i, value in enumerate(entry):
            cells[i].text = value
            for p in cells[i].paragraphs:
                for run in p.runs:
                    run.font.name = "Arial"
                    run.font.size = Pt(10)
    return table


def heading(document, text, level=1):
    p = document.add_heading(text, level=level)
    return p


def build(path, rev, *, header_title, doc_title, revision_rows, storage_text,
          acceptance_pct, handling_text, references, extra_defects):
    document = docx.Document()
    document.core_properties.title = doc_title

    build_header_footer(document, "DES-9999", header_title, rev)

    t = document.add_paragraph(doc_title)
    t.style = document.styles["Title"]

    heading(document, "1 Purpose", 1)
    purpose = (
        "This specification defines the design requirements and acceptance criteria "
        "for the DAISY NIPT assay. Sample handling is described in Section 7."
    )
    if extra_defects:
        purpose += " Environmental monitoring is described in Section 9."
    document.add_paragraph(purpose)

    heading(document, "2 References", 1)
    for ref in references:
        document.add_paragraph(ref, style="List Bullet")

    heading(document, "3 Storage Conditions", 1)
    document.add_paragraph(storage_text)

    heading(document, "4 Materials", 1)
    heading(document, "4.1 Reagents", 2)
    document.add_paragraph("Reagent module A and reagent module B are supplied in the kit.")
    heading(document, "4.2 Consumables", 2)
    document.add_paragraph("Consumables are supplied separately and are listed in Table 1.")
    heading(document, "4.4 Equipment" if extra_defects else "4.3 Equipment", 2)
    document.add_paragraph("A thermal cycler and a sequencing instrument are required.")

    heading(document, "5 Performance Requirements", 1)
    heading(document, "5.1 Analytical Sensitivity", 2)
    document.add_paragraph(
        "The limit of detection shall be 4.0% fetal fraction or lower, determined "
        "in accordance with CLSI EP17-A2."
    )
    heading(document, "5.2 Acceptance Criteria", 2)
    document.add_paragraph(
        f"The assay call rate shall be greater than or equal to {acceptance_pct}% "
        "across the validation sample set."
    )

    heading(document, "6 Verification Activities", 1)
    document.add_paragraph(
        "Verification is performed under protocol VER-0045."
        + (" The regression sample set is TBD." if extra_defects else
           " The regression sample set is defined in VER-0045 Appendix A.")
    )

    heading(document, "7 Sample Handling", 1)
    document.add_paragraph(handling_text)

    heading(document, "8 Revision History", 1)
    add_revision_table(document, revision_rows)

    document.save(str(path))
    return path


REFERENCES_COMMON = [
    "ISO 13485:2003 Medical devices - Quality management systems",
    "ISO 14971:2019 Application of risk management to medical devices",
    "ISO 15223-1:2016 Symbols to be used with information supplied by the manufacturer",
    "21 CFR 820 Quality System Regulation",
    "CLSI EP17-A2 Evaluation of Detection Capability",
]


def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)

    rev1 = build(
        out / "DES-9999_Rev1.docx",
        rev="1",
        header_title="NIPT Assay Specification",
        doc_title="NIPT Assay Specification",
        revision_rows=[
            ["1", "2024-03-15", "J. Chen", "Initial release."],
        ],
        storage_text="Store the kit at 15-25 degrees C until the expiration date printed on the label.",
        acceptance_pct="95",
        handling_text=(
            "Plasma samples shall be processed within 72 hours of collection. "
            "Store samples at room temperature (15-25 degrees C) until use."
        ),
        references=REFERENCES_COMMON,
        extra_defects=False,
    )

    rev2 = build(
        out / "DES-9999_Rev2.docx",
        rev="2",
        # Defect 1: header title differs from the document title.
        header_title="NIPT Assay Specification",
        doc_title="NIPT Assay Validation Specification",
        revision_rows=[
            ["1", "2024-03-15", "J. Chen", "Initial release."],
            # Defect 2: revision column reads 3 while the header reads Rev. 2.
            # Defect 3: claims Section 9 was added; it was not.
            ["3", "2026-01-20", "M. Lopez",
             "Updated the storage condition in Section 3 to 2-8 degrees C. "
             "Added Section 9 Environmental Monitoring."],
        ],
        storage_text="Store the kit at 2-8 degrees C until the expiration date printed on the label.",
        # Defect 4: changed from 95 with no revision history entry.
        acceptance_pct="98",
        # Defect 5: conflicts with Section 3.
        handling_text=(
            "Plasma samples shall be processed within 72 hours of collection. "
            "Store samples at room temperature (15-25 degrees C) until use."
        ),
        # Defect 6: superseded editions, carried over from Rev 1.
        references=REFERENCES_COMMON,
        # Defect 7: TBD placeholder, 4.3 -> 4.4 numbering gap, broken Section 9 reference.
        extra_defects=True,
    )

    print(f"wrote {rev1}")
    print(f"wrote {rev2}")


if __name__ == "__main__":
    main()
