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
    from docx.shared import Inches, Pt
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


# --------------------------------------------------------------------------
# Second pair: shaped like a real controlled-document template
#
# Reproduces the structural traits that the first pair does not have and that
# a real Pillar protocol (TP-0016) turned out to rely on:
#   * the revision in an unlabelled header-table cell, with no "Rev." anywhere;
#   * heading numbers generated by the Heading styles, absent from the text;
#   * a Rev | Description | Date | Name revision history with a minor-uprev
#     tail, so the header major (8) and the latest entry (8.04) differ legally;
#   * a self-labelled APPENDIX heading whose content must stay unnumbered;
#   * a footer holding only PAGE / NUMPAGES fields;
#   * numbered steps typed into body paragraphs, which must not be mistaken
#     for sections.
# No content is taken from any real document.
# --------------------------------------------------------------------------

HEADING_NUM_ID = 90  # above the default template's existing numIds (1-9)


def add_heading_numbering(document):
    """Link Heading 1-3 to a multilevel list, as a controlled template does."""
    numbering = document.part.numbering_part.element

    abstract = OxmlElement("w:abstractNum")
    abstract.set(qn("w:abstractNumId"), str(HEADING_NUM_ID))
    multi = OxmlElement("w:multiLevelType")
    multi.set(qn("w:val"), "multilevel")
    abstract.append(multi)
    for ilvl in range(3):
        lvl = OxmlElement("w:lvl")
        lvl.set(qn("w:ilvl"), str(ilvl))
        start = OxmlElement("w:start")
        start.set(qn("w:val"), "1")
        fmt = OxmlElement("w:numFmt")
        fmt.set(qn("w:val"), "decimal")
        text = OxmlElement("w:lvlText")
        text.set(qn("w:val"), ".".join(f"%{i + 1}" for i in range(ilvl + 1)))
        just = OxmlElement("w:lvlJc")
        just.set(qn("w:val"), "left")
        for el in (start, fmt, text, just):
            lvl.append(el)
        abstract.append(lvl)
    numbering.append(abstract)

    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(HEADING_NUM_ID))
    aid = OxmlElement("w:abstractNumId")
    aid.set(qn("w:val"), str(HEADING_NUM_ID))
    num.append(aid)
    numbering.append(num)

    # Point the Heading styles at it, which is what keeps the numbers out of
    # the paragraph text entirely.
    for level in (1, 2, 3):
        style = document.styles[f"Heading {level}"]
        ppr = style.element.get_or_add_pPr()
        numpr = OxmlElement("w:numPr")
        ilvl = OxmlElement("w:ilvl")
        ilvl.set(qn("w:val"), str(level - 1))
        nid = OxmlElement("w:numId")
        nid.set(qn("w:val"), str(HEADING_NUM_ID))
        numpr.append(ilvl)
        numpr.append(nid)
        ppr.append(numpr)


def build_template_header_footer(document, doc_id, title, rev):
    """Header as a table: title row, then a bare revision cell with no label."""
    section = document.sections[0]

    header = section.header
    table = header.add_table(rows=2, cols=2, width=Inches(6.5))
    table.rows[0].cells[0].text = ""
    table.rows[0].cells[1].text = f"{doc_id} – {title},"
    table.rows[1].cells[0].text = ""
    table.rows[1].cells[1].text = str(rev)  # no "Rev." label anywhere

    footer = section.footer
    fp = footer.paragraphs[0]
    add_field(fp, "PAGE")          # fields only: no controlled-copy statement,
    add_field(fp, "NUMPAGES")      # no template ID, no readable text


def build_template_shaped(path, rev, *, revision_rows, hold_minutes, extra_step):
    document = docx.Document()
    add_heading_numbering(document)
    build_template_header_footer(document, "TP-9999", "Widget Qualification", rev)

    document.add_heading("PURPOSE AND SCOPE", 1)
    document.add_paragraph(
        "This protocol qualifies the widget. Sample handling is covered in Section 4."
    )
    document.add_heading("REFERENCES", 1)
    document.add_paragraph("FORM-0001 Qualification Report", style="List Bullet")
    document.add_heading("DEFINITIONS AND ACRONYMS", 1)
    deftable = document.add_table(rows=1, cols=2)
    deftable.style = "Table Grid"
    deftable.rows[0].cells[0].text = "Term"
    deftable.rows[0].cells[1].text = "Definition"
    for term, meaning in (("WQT", "Widget Qualification Test"),
                          ("LBx", "Liquid Biopsy")):
        cells = deftable.add_row().cells
        cells[0].text = term
        cells[1].text = meaning

    document.add_heading("PROCEDURE", 1)
    document.add_heading("Incoming Inspection", 2)
    document.add_paragraph(f"Hold the widget for {hold_minutes} minutes before inspection.")
    document.add_heading("Functional Test", 2)
    document.add_paragraph("Run the WQT and record the result.")
    document.add_heading("Product Assessment Documentation", 2)
    document.add_paragraph("File the completed FORM-0001 with the batch record.")
    if extra_step:
        # A numbered step typed into a body paragraph: must NOT become a section.
        document.add_paragraph(
            "5.4.1 If acceptance criteria are not met, escalate to QA for disposition."
        )

    document.add_heading("RECORDS", 1)
    document.add_paragraph("Records are filed with the batch record.")
    document.add_heading("REVISION HISTORY", 1)

    table = document.add_table(rows=1, cols=4)
    table.style = "Table Grid"
    for i, label in enumerate(["Rev", "Description", "Date", "Name"]):
        table.rows[0].cells[i].text = label
    for entry in revision_rows:
        cells = table.add_row().cells
        for i, value in enumerate(entry):
            cells[i].text = value
            for p in cells[i].paragraphs:
                for run in p.runs:
                    run.font.name = "Calibri"
                    run.font.size = Pt(9)

    # Self-labelled appendix: unnumbered, and its steps must stay unnumbered too.
    document.add_heading("APPENDIX A – INSTRUMENT SETUP", 1)
    document.add_heading("Switch the instrument on.", 2)
    document.add_heading("Load the test fixture.", 2)

    document.save(str(path))
    return path


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

    # ---- second pair: shaped like a real controlled-document template ----
    tpl_rows_old = [
        ["7", "Initial release.", "2025-01-10", "A. Reviewer"],
    ]
    tpl_rows_new = [
        ["7", "Initial release.", "2025-01-10", "A. Reviewer"],
        ["8", "Section 2 References - added FORM-0001. "
              "Section 5.3 Product Assessment Documentation - added the filing step.",
         "2026-02-15", "A. Reviewer"],
        ["8.01", "Correct Typo: fixed a spelling error in Section 1 Purpose and Scope.",
         "2026-03-05", "B. Checker"],
        ["8.04", "Minor Revision: extended the incoming inspection hold time.",
         "2026-06-15", "B. Checker"],
    ]
    tpl_old = build_template_shaped(
        out / "TP-9999_Rev7.docx", rev="7",
        revision_rows=tpl_rows_old, hold_minutes="15", extra_step=False,
    )
    tpl_new = build_template_shaped(
        out / "TP-9999_Rev8.docx", rev="8",
        revision_rows=tpl_rows_new, hold_minutes="30", extra_step=True,
    )

    print(f"wrote {rev1}")
    print(f"wrote {rev2}")
    print(f"wrote {tpl_old}")
    print(f"wrote {tpl_new}")


if __name__ == "__main__":
    main()
