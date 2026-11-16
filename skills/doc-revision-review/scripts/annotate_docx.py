#!/usr/bin/env python3
"""Write review findings into a .docx as a reviewer would mark up a draft.

Usage:
    python3 annotate_docx.py <new.docx> --findings findings.json --out <annotated.docx>
    python3 annotate_docx.py <annotated.docx> --strip-annotations --out <clean.docx>

Red  - notes written into the revision history table, next to the entry they
       concern, in the font already used in that cell.
Appended at the end of the document, a "Claude Review" block holding two
tables:
  * Errors (red)          - Section | Error.          What is wrong.
  * Recommendations (orange) - Section | Recommendation. What could be better.

The input file is never modified; a new file is always written.

findings.json schema
--------------------
{
  "document":  {"doc_id": "DES-9999", "revision": "2",
                "title": "NIPT Assay Specification"},
  "revision_history_table_index": 0,
  "revision_history_notes": [
      {"row_index": 2, "col_index": 3, "text": "Undocumented change: ..."}
  ],
  "errors": [
      {"section": "§5.2", "error": "...",
       "severity": "Major", "qa_q": "QA-Q6"}
  ],
  "recommendations": [
      {"section": "Page header", "recommendation": "..."}
  ],
  "checks_not_performed": ["Check 5 - internal consistency: not completed"]
}

"severity" and "qa_q" are optional. They are not printed - the tables have
exactly the two columns above - but severity orders the Errors table, putting
Major first, and both are kept in findings.json for the change order.

Every key except "document" is optional; an empty list simply writes nothing.
The older "summary" / "location" / "finding" / "suggested_action" keys are
still read so an existing findings file does not break.
"""
import argparse
import copy
import json
import sys
from datetime import date
from pathlib import Path

try:
    import docx
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.shared import Pt, RGBColor
except ImportError:  # pragma: no cover
    sys.exit("python-docx is required: pip install python-docx")

RED = RGBColor(0xFF, 0x00, 0x00)
ORANGE = RGBColor(0xED, 0x7D, 0x31)  # Word standard "Orange, Accent 2"

MARKER = "[Claude]"
BLOCK_HEADING = "Claude Review — not part of the approved document"
#: Older annotated files used this heading; still recognised when stripping.
LEGACY_BLOCK_HEADING = "Claude Recommendations — not part of the approved document"

SEVERITY_ORDER = {"major": 0, "minor": 1, "observation": 2}


# --------------------------------------------------------------------------
# Font-preserving run insertion
# --------------------------------------------------------------------------

def last_run_with_text(*containers):
    """The last run that actually carries text, searched across containers."""
    for container in containers:
        if container is None:
            continue
        paragraphs = getattr(container, "paragraphs", None)
        if paragraphs is None:
            paragraphs = [container]
        for p in reversed(list(paragraphs)):
            for r in reversed(list(p.runs)):
                if r.text and r.text.strip():
                    return r
    return None


def styled_run(paragraph, text, rgb, template_run=None):
    """Append a coloured run that keeps the template run's character format.

    Setting ``run.font.name`` is not enough: the effective font often comes
    from the paragraph style or from the east-asian / complex-script
    attributes of ``w:rFonts``, none of which that touches. Copying the whole
    ``w:rPr`` carries family, size, bold, italic and style across, and the
    colour is then overridden on the copy.
    """
    run = paragraph.add_run(text)
    if template_run is not None:
        src = template_run._element.find(qn("w:rPr"))
        if src is not None:
            clone = copy.deepcopy(src)
            # Drop formatting we are about to set ourselves, plus anything
            # that would make the note look like original content.
            for tag in ("w:color", "w:highlight", "w:u", "w:strike", "w:dstrike"):
                for el in clone.findall(qn(tag)):
                    clone.remove(el)
            existing = run._element.find(qn("w:rPr"))
            if existing is not None:
                run._element.remove(existing)
            run._element.insert(0, clone)
    run.font.color.rgb = rgb
    return run


def copy_paragraph_format(source_p, target_p):
    """Carry indentation / spacing / style from one paragraph to another."""
    if source_p is None:
        return
    try:
        target_p.style = source_p.style
    except Exception:
        pass
    src_ppr = source_p._p.find(qn("w:pPr"))
    if src_ppr is None:
        return
    clone = copy.deepcopy(src_ppr)
    # Numbering would restart a list; drop it.
    for el in clone.findall(qn("w:numPr")):
        clone.remove(el)
    existing = target_p._p.find(qn("w:pPr"))
    if existing is not None:
        target_p._p.remove(existing)
    target_p._p.insert(0, clone)


# --------------------------------------------------------------------------
# Red notes in the revision history table
# --------------------------------------------------------------------------

def annotate_revision_history(document, findings):
    notes = findings.get("revision_history_notes") or []
    if not notes:
        return 0, []

    t_index = findings.get("revision_history_table_index")
    if t_index is None:
        raise SystemExit(
            "findings.json has revision_history_notes but no "
            "revision_history_table_index - refusing to guess which table to annotate."
        )
    tables = document.tables
    if t_index >= len(tables):
        raise SystemExit(
            f"revision_history_table_index {t_index} is out of range: "
            f"the document has {len(tables)} tables."
        )
    table = tables[t_index]

    written, problems = 0, []
    for note in notes:
        row_index = note.get("row_index")
        text = (note.get("text") or "").strip()
        if row_index is None or not text:
            problems.append(f"skipped a note without row_index or text: {note!r}")
            continue
        if row_index >= len(table.rows):
            problems.append(
                f"row_index {row_index} is out of range (table has {len(table.rows)} rows); "
                "note not written"
            )
            continue
        row = table.rows[row_index]
        cells = list(row.cells)
        col_index = note.get("col_index")
        if col_index is None or col_index >= len(cells):
            col_index = len(cells) - 1  # the description column, by convention
        cell = cells[col_index]

        source_p = cell.paragraphs[-1] if cell.paragraphs else None
        template = last_run_with_text(cell, table.rows[min(1, len(table.rows) - 1)].cells[col_index])

        target = cell.add_paragraph()
        copy_paragraph_format(source_p, target)
        styled_run(target, f"{MARKER} {text}", RED, template)
        written += 1
    return written, problems


# --------------------------------------------------------------------------
# Orange recommendations block
# --------------------------------------------------------------------------

def add_coloured_table(document, headers, rows, template_run, rgb):
    table = document.add_table(rows=1, cols=len(headers))
    try:
        table.style = "Table Grid"
    except KeyError:  # pragma: no cover - template without that style
        pass
    for i, label in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ""
        run = styled_run(cell.paragraphs[0], label, rgb, template_run)
        run.font.bold = True
    for row_values in rows:
        cells = table.add_row().cells
        for i, value in enumerate(row_values):
            cells[i].text = ""
            styled_run(cells[i].paragraphs[0], str(value or ""), rgb, template_run)
    return table


def _section_of(item):
    """The section an item belongs to, accepting the older 'location' key."""
    return item.get("section") or item.get("location") or ""


def _error_text(item):
    text = item.get("error") or item.get("finding") or ""
    return str(text).strip()


def _recommendation_text(item):
    text = (
        item.get("recommendation")
        or item.get("suggested_action")
        or item.get("finding")
        or ""
    )
    return str(text).strip()


def append_review_block(document, findings, source_name):
    """Append the review as two tables: errors, then recommendations.

    Errors are defects - something in the document is wrong and needs
    correcting. Recommendations are improvements - the document is not wrong,
    but it could be clearer or better presented. Keeping them in separate
    tables lets a reviewer deal with the must-fix list without reading past
    the nice-to-have list.
    """
    errors = findings.get("errors") or []
    recs = findings.get("recommendations") or []
    not_performed = findings.get("checks_not_performed") or []
    if not errors and not recs and not not_performed:
        return False

    template_run = last_run_with_text(document)
    doc_meta = findings.get("document") or {}

    document.add_paragraph()
    heading = document.add_paragraph()
    run = styled_run(heading, BLOCK_HEADING, ORANGE, template_run)
    run.font.bold = True
    run.font.size = Pt(13)

    provenance = document.add_paragraph()
    styled_run(
        provenance,
        "Generated by Claude on {d} from {f}{rev}. Everything in this block is a "
        "Claude recommendation for the reviewer to consider. It is not approved "
        "content, it is not part of the controlled document, and it must be removed "
        "before the document is released.".format(
            d=date.today().isoformat(),
            f=source_name,
            rev=(
                f" ({doc_meta.get('doc_id')} Rev {doc_meta.get('revision')})"
                if doc_meta.get("doc_id") else ""
            ),
        ),
        ORANGE,
        template_run,
    )

    # ---- Table 1: errors, in red ----------------------------------------
    p = document.add_paragraph()
    run = styled_run(p, "Errors", RED, template_run)
    run.font.bold = True
    if errors:
        ordered = sorted(
            errors,
            key=lambda e: SEVERITY_ORDER.get(str(e.get("severity", "")).lower(), 3),
        )
        add_coloured_table(
            document,
            ["Section", "Error"],
            [[_section_of(e), _error_text(e)] for e in ordered],
            template_run,
            RED,
        )
    else:
        styled_run(
            document.add_paragraph(),
            "No errors identified by the checks that were performed.",
            RED,
            template_run,
        )

    # ---- Table 2: recommendations, in orange -----------------------------
    document.add_paragraph()
    p = document.add_paragraph()
    run = styled_run(p, "Recommendations", ORANGE, template_run)
    run.font.bold = True
    if recs:
        add_coloured_table(
            document,
            ["Section", "Recommendation"],
            [[_section_of(r), _recommendation_text(r)] for r in recs],
            template_run,
            ORANGE,
        )
    else:
        styled_run(
            document.add_paragraph(), "None.", ORANGE, template_run
        )

    # Checks that did not run must stay visible, or a reviewer reads an empty
    # Errors table as "nothing wrong" when it means "not looked at".
    if not_performed:
        document.add_paragraph()
        p = document.add_paragraph()
        run = styled_run(p, "Not checked", ORANGE, template_run)
        run.font.bold = True
        for item in not_performed:
            styled_run(document.add_paragraph(), f"• {item}", ORANGE, template_run)

    closing = document.add_paragraph()
    closing.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = styled_run(
        closing,
        "Note: the above is a Claude recommendation. It has not been reviewed or "
        "approved by QA. Confirm every regulatory citation against the controlled "
        "standards list, and confirm every finding against the source documents, "
        "before acting on it.",
        ORANGE,
        template_run,
    )
    run.font.italic = True
    return True


# --------------------------------------------------------------------------
# Strip
# --------------------------------------------------------------------------

def strip_annotations(document):
    """Remove a previous round of annotations so a document can be re-reviewed."""
    removed_runs = 0
    body = document.element.body

    for r in body.findall(f".//{qn('w:r')}"):
        text = "".join(t.text or "" for t in r.findall(qn("w:t")))
        if text.strip().startswith(MARKER):
            parent = r.getparent()
            if parent is not None:
                parent.remove(r)
                removed_runs += 1

    # Delete empty paragraphs left behind inside table cells.
    removed_paragraphs = 0
    for p in body.findall(f".//{qn('w:p')}"):
        if p.findall(qn("w:r")):
            continue
        parent = p.getparent()
        if parent is not None and parent.tag == qn("w:tc") and len(parent.findall(qn("w:p"))) > 1:
            parent.remove(p)
            removed_paragraphs += 1

    # The recommendations block is always appended last: find its heading and
    # drop it and everything after it.
    removed_block = False
    children = list(body.iterchildren())
    start = None
    prefixes = (BLOCK_HEADING[:20], LEGACY_BLOCK_HEADING[:20])
    for i, child in enumerate(children):
        if child.tag != qn("w:p"):
            continue
        text = "".join(t.text or "" for t in child.findall(f".//{qn('w:t')}"))
        if text.strip().startswith(prefixes):
            start = i
            break
    if start is not None:
        for child in children[start:]:
            if child.tag == qn("w:sectPr"):
                continue
            body.remove(child)
        removed_block = True

    return removed_runs, removed_paragraphs, removed_block


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("document", help="the .docx to annotate (never modified)")
    ap.add_argument("--findings", help="path to findings.json")
    ap.add_argument("--out", required=True, help="path for the annotated copy")
    ap.add_argument(
        "--strip-annotations",
        action="store_true",
        help="remove previous Claude annotations instead of adding any",
    )
    a = ap.parse_args()

    src = Path(a.document)
    out = Path(a.out)
    if not src.exists():
        sys.exit(f"No such file: {src}")
    if src.resolve() == out.resolve():
        sys.exit("--out must differ from the input file; the source is never modified.")
    if src.suffix.lower() != ".docx":
        sys.exit(
            f"Only .docx can be annotated, not '{src.suffix}'. For a PDF, build a "
            "separate .docx review record instead."
        )
    out.parent.mkdir(parents=True, exist_ok=True)

    document = docx.Document(str(src))

    if a.strip_annotations:
        runs, paras, block = strip_annotations(document)
        document.save(str(out))
        print(f"removed {runs} annotation run(s), {paras} empty paragraph(s)")
        print(f"removed recommendations block: {block}")
        print(f"wrote {out}")
        return

    if not a.findings:
        sys.exit("--findings is required unless --strip-annotations is given.")
    findings_path = Path(a.findings)
    if not findings_path.exists():
        sys.exit(f"No such findings file: {findings_path}")
    try:
        findings = json.loads(findings_path.read_text())
    except json.JSONDecodeError as exc:
        sys.exit(f"findings.json is not valid JSON: {exc}")

    written, problems = annotate_revision_history(document, findings)
    block = append_review_block(document, findings, src.name)
    document.save(str(out))

    print(f"red notes in revision history : {written}")
    print(f"errors table                   : {len(findings.get('errors') or [])} row(s)")
    print(f"recommendations table          : {len(findings.get('recommendations') or [])} row(s)")
    print(f"review block appended          : {'yes' if block else 'no'}")
    for p in problems:
        print(f"[WARN  ] {p}")
    print(f"wrote {out}")
    if problems:
        sys.exit(1)


if __name__ == "__main__":
    main()
