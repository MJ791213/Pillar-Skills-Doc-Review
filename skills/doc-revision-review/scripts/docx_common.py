#!/usr/bin/env python3
"""Shared .docx traversal helpers for the doc-revision-review scripts.

Deliberately mirrors the helpers proven in the label-spec-review skill
(iter_block_items / table_rows / text_of / para_level) so both skills parse
Word documents the same way.
"""
import re

try:
    from docx.oxml.ns import qn
    from docx.table import Table
    from docx.text.paragraph import Paragraph
except ImportError:  # pragma: no cover
    import sys

    sys.exit("python-docx is required: pip install python-docx")


# --------------------------------------------------------------------------
# Patterns shared across the skill
# --------------------------------------------------------------------------

#: Pillar document identifiers: DES-0028, UM-0106, RSK-0017, LB-0669, SRS-219 ...
DOC_ID_RE = re.compile(r"\b([A-Z]{2,5})-(\d{3,6})\b")

#: "Rev. 2", "Rev 2", "Revision B", "Version 1.0"
REV_RE = re.compile(
    r"\b(?:Rev(?:ision)?|Ver(?:sion)?)\.?\s*[:\-]?\s*([0-9]+(?:\.[0-9]+)?[A-Z]?|[A-Z])\b",
    re.I,
)

#: A numbered heading: "5.2 Acceptance Criteria"
HEADING_NUM_RE = re.compile(r"^(\d+(?:\.\d+)*)\.?\s+(\S.*)$")

#: Lines that start with a number but are citations, dates or quantities,
#: not headings. Checked before the numbered-heading guess.
NOT_A_HEADING_RE = re.compile(
    r"^\d+\s*(?:CFR|USC|%|degrees?|deg\b|°|mL|uL|µL|mg|ng|mm|cm|hours?|hrs?|days?|"
    r"minutes?|min\b|months?|years?|x\b)",
    re.I,
)


def revision_sort_key(rev):
    """Order revisions so the newest sorts highest.

    Handles numeric ("1", "2", "10", "1.1") and alphabetic ("A", "B") schemes.
    Returns a tuple so numeric and alphabetic revisions never compare directly.
    """
    if rev is None:
        return (-1, ())
    r = str(rev).strip().upper()
    if not r:
        return (-1, ())
    m = re.fullmatch(r"(\d+(?:\.\d+)*)([A-Z]?)", r)
    if m:
        parts = tuple(int(p) for p in m.group(1).split("."))
        suffix = ord(m.group(2)) - 64 if m.group(2) else 0
        return (1, parts + (suffix,))
    if re.fullmatch(r"[A-Z]+", r):
        value = 0
        for ch in r:
            value = value * 26 + (ord(ch) - 64)
        return (0, (value,))
    return (-1, (r,))


# --------------------------------------------------------------------------
# Traversal
# --------------------------------------------------------------------------

def iter_block_items(parent):
    """Yield paragraphs and tables of a body / cell / header in document order.

    ``document.paragraphs`` does not reach header or footer parts, and it
    silently drops tables; this walks the XML children instead.
    """
    elm = parent.element.body if hasattr(parent, "element") and hasattr(parent.element, "body") else parent._element
    for child in elm.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, parent)
        elif child.tag == qn("w:tbl"):
            yield Table(child, parent)


def para_level(p):
    """List nesting level of a paragraph (0 = top level, 1 = sub-item ...)."""
    ppr = p._p.pPr
    if ppr is not None and ppr.numPr is not None and ppr.numPr.ilvl is not None:
        return int(ppr.numPr.ilvl.val)
    indent = p.paragraph_format.left_indent
    if indent is not None and indent.pt > 54:
        return 1
    return 0


def table_rows(tbl):
    """Rows of a table as lists of cells, with merged cells de-duplicated."""
    rows = []
    for row in tbl.rows:
        cells, seen = [], set()
        for c in row.cells:  # merged cells repeat; keep the first occurrence
            if id(c._tc) in seen:
                continue
            seen.add(id(c._tc))
            cells.append(c)
        rows.append(cells)
    return rows


def table_text_rows(tbl):
    """Rows of a table as lists of plain strings."""
    return [[c.text.strip() for c in row] for row in table_rows(tbl)]


def text_of(container):
    """All text of a body / header / footer, tables included, in order."""
    parts = []
    for block in iter_block_items(container):
        if isinstance(block, Paragraph):
            parts.append(block.text)
        else:
            for row in table_text_rows(block):
                parts.append(" | ".join(row))
    return "\n".join(t for t in parts if t.strip())


def heading_level(p):
    """Outline level of a paragraph, or None if it is not a heading.

    Uses the style name first, then an explicit outline level, then a
    numbered-heading text pattern, so documents built from a template that
    does not use Word's Heading styles still produce a section tree.
    """
    style = (p.style.name if p.style is not None else "") or ""
    m = re.match(r"^Heading\s+(\d+)$", style, re.I)
    if m:
        return int(m.group(1))
    if style.strip().lower() == "title":
        return 0
    ppr = p._p.pPr
    if ppr is not None:
        outline = ppr.find(qn("w:outlineLvl"))
        if outline is not None:
            val = outline.get(qn("w:val"))
            if val is not None and val.isdigit() and int(val) < 9:
                return int(val) + 1

    # Below here we are guessing from the text alone. A bulleted or numbered
    # list item is never a heading, however much it looks like one - a
    # reference list entry such as "21 CFR 820 Quality System Regulation"
    # otherwise becomes "Section 21".
    if ppr is not None and ppr.numPr is not None:
        return None
    if "list" in style.strip().lower():
        return None

    text = p.text.strip()
    if NOT_A_HEADING_RE.match(text):
        return None
    m = HEADING_NUM_RE.match(text)
    if m and len(text) < 120:
        return m.group(1).count(".") + 1
    return None


def normalise(text):
    """Collapse whitespace and unify the dashes/quotes Word likes to swap."""
    if not text:
        return ""
    t = str(text)
    for a, b in (("–", "-"), ("—", "-"), ("−", "-"),
                 ("‘", "'"), ("’", "'"),
                 ("“", '"'), ("”", '"'),
                 (" ", " ")):
        t = t.replace(a, b)
    return re.sub(r"\s+", " ", t).strip()
