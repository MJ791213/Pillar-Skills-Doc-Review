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


def major_revision(rev):
    """The major component of a revision: "8.04" -> "8", "B" -> "B"."""
    if rev is None:
        return None
    r = str(rev).strip().upper()
    m = re.match(r"^(\d+)", r)
    return m.group(1) if m else r


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


def heading_level(p, allow_text_fallback=True):
    """Outline level of a paragraph, or None if it is not a heading.

    Uses the style name first, then an explicit outline level, then - only when
    ``allow_text_fallback`` - a numbered-heading text pattern, so documents
    built from a template that does not use Word's Heading styles still produce
    a section tree.

    Pass ``allow_text_fallback=False`` for a document whose Heading styles
    already carry the numbering. There, a body paragraph beginning with a
    dotted number is a numbered *step* ("5.6.6 If acceptance criteria are not
    met ..."), not a section, and promoting it would invent sections and make
    ordinary edits look like structural changes.
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
    if not allow_text_fallback:
        return None
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


# --------------------------------------------------------------------------
# Heading numbers
#
# Most controlled templates number their headings through the Heading styles'
# linked multilevel list, which means the number Word displays appears nowhere
# in the file's text and nowhere on the paragraph. It has to be rebuilt from
# styles.xml and numbering.xml the way Word builds it, or every finding that
# cites a section number is guesswork.
# --------------------------------------------------------------------------

#: Headings that carry their own label and are not part of the numbered run.
SELF_LABELLED_HEADING_RE = re.compile(r"^\s*(APPENDIX|ANNEX|ATTACHMENT|EXHIBIT)\b", re.I)

_ROMAN = [
    (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"),
    (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I"),
]


def _roman(n):
    out = []
    for value, letter in _ROMAN:
        while n >= value:
            out.append(letter)
            n -= value
    return "".join(out)


def _letter(n):
    out = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out


def format_counter(fmt, n):
    """Render one counter in a Word numFmt."""
    fmt = (fmt or "decimal").lower()
    if fmt == "upperroman":
        return _roman(n)
    if fmt == "lowerroman":
        return _roman(n).lower()
    if fmt == "upperletter":
        return _letter(n)
    if fmt == "lowerletter":
        return _letter(n).lower()
    if fmt == "decimalzero":
        return f"{n:02d}"
    if fmt == "none":
        return ""
    return str(n)


def heading_style_numbering(document):
    """Heading outline level -> (numId, ilvl), read from the Heading styles."""
    out = {}
    try:
        styles_root = document.styles.element
    except Exception:
        return out
    for st in styles_root.findall(qn("w:style")):
        name_el = st.find(qn("w:name"))
        if name_el is None:
            continue
        name = name_el.get(qn("w:val")) or ""
        m = re.match(r"^heading\s*(\d)$", name.strip(), re.I)
        if not m:
            continue
        level = int(m.group(1))
        ppr = st.find(qn("w:pPr"))
        if ppr is None:
            continue
        numpr = ppr.find(qn("w:numPr"))
        if numpr is None:
            continue
        num_id = numpr.find(qn("w:numId"))
        if num_id is None or num_id.get(qn("w:val")) is None:
            continue
        ilvl_el = numpr.find(qn("w:ilvl"))
        ilvl = int(ilvl_el.get(qn("w:val"))) if ilvl_el is not None else level - 1
        out[level] = (int(num_id.get(qn("w:val"))), ilvl)
    return out


def numbering_definitions(document):
    """numId -> {ilvl: {start, numFmt, lvlText}}, following abstractNum and overrides."""
    try:
        root = document.part.numbering_part.element
    except Exception:
        return {}

    def read_levels(parent):
        levels = {}
        for lvl in parent.findall(qn("w:lvl")):
            if lvl.get(qn("w:ilvl")) is None:
                continue
            ilvl = int(lvl.get(qn("w:ilvl")))
            start = lvl.find(qn("w:start"))
            fmt = lvl.find(qn("w:numFmt"))
            text = lvl.find(qn("w:lvlText"))
            levels[ilvl] = {
                "start": int(start.get(qn("w:val"))) if start is not None and start.get(qn("w:val")) else 1,
                "numFmt": fmt.get(qn("w:val")) if fmt is not None else "decimal",
                "lvlText": text.get(qn("w:val")) if text is not None else "%1.",
            }
        return levels

    abstract = {}
    for an in root.findall(qn("w:abstractNum")):
        if an.get(qn("w:abstractNumId")) is None:
            continue
        abstract[int(an.get(qn("w:abstractNumId")))] = read_levels(an)

    nums = {}
    for num in root.findall(qn("w:num")):
        if num.get(qn("w:numId")) is None:
            continue
        nid = int(num.get(qn("w:numId")))
        ai = num.find(qn("w:abstractNumId"))
        if ai is None or ai.get(qn("w:val")) is None:
            continue
        levels = {k: dict(v) for k, v in abstract.get(int(ai.get(qn("w:val"))), {}).items()}
        for ov in num.findall(qn("w:lvlOverride")):
            if ov.get(qn("w:ilvl")) is None:
                continue
            ilvl = int(ov.get(qn("w:ilvl")))
            level = levels.setdefault(ilvl, {"start": 1, "numFmt": "decimal", "lvlText": "%1."})
            so = ov.find(qn("w:startOverride"))
            if so is not None and so.get(qn("w:val")):
                level["start"] = int(so.get(qn("w:val")))
            inner = ov.find(qn("w:lvl"))
            if inner is not None:
                sub = read_levels(ov).get(ilvl)
                if sub:
                    level.update(sub)
        nums[nid] = levels
    return nums


class HeadingNumberer:
    """Rebuilds the heading number Word would display, one heading at a time.

    Call `number_for(paragraph, level)` for each heading **in document order**;
    it advances the counters and returns the number, or None for a heading that
    Word leaves unnumbered.

    Deliberately stateful rather than returning a paragraph -> number mapping:
    lxml creates a fresh Python proxy for the same underlying node on each
    traversal, so `id(paragraph._p)` is not stable between two walks of the
    document and using it as a key silently attaches numbers to the wrong
    paragraphs.

    Left out of the numbered sequence, because Word leaves them out too:
      * empty headings;
      * heading-styled paragraphs carrying their *own* numPr, which are list
        items that borrowed a Heading style rather than real headings;
      * headings that label themselves ("APPENDIX A - ..."), which templates
        keep outside the numbered run.
    """

    def __init__(self, document):
        style_num = heading_style_numbering(document)
        defs = numbering_definitions(document)
        self.level_defs = {}
        for level, (num_id, ilvl) in style_num.items():
            spec = defs.get(num_id, {}).get(ilvl)
            if spec:
                self.level_defs[level] = (ilvl, spec)
        self.source = "resolved" if self.level_defs else "inferred"
        self.counters = {}
        self._spec_by_ilvl = {li: sp for (li, sp) in self.level_defs.values()}
        # Content under an unnumbered heading is unnumbered too: steps inside
        # "APPENDIX C" must not come out as 7.1, 7.2 ... and collide with the
        # real section 7.
        self._suppress_below = None

    def _start(self, ilvl):
        spec = self._spec_by_ilvl.get(ilvl)
        return spec.get("start", 1) if spec else 1

    def number_for(self, paragraph, level):
        if level is None or level == 0:
            return None
        text = (paragraph.text or "").strip()
        if not text:
            return None
        ppr = paragraph._p.pPr
        if ppr is not None and ppr.numPr is not None:
            return None

        ilvl = self.level_defs.get(level, (level - 1, None))[0]

        if SELF_LABELLED_HEADING_RE.match(text):
            # Unnumbered itself, and everything nested under it is unnumbered.
            self._suppress_below = ilvl
            return None
        if self._suppress_below is not None:
            if ilvl > self._suppress_below:
                return None
            self._suppress_below = None

        self.counters[ilvl] = self.counters.get(ilvl, self._start(ilvl) - 1) + 1
        for deeper in [k for k in self.counters if k > ilvl]:
            del self.counters[deeper]

        parts = []
        for i in range(ilvl + 1):
            spec = self._spec_by_ilvl.get(i)
            value = self.counters.get(i, self._start(i))
            parts.append(format_counter(spec.get("numFmt") if spec else "decimal", value))

        spec = self.level_defs.get(level, (None, None))[1]
        lvl_text = spec.get("lvlText") if spec else None
        if lvl_text:
            rendered = lvl_text
            for i, value in enumerate(parts, start=1):
                rendered = rendered.replace(f"%{i}", value)
            rendered = re.sub(r"%\d", "", rendered)
        else:
            rendered = ".".join(parts)
        rendered = rendered.strip().strip(".").strip()
        return rendered or None


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
