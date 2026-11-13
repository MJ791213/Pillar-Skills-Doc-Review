#!/usr/bin/env python3
"""Extract the reviewable structure of a controlled document.

Usage:
    python3 extract_doc.py <document.docx|.pdf> --out <dir>

Writes <dir>/extract.json containing the page header and footer (read from the
real w:hdr/w:ftr parts, which document.paragraphs never reaches), the document
title, the revision history table, a numbered section tree, every standard /
regulation cited, and mechanical hygiene checks.

Everything this script reports is text it actually found. Judgement calls are
left to the reviewer.
"""
import argparse
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from docx_common import (  # noqa: E402
    DOC_ID_RE,
    HEADING_NUM_RE,
    NOT_A_HEADING_RE,
    REV_RE,
    heading_level,
    iter_block_items,
    normalise,
    revision_sort_key,
    table_text_rows,
    text_of,
)

try:
    import docx
    from docx.oxml.ns import qn
    from docx.text.paragraph import Paragraph
except ImportError:  # pragma: no cover
    sys.exit("python-docx is required: pip install python-docx")


# --------------------------------------------------------------------------
# Revision history table
# --------------------------------------------------------------------------

REV_COL = re.compile(r"\b(rev(ision)?|version|issue)\b", re.I)
DATE_COL = re.compile(r"\b(date|effective|issued|released)\b", re.I)
AUTHOR_COL = re.compile(r"\b(author|by|owner|prepared|originator|revised by)\b", re.I)
DESC_COL = re.compile(
    r"\b(description|change|summary|reason|details?|nature|modification|revision history)\b", re.I
)
DCR_COL = re.compile(r"\b(dcr|change order|co\b|eco|ecn|reference)\b", re.I)


def find_revision_history(tables):
    """Locate the revision history table and map its columns.

    Returns None when no table looks like a revision history; the caller must
    then stop rather than annotate the wrong table.
    """
    best = None
    for t_index, tbl in enumerate(tables):
        rows = table_text_rows(tbl)
        if len(rows) < 2:
            continue
        # The header row is usually row 0, but some templates put a title row
        # above it. Try the first three rows.
        for h in range(min(3, len(rows) - 1)):
            head = [normalise(c).lower() for c in rows[h]]
            if not any(REV_COL.search(c) for c in head):
                continue
            if not any(DESC_COL.search(c) for c in head):
                continue
            cols = {}
            for i, c in enumerate(head):
                if "rev" not in cols and REV_COL.search(c):
                    cols["rev"] = i
                elif "date" not in cols and DATE_COL.search(c):
                    cols["date"] = i
                elif "author" not in cols and AUTHOR_COL.search(c):
                    cols["author"] = i
                elif "dcr" not in cols and DCR_COL.search(c):
                    cols["dcr"] = i
                elif "description" not in cols and DESC_COL.search(c):
                    cols["description"] = i
            if "rev" not in cols or "description" not in cols:
                continue
            score = len(cols) * 100 + len(rows)
            if best is None or score > best["score"]:
                best = {
                    "score": score,
                    "table_index": t_index,
                    "header_row_index": h,
                    "columns": cols,
                    "header_row": rows[h],
                    "rows": rows,
                }
            break
    if best is None:
        return None

    entries = []
    for r in range(best["header_row_index"] + 1, len(best["rows"])):
        cells = best["rows"][r]
        if not any(normalise(c) for c in cells):
            continue

        def get(key):
            i = best["columns"].get(key)
            return normalise(cells[i]) if i is not None and i < len(cells) else ""

        rev_raw = get("rev")
        m = REV_RE.search(rev_raw) if rev_raw else None
        entries.append(
            {
                "row_index": r,
                "rev_raw": rev_raw,
                "rev": (m.group(1) if m else rev_raw).upper() or None,
                "date": get("date"),
                "author": get("author"),
                "dcr": get("dcr"),
                "description": get("description"),
                "cells": cells,
            }
        )

    latest = None
    if entries:
        latest = max(entries, key=lambda e: revision_sort_key(e["rev"]))
    return {
        "table_index": best["table_index"],
        "header_row_index": best["header_row_index"],
        "header_row": best["header_row"],
        "columns": best["columns"],
        "entries": entries,
        "latest_entry_row_index": latest["row_index"] if latest else None,
        "latest_rev": latest["rev"] if latest else None,
    }


# --------------------------------------------------------------------------
# Citations
# --------------------------------------------------------------------------

CITATION_PATTERNS = [
    ("iso_iec", re.compile(
        r"\b(?:EN\s+)?(?:ANSI/)?(?:AAMI/)?(ISO|IEC)(?:/(?:TR|TS|IEEE|ASTM))?\s*"
        r"(\d{3,5}(?:-\d+)*)\s*(?::\s*(\d{4}))?", re.I)),
    ("cfr", re.compile(r"\b21\s*CFR\s*(?:Part\s*)?(\d{1,4}(?:\.\d+)?)", re.I)),
    ("eu", re.compile(r"\(?EU\)?\s*(20\d{2}/\d{3,4})", re.I)),
    ("clsi", re.compile(r"\bCLSI\s+([A-Z]{2}\d{2}(?:-A\d?|\s+\d(?:st|nd|rd|th)\s+ed\.?)?)", re.I)),
    ("mdsap", re.compile(r"\bMDSAP\b", re.I)),
]


def harvest_citations(sections):
    out = []
    seen = set()
    for sec in sections:
        blob = " ".join([sec["heading"]] + sec["paragraphs"])
        for kind, pat in CITATION_PATTERNS:
            for m in pat.finditer(blob):
                text = normalise(m.group(0))
                key = (kind, text.upper(), sec["number"] or sec["heading"])
                if key in seen:
                    continue
                seen.add(key)
                year = None
                if kind == "iso_iec":
                    year = m.group(3)
                out.append(
                    {
                        "kind": kind,
                        "text": text,
                        "edition_year": year,
                        "section": sec["number"] or None,
                        "heading": sec["heading"],
                    }
                )
    return out


# --------------------------------------------------------------------------
# Hygiene checks
# --------------------------------------------------------------------------

PLACEHOLDER_RE = re.compile(
    r"(\bTBD\b|\bTBC\b|\bXXX+\b|\bLorem ipsum\b|\?\?\?|<insert[^>]*>|\[insert[^\]]*\]|\bPLACEHOLDER\b)",
    re.I,
)
XREF_SECTION_RE = re.compile(r"\b(?:see\s+|refer\s+to\s+|per\s+|in\s+)?Section\s+(\d+(?:\.\d+)*)", re.I)
XREF_TABLE_RE = re.compile(r"\bTable\s+(\d+)", re.I)
XREF_FIGURE_RE = re.compile(r"\b(?:Figure|Fig\.?)\s+(\d+)", re.I)
ACRONYM_DEF_RE = re.compile(r"([A-Z][A-Za-z][A-Za-z \-/]{3,60}?)\s*\(([A-Z]{2,6}s?)\)")
ACRONYM_USE_RE = re.compile(r"\b([A-Z]{2,6})\b")

ACRONYM_STOPLIST = {
    "AND", "THE", "FOR", "NOT", "ALL", "ANY", "USE", "MAY", "SHALL", "MUST",
    "ISO", "IEC", "EN", "CFR", "EU", "US", "USA", "FDA", "CE", "UDI", "QA", "QC",
    "ID", "OK", "NA", "TBD", "TBC", "PDF", "XML", "HTML", "CSV", "URL", "REF",
    "LOT", "MIN", "MAX", "AVG", "STD", "DNA", "RNA", "PCR", "NGS", "IVD",
    "CLSI", "AAMI", "ANSI", "CEN", "MDSAP", "WHO", "EC", "UK", "MDR", "IVDR",
    "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X",
}


def check_numbering(sections):
    """Report gaps and duplicates in the numbered heading sequence."""
    issues = []
    seen = {}
    numbers = [s["number"] for s in sections if s["number"]]
    for n in numbers:
        if n in seen:
            issues.append({"type": "duplicate_heading_number", "number": n})
        seen[n] = True
    by_parent = {}
    for n in numbers:
        parts = n.split(".")
        parent = ".".join(parts[:-1])
        by_parent.setdefault(parent, []).append(int(parts[-1]) if parts[-1].isdigit() else None)
    for parent, kids in by_parent.items():
        kids = [k for k in kids if k is not None]
        if not kids:
            continue
        expected = list(range(min(kids), max(kids) + 1))
        missing = sorted(set(expected) - set(kids))
        if missing:
            issues.append(
                {
                    "type": "heading_number_gap",
                    "parent": parent or "(top level)",
                    "missing": [f"{parent}.{m}" if parent else str(m) for m in missing],
                    "present": sorted(set(kids)),
                }
            )
    return issues


def check_cross_references(sections, table_count, figure_numbers):
    known_sections = {s["number"] for s in sections if s["number"]}
    issues = []
    for sec in sections:
        blob = " ".join(sec["paragraphs"])
        for m in XREF_SECTION_RE.finditer(blob):
            target = m.group(1)
            if known_sections and target not in known_sections:
                issues.append(
                    {
                        "type": "broken_section_reference",
                        "in_section": sec["number"] or sec["heading"],
                        "target": f"Section {target}",
                        "context": normalise(blob[max(0, m.start() - 60):m.end() + 60]),
                    }
                )
        for m in XREF_TABLE_RE.finditer(blob):
            if table_count and int(m.group(1)) > table_count:
                issues.append(
                    {
                        "type": "broken_table_reference",
                        "in_section": sec["number"] or sec["heading"],
                        "target": f"Table {m.group(1)}",
                        "tables_in_document": table_count,
                    }
                )
        for m in XREF_FIGURE_RE.finditer(blob):
            if figure_numbers and int(m.group(1)) not in figure_numbers:
                issues.append(
                    {
                        "type": "broken_figure_reference",
                        "in_section": sec["number"] or sec["heading"],
                        "target": f"Figure {m.group(1)}",
                        "figures_in_document": sorted(figure_numbers),
                    }
                )
    return issues


def check_acronyms(sections):
    defined, first_use = {}, {}
    for sec in sections:
        blob = " ".join([sec["heading"]] + sec["paragraphs"])
        # Document identifiers (VER-0045, DES-0028) are not acronyms.
        blob = DOC_ID_RE.sub(" ", blob)
        for m in ACRONYM_DEF_RE.finditer(blob):
            acr = m.group(2).upper()
            defined.setdefault(acr, {"expansion": normalise(m.group(1)),
                                     "section": sec["number"] or sec["heading"]})
        for m in ACRONYM_USE_RE.finditer(blob):
            acr = m.group(1).upper()
            if acr in ACRONYM_STOPLIST or len(acr) < 2:
                continue
            first_use.setdefault(acr, sec["number"] or sec["heading"])
    undefined = [
        {"acronym": a, "first_used_in": s}
        for a, s in sorted(first_use.items())
        if a not in defined
    ]
    unused = [
        {"acronym": a, "defined_in": d["section"], "expansion": d["expansion"]}
        for a, d in sorted(defined.items())
        if a not in first_use
    ]
    return {"defined": defined, "undefined": undefined, "defined_but_unused": unused}


def check_placeholders(sections):
    hits = []
    for sec in sections:
        for p in sec["paragraphs"]:
            for m in PLACEHOLDER_RE.finditer(p):
                hits.append(
                    {
                        "section": sec["number"] or sec["heading"],
                        "match": m.group(0),
                        "context": normalise(p[max(0, m.start() - 60):m.end() + 60]),
                    }
                )
    return hits


def check_markup_left_in(document):
    """Tracked changes, comments and highlighted runs still in the file."""
    body = document.element.body
    out = {
        "tracked_insertions": len(body.findall(f".//{qn('w:ins')}")),
        "tracked_deletions": len(body.findall(f".//{qn('w:del')}")),
        "highlighted_runs": [],
        "has_comments": False,
    }
    for hl in body.findall(f".//{qn('w:highlight')}"):
        val = hl.get(qn("w:val"))
        if val and val != "none":
            run = hl.getparent()
            run = run.getparent() if run is not None else None
            text = "".join(t.text or "" for t in run.findall(qn("w:t"))) if run is not None else ""
            out["highlighted_runs"].append({"colour": val, "text": normalise(text)[:120]})
    try:
        for rel in document.part.rels.values():
            if "comments" in rel.reltype:
                out["has_comments"] = True
    except Exception:  # pragma: no cover - relationship shapes vary
        pass
    return out


# --------------------------------------------------------------------------
# Header / footer / title
# --------------------------------------------------------------------------

def collect_headers_footers(document):
    headers, footers = [], []
    for i, sec in enumerate(document.sections):
        for kind, part, bucket in (
            ("default", sec.header, headers),
            ("first_page", sec.first_page_header, headers),
            ("even_page", sec.even_page_header, headers),
            ("default", sec.footer, footers),
            ("first_page", sec.first_page_footer, footers),
            ("even_page", sec.even_page_footer, footers),
        ):
            try:
                if part.is_linked_to_previous and i > 0:
                    continue
                t = text_of(part)
            except Exception:
                continue
            if t.strip():
                bucket.append({"section": i, "type": kind, "text": t})
    return headers, footers


def has_page_fields(document):
    """True when the footer/header uses PAGE and NUMPAGES fields."""
    found = {"PAGE": False, "NUMPAGES": False}
    for sec in document.sections:
        for part in (sec.header, sec.footer, sec.first_page_header, sec.first_page_footer):
            try:
                xml = part._element.xml
            except Exception:
                continue
            for field in found:
                if re.search(rf"\b{field}\b", xml):
                    found[field] = True
    return found


def parse_header(header_text):
    """Pull the document ID, revision and residual title out of header text."""
    ids = sorted({f"{m.group(1)}-{m.group(2)}" for m in DOC_ID_RE.finditer(header_text)})
    revs = sorted({m.group(1).upper() for m in REV_RE.finditer(header_text)})
    title_lines = []
    for line in header_text.splitlines():
        stripped = normalise(line)
        if not stripped:
            continue
        cleaned = DOC_ID_RE.sub("", stripped)
        cleaned = REV_RE.sub("", cleaned)
        cleaned = re.sub(r"Page\s*\d*\s*of\s*\d*", "", cleaned, flags=re.I)
        cleaned = normalise(re.sub(r"[\-–|:,]+", " ", cleaned))
        if len(cleaned) > 3:
            title_lines.append(cleaned)
    return {
        "doc_ids": ids,
        "revisions": revs,
        "title_candidates": title_lines,
    }


# --------------------------------------------------------------------------
# Body
# --------------------------------------------------------------------------

def build_sections(document):
    """Walk the body once, producing an ordered section tree."""
    sections = []
    tables = []
    figure_numbers = set()
    title_paragraph = None
    current = {
        "number": None,
        "heading": "(front matter)",
        "level": 0,
        "paragraphs": [],
        "tables": [],
    }
    for block in iter_block_items(document):
        if isinstance(block, Paragraph):
            text = normalise(block.text)
            if not text:
                continue
            lvl = heading_level(block)
            if lvl == 0:
                # The document title is not a section of its own; recording it
                # as one makes a retitled document look like a section was
                # added and another removed.
                if title_paragraph is None:
                    title_paragraph = text
                current["paragraphs"].append(text)
                continue
            if lvl is not None:
                sections.append(current)
                m = HEADING_NUM_RE.match(text)
                current = {
                    "number": m.group(1) if m else None,
                    "heading": text,
                    "title": m.group(2) if m else text,
                    "level": lvl,
                    "paragraphs": [],
                    "tables": [],
                }
                continue
            m = re.match(r"^(?:Figure|Fig\.?)\s+(\d+)", text, re.I)
            if m:
                figure_numbers.add(int(m.group(1)))
            current["paragraphs"].append(text)
        else:
            t_index = len(tables)
            tables.append(block)
            current["tables"].append(
                {"table_index": t_index, "rows": table_text_rows(block)}
            )
    sections.append(current)
    for s in sections:
        s.setdefault("title", s["heading"])
    return sections, tables, figure_numbers, title_paragraph


def document_title(document, sections, title_paragraph=None):
    core = None
    try:
        core = normalise(document.core_properties.title) or None
    except Exception:
        pass
    first_heading = None
    for s in sections:
        if s["heading"] and s["heading"] != "(front matter)":
            first_heading = s["heading"]
            break
    front = sections[0]["paragraphs"] if sections else []
    first_body_line = front[0] if front else None
    return {
        "core_properties_title": core,
        "title_paragraph": title_paragraph,
        "first_heading": first_heading,
        "first_body_line": first_body_line,
    }


# --------------------------------------------------------------------------
# Auto checks (mechanical only)
# --------------------------------------------------------------------------

def auto_checks(header_info, title, revision_history, page_fields, footers):
    checks = []

    header_revs = header_info["revisions"]
    header_ids = header_info["doc_ids"]

    # Check 1 - title vs header
    title_values = [v for v in (
        title.get("core_properties_title"),
        title.get("title_paragraph"),
        title.get("first_heading"),
        title.get("first_body_line"),
    ) if v]
    header_titles = header_info["title_candidates"]

    def squash(s):
        return re.sub(r"[^a-z0-9]", "", (s or "").lower())

    title_match = None
    if header_titles and title_values:
        title_match = any(
            squash(h) and squash(t) and (squash(h) in squash(t) or squash(t) in squash(h))
            for h in header_titles
            for t in title_values
        )
    checks.append(
        {
            "check": "1 Title vs header",
            "result": "PASS" if title_match else ("REVIEW" if title_match is None else "FAIL"),
            "header_title_candidates": header_titles,
            "document_title_candidates": title_values,
            "note": "Compare wording exactly; the automated match ignores punctuation and case.",
        }
    )

    # Check 2 - header revision vs revision history
    hist_rev = revision_history["latest_rev"] if revision_history else None
    if not header_revs:
        result = "REVIEW"
    elif hist_rev is None:
        result = "REVIEW"
    elif len(set(header_revs)) == 1 and header_revs[0] == hist_rev:
        result = "PASS"
    else:
        result = "FAIL"
    checks.append(
        {
            "check": "2 Header revision vs revision history",
            "result": result,
            "header_revisions": header_revs,
            "revision_history_latest": hist_rev,
            "revision_history_row_index": revision_history["latest_entry_row_index"] if revision_history else None,
        }
    )

    checks.append(
        {
            "check": "2b Document ID in header",
            "result": "PASS" if len(header_ids) == 1 else ("FAIL" if len(header_ids) > 1 else "REVIEW"),
            "doc_ids_in_header": header_ids,
        }
    )

    checks.append(
        {
            "check": "7a Page x of y fields present",
            "result": "PASS" if page_fields["PAGE"] and page_fields["NUMPAGES"] else "REVIEW",
            "fields": page_fields,
            "note": "A literal 'Page 1 of 7' typed as text will not renumber; PAGE/NUMPAGES fields do.",
        }
    )

    checks.append(
        {
            "check": "7b Footer present",
            "result": "PASS" if footers else "FAIL",
            "footer_text": [f["text"] for f in footers],
        }
    )

    return checks


# --------------------------------------------------------------------------
# PDF path (review only - a PDF cannot carry the red/orange annotations)
# --------------------------------------------------------------------------

def running_blocks(pages, band=4):
    """Recover the running header and footer from a PDF's page text.

    A PDF has no header part: the header is simply the first few lines of every
    page. Lines that recur in the top (or bottom) band of most pages are the
    running block. Digits are masked for the recurrence test so that
    "Page 1 of 7" still matches "Page 2 of 7", but the text is reported as it
    appears on the first page that carries it.
    """
    real = [p for p in pages if p.strip()]
    if not real:
        return "", ""
    threshold = max(2, int(round(0.6 * len(real)))) if len(real) > 1 else 1

    def collect(top):
        counts, first_seen, order = {}, {}, []
        for page in real:
            lines = [ln for ln in page.splitlines() if ln.strip()]
            band_lines = lines[:band] if top else lines[-band:]
            for pos, line in enumerate(band_lines):
                key = re.sub(r"\d", "#", normalise(line))
                if not key:
                    continue
                counts[key] = counts.get(key, 0) + 1
                if key not in first_seen:
                    first_seen[key] = normalise(line)
                    order.append((pos, key))
        kept = [first_seen[k] for _, k in sorted(order) if counts[k] >= threshold]
        # De-duplicate while keeping order.
        seen, out = set(), []
        for line in kept:
            if line not in seen:
                seen.add(line)
                out.append(line)
        return "\n".join(out)

    return collect(True), collect(False)


def extract_pdf(path):
    try:
        import pdfplumber
    except ImportError:  # pragma: no cover
        sys.exit("pdfplumber is required for PDF input: pip install pdfplumber")

    pages, tables = [], []
    with pdfplumber.open(str(path)) as pdf:
        for page in pdf.pages:
            pages.append(page.extract_text() or "")
            for t in page.extract_tables() or []:
                tables.append([[normalise(c) for c in row] for row in t])

    full = "\n".join(pages)
    text_chars = len(re.sub(r"\s", "", full))
    has_text_layer = text_chars > 50 * max(1, len(pages))

    header_text, footer_text = running_blocks(pages)

    # The running header and footer repeat on every page; leaving them in the
    # body would make them look like content and pollute every check.
    running = {normalise(l) for l in (header_text + "\n" + footer_text).splitlines() if l.strip()}
    running_masked = {re.sub(r"\d", "#", r) for r in running}

    sections = []
    current = {"number": None, "heading": "(front matter)", "title": "(front matter)",
               "level": 0, "paragraphs": [], "tables": []}
    for line in full.splitlines():
        text = normalise(line)
        if not text:
            continue
        if text in running or re.sub(r"\d", "#", text) in running_masked:
            continue
        m = HEADING_NUM_RE.match(text)
        # A PDF has no styles, so a heading can only be guessed from the text.
        # Table rows arrive inline: "1 2024-03-15 J. Chen Initial release."
        # would otherwise become "Section 1" a second time.
        if m and (
            NOT_A_HEADING_RE.match(text)
            or re.search(r"\d{4}-\d{2}-\d{2}|\d{1,2}/\d{1,2}/\d{2,4}", text)
            or not re.match(r"^[A-Z(]", m.group(2))
        ):
            m = None
        if m and len(text) < 120:
            sections.append(current)
            current = {
                "number": m.group(1),
                "heading": text,
                "title": m.group(2),
                "level": m.group(1).count(".") + 1,
                "paragraphs": [],
                "tables": [],
            }
        else:
            current["paragraphs"].append(text)
    sections.append(current)

    rev_hist = None
    for i, rows in enumerate(tables):
        head = [normalise(c).lower() for c in rows[0]] if rows else []
        if any(REV_COL.search(c) for c in head) and any(DESC_COL.search(c) for c in head):
            cols = {}
            for j, c in enumerate(head):
                if "rev" not in cols and REV_COL.search(c):
                    cols["rev"] = j
                elif "date" not in cols and DATE_COL.search(c):
                    cols["date"] = j
                elif "author" not in cols and AUTHOR_COL.search(c):
                    cols["author"] = j
                elif "description" not in cols and DESC_COL.search(c):
                    cols["description"] = j
            entries = []
            for r, row in enumerate(rows[1:], start=1):
                def get(key):
                    k = cols.get(key)
                    return normalise(row[k]) if k is not None and k < len(row) else ""
                rev_raw = get("rev")
                m = REV_RE.search(rev_raw) if rev_raw else None
                entries.append({
                    "row_index": r, "rev_raw": rev_raw,
                    "rev": (m.group(1) if m else rev_raw).upper() or None,
                    "date": get("date"), "author": get("author"), "dcr": "",
                    "description": get("description"), "cells": row,
                })
            latest = max(entries, key=lambda e: revision_sort_key(e["rev"])) if entries else None
            rev_hist = {
                "table_index": i, "header_row_index": 0, "header_row": rows[0],
                "columns": cols, "entries": entries,
                "latest_entry_row_index": latest["row_index"] if latest else None,
                "latest_rev": latest["rev"] if latest else None,
            }
            break

    header_info = parse_header(header_text)
    title = {
        "core_properties_title": None,
        "title_paragraph": sections[0]["paragraphs"][0] if sections[0]["paragraphs"] else None,
        "first_heading": next((s["heading"] for s in sections if s["heading"] != "(front matter)"), None),
        "first_body_line": sections[0]["paragraphs"][0] if sections[0]["paragraphs"] else None,
    }
    # A PDF has no PAGE/NUMPAGES fields to inspect; a printed "Page x of y" in
    # the recovered footer is the most that can be said.
    page_fields = {"PAGE": False, "NUMPAGES": False}
    footers = (
        [{"section": 0, "type": "running", "text": footer_text}] if footer_text else []
    )

    return {
        "file": str(path),
        "format": "pdf",
        "annotatable": False,
        "has_text_layer": has_text_layer,
        "page_count": len(pages),
        "headers": [{"section": 0, "type": "running", "text": header_text}] if header_text else [],
        "footers": footers,
        "header_info": header_info,
        "title": title,
        "revision_history": rev_hist,
        "sections": sections,
        "tables": [{"table_index": i, "rows": rows} for i, rows in enumerate(tables)],
        "citations": harvest_citations(sections),
        "hygiene": {
            "placeholders": check_placeholders(sections),
            "numbering": check_numbering(sections),
            "cross_references": check_cross_references(sections, len(tables), set()),
            "acronyms": check_acronyms(sections),
            "markup_left_in": None,
        },
        "auto_checks": auto_checks(header_info, title, rev_hist, page_fields, footers),
        "warnings": ([] if has_text_layer else [
            "This PDF has little or no text layer - it is probably a scan. "
            "Do not report PASS on any check; the document could not be read."
        ]) + [
            "PDF input: the document cannot be annotated in place. "
            "Produce a separate .docx review record instead."
        ],
    }


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------

def extract_docx(path):
    document = docx.Document(str(path))
    headers, footers = collect_headers_footers(document)
    header_text = "\n".join(h["text"] for h in headers)
    header_info = parse_header(header_text)
    sections, tables, figure_numbers, title_paragraph = build_sections(document)
    title = document_title(document, sections, title_paragraph)
    rev_hist = find_revision_history(tables)
    page_fields = has_page_fields(document)

    warnings = []
    if rev_hist is None:
        warnings.append(
            "No revision history table was found. Checks 2, 3 and 4 cannot run and "
            "nothing may be annotated - report this and stop."
        )
    if not headers:
        warnings.append("No page header was found in any w:hdr part.")

    return {
        "file": str(path),
        "format": "docx",
        "annotatable": True,
        "headers": headers,
        "footers": footers,
        "header_info": header_info,
        "title": title,
        "revision_history": rev_hist,
        "sections": sections,
        "tables": [
            {"table_index": i, "rows": table_text_rows(t)} for i, t in enumerate(tables)
        ],
        "citations": harvest_citations(sections),
        "hygiene": {
            "placeholders": check_placeholders(sections),
            "numbering": check_numbering(sections),
            "cross_references": check_cross_references(sections, len(tables), figure_numbers),
            "acronyms": check_acronyms(sections),
            "markup_left_in": check_markup_left_in(document),
        },
        "auto_checks": auto_checks(header_info, title, rev_hist, page_fields, footers),
        "warnings": warnings,
    }


def extract(path):
    suffix = path.suffix.lower()
    if suffix == ".docx":
        return extract_docx(path)
    if suffix == ".pdf":
        return extract_pdf(path)
    if suffix == ".doc":
        sys.exit(
            ".doc (Word 97-2003) is not supported. Ask the author to save as .docx."
        )
    sys.exit(f"Unsupported file type '{suffix}'. Provide a .docx or .pdf.")


def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("document")
    ap.add_argument("--out", required=True, help="directory for extract.json")
    a = ap.parse_args()

    path = Path(a.document)
    if not path.exists():
        sys.exit(f"No such file: {path}")
    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)

    data = extract(path)
    target = out / "extract.json"
    target.write_text(json.dumps(data, indent=2, ensure_ascii=False))

    hi = data["header_info"]
    print(f"format        : {data['format']}")
    print(f"doc id        : {', '.join(hi['doc_ids']) or '(none found in header)'}")
    print(f"header rev    : {', '.join(hi['revisions']) or '(none found in header)'}")
    rh = data["revision_history"]
    print(f"rev history   : {len(rh['entries'])} entries, latest Rev {rh['latest_rev']}"
          if rh else "rev history   : NOT FOUND")
    print(f"sections      : {len(data['sections'])}")
    print(f"citations     : {len(data['citations'])}")
    for c in data["auto_checks"]:
        print(f"[{c['result']:6}] {c['check']}")
    for w in data["warnings"]:
        print(f"[WARN  ] {w}")
    print(f"wrote {target}")


if __name__ == "__main__":
    main()
