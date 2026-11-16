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
    HeadingNumberer,
    heading_level,
    iter_block_items,
    major_revision,
    normalise,
    revision_sort_key,
    table_rows,
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
AUTHOR_COL = re.compile(r"\b(author|by|owner|prepared|originator|revised by|name|initials)\b", re.I)
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
# Numbering self-validation
# --------------------------------------------------------------------------

#: "§ 5.6 Product Assessment", "section 2 References", "Section 5.5.1".
#: The keyword is matched case-insensitively but the title must genuinely start
#: with a capital, so "Section 3 to 2-8 degrees C" does not read as a title.
CITED_SECTION_RE = re.compile(
    r"(?:§|[Ss]ections?|[Ss]teps?)\s*(\d+(?:\.\d+)*)\s*[-,:–]?\s*"
    r"([A-Z][A-Za-z][A-Za-z \-/&]{2,48})?"
)

#: Words that begin a clause, not a section title.
_NOT_A_TITLE_START = {
    "to", "and", "the", "for", "in", "of", "is", "was", "were", "see", "as", "at",
    "by", "from", "with", "this", "that", "it", "has", "have", "had", "a", "an",
    "corrected", "added", "removed", "updated", "edited", "changed",
}


def squash_title(s):
    return re.sub(r"[^a-z0-9]", "", (s or "").lower())


def find_section_by_title(sections, cited):
    """The section whose title best matches a title cited in prose.

    A citation's title is rarely captured exactly: the regex may stop short
    ("Edit and delete any") or run on into the next clause ("Product Assessment
    Documentation -added"). Exact matching therefore misses the section and the
    caller wrongly concludes it was removed, so match on containment either way
    and keep the longest candidate.
    """
    b = squash_title(cited)
    if not b or len(b) < 6:
        return None
    best, best_len = None, -1
    for s in sections:
        a = squash_title(s.get("title"))
        if not a:
            continue
        if a == b or a.startswith(b) or b.startswith(a) or b in a or a in b:
            if len(a) > best_len:
                best, best_len = s, len(a)
    return best


def validate_numbering(sections, revision_history):
    """Check reconstructed section numbers against the document's own citations.

    A revision history usually cites its own sections by number *and* title
    ("§ 2 References", "section 5.6 Product Assessment Documentation"). That is
    free ground truth for the numbering rebuilt from the styles, so use it: if
    the pairs agree the numbers can be quoted in findings, and if they disagree
    they must not be.
    """
    by_number = {s["number"]: s for s in sections if s.get("number")}
    if not revision_history or not by_number:
        return {
            "checked": 0, "agreed": 0, "disagreed": [], "unmatched": [],
            "reliable": False,
            "note": "No revision history or no section numbers to validate against.",
        }

    def squash(s):
        return re.sub(r"[^a-z0-9]", "", (s or "").lower())

    def find_by_title(cited):
        return find_section_by_title(sections, cited)

    agreed, mismatches, stale = [], [], []
    for entry in revision_history["entries"]:
        for m in CITED_SECTION_RE.finditer(entry["description"] or ""):
            number, cited_title = m.group(1), (m.group(2) or "").strip()
            if not cited_title or len(cited_title) < 4:
                continue
            if cited_title.split()[0].lower() in _NOT_A_TITLE_START:
                continue
            section = by_number.get(number)
            if section is not None:
                a, b = squash(section["title"]), squash(cited_title)
                if a and b and (a.startswith(b) or b.startswith(a) or b in a):
                    agreed.append({"number": number, "title": section["title"], "cited": cited_title})
                    continue
            # The cited number does not carry the cited title. If that title
            # exists elsewhere in the document, the section has moved since the
            # entry was written: the entry is stale, the numbering is not wrong.
            moved_to = find_by_title(cited_title)
            if moved_to is not None:
                stale.append(
                    {
                        "rev": entry["rev"],
                        "row_index": entry["row_index"],
                        "cited_as": number,
                        "title": moved_to["title"],
                        "now": moved_to.get("number") or "unnumbered",
                    }
                )
            else:
                mismatches.append(
                    {
                        "rev": entry["rev"],
                        "cited_as": number,
                        "cited_title": cited_title,
                        "reconstructed_title": section["title"] if section else None,
                    }
                )

    checked = len(agreed) + len(mismatches)
    return {
        "checked": checked,
        "agreed": agreed,
        "mismatches": mismatches,
        # Not a numbering problem: historical entries citing numbers that have
        # since changed. Report these under check 3.
        "stale_citations": stale,
        # Numbers may be quoted only when the document's own citations
        # corroborate them and none genuinely contradicts them.
        "reliable": len(agreed) > 0 and not mismatches,
        "note": (
            f"{len(agreed)} of {checked} number/title pairs cited in the revision history "
            f"match the reconstructed numbering"
            + (f"; {len(stale)} further citation(s) name a section that has since moved"
               if stale else "")
            + "."
            if checked
            else "The revision history cites no section number with a title, so the "
                 "reconstructed numbering could not be corroborated."
        ),
    }


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
    # Ordinary words that appear capitalised in templates and headings.
    "OF", "AUTO", "REPORT", "SYSTEM", "FILL", "RAW", "TUBES", "PRIMER", "SCOPE",
    "EXAMPLE", "METRIC", "DATA", "TOTAL", "NOTE", "STEP", "STEPS", "TABLE",
    "FIGURE", "PASS", "FAIL", "YES", "PER", "VIA", "ALL", "NEW", "OLD", "END",
    "AND", "OR", "IF", "SEE", "TYPE", "NAME", "DATE", "REV", "PAGE", "TITLE",
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


def check_cross_references(sections, table_count, figure_numbers, numbering_reliable=True):
    known_sections = {s["number"] for s in sections if s["number"]}
    issues = []
    for sec in sections:
        blob = " ".join(sec["paragraphs"])
        for m in XREF_SECTION_RE.finditer(blob):
            target = m.group(1)
            if known_sections and target not in known_sections:
                # Calling a reference broken when the section numbering could
                # not be rebuilt reliably invents findings: the section very
                # likely exists and simply has no number we can read.
                issues.append(
                    {
                        "type": (
                            "broken_section_reference" if numbering_reliable
                            else "unverified_section_reference"
                        ),
                        "severity_hint": "Minor" if numbering_reliable else "Observation",
                        "in_section": sec["number"] or sec["heading"],
                        "target": f"Section {target}",
                        "context": normalise(blob[max(0, m.start() - 60):m.end() + 60]),
                        **({} if numbering_reliable else {
                            "note": "Section numbering could not be corroborated, so this "
                                    "reference could not be resolved either way. Do not "
                                    "report it as broken."
                        }),
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

    def is_all_caps(s):
        letters = [c for c in s if c.isalpha()]
        return bool(letters) and all(c.isupper() for c in letters)

    for sec in sections:
        where = sec["number"] or sec["heading"]

        # A definitions section usually holds its terms in a table, not in
        # "Expansion (ACR)" prose. Read those rows as definitions, or every
        # term the document properly defines is reported as undefined.
        for tbl in sec.get("tables", []):
            for row in tbl.get("rows", []):
                cells = [normalise(c) for c in row if normalise(c)]
                if len(cells) < 2:
                    continue
                term = cells[0].strip().rstrip(":")
                if 2 <= len(term) <= 12 and re.fullmatch(r"[A-Za-z0-9/\-]+", term):
                    defined.setdefault(
                        term.upper(), {"expansion": cells[1][:120], "section": where}
                    )

        for text in [sec["heading"]] + sec["paragraphs"]:
            # Document identifiers (VER-0045, DES-0028) are not acronyms.
            blob = DOC_ID_RE.sub(" ", text)
            for m in ACRONYM_DEF_RE.finditer(blob):
                acr = m.group(2).upper()
                defined.setdefault(
                    acr, {"expansion": normalise(m.group(1)), "section": where}
                )
            # In an ALL-CAPS heading or caption every word looks like an
            # acronym ("RAW INDEXING PRIMER CATEGORIES"), so skip those
            # entirely rather than invent a finding per word.
            if is_all_caps(blob):
                continue
            for m in ACRONYM_USE_RE.finditer(blob):
                acr = m.group(1).upper()
                if acr in ACRONYM_STOPLIST or len(acr) < 2:
                    continue
                first_use.setdefault(acr, where)
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
                xml = part._element.xml
            except Exception:
                continue
            has_fields = bool(re.search(r"\b(PAGE|NUMPAGES|SECTIONPAGES)\b", xml))
            has_runs = qn("w:r") in xml or "<w:r>" in xml or "<w:r " in xml
            # A footer holding only PAGE fields has no text at all. Dropping it
            # would report "no footer found" for a document that has one, so
            # keep it and let the check say what it could not read.
            if t.strip() or has_fields or has_runs:
                bucket.append(
                    {"section": i, "type": kind, "text": t, "has_fields": has_fields}
                )
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


#: A header cell holding nothing but a revision: "8", "8.04", "B".
BARE_REVISION_RE = re.compile(r"^(\d+(?:\.\d+)?|[A-Z])$")


def cell_has_page_field(cell):
    """True if the cell's content comes from a PAGE / NUMPAGES field."""
    try:
        xml = cell._tc.xml
    except Exception:
        return False
    return bool(re.search(r"\b(PAGE|NUMPAGES|SECTIONPAGES)\b", xml))


def find_header_revision(document, header_text, revision_history):
    """Find the revision in the page header, labelled or not.

    Templates often put the revision in a header table cell as a bare number
    with no "Rev." label at all. A bare number in a header is ambiguous - it
    could be a page number - so an unlabelled candidate is only trusted when
    the revision history corroborates it.
    """
    labelled = sorted({m.group(1).upper() for m in REV_RE.finditer(header_text)})
    if labelled:
        return {
            "revisions": labelled,
            "revision": labelled[0] if len(labelled) == 1 else None,
            "revision_source": "labelled in the page header",
            "revision_confidence": "labelled",
        }

    doc_ids = {f"{m.group(1)}-{m.group(2)}" for m in DOC_ID_RE.finditer(header_text)}
    candidates = []
    for sec in document.sections:
        for part in (sec.header, sec.first_page_header, sec.even_page_header):
            try:
                tables = part.tables
            except Exception:
                continue
            for tbl in tables:
                for row in table_rows(tbl):
                    for cell in row:
                        text = normalise(cell.text)
                        if not text or text in doc_ids:
                            continue
                        if not BARE_REVISION_RE.match(text.upper()):
                            continue
                        if cell_has_page_field(cell):
                            continue  # a page number, not a revision
                        if text.upper() not in candidates:
                            candidates.append(text.upper())

    if not candidates:
        return {
            "revisions": [], "revision": None,
            "revision_source": None,
            "revision_confidence": "not found",
        }

    history_revs = set()
    latest_major = None
    if revision_history:
        history_revs = {e["rev"] for e in revision_history["entries"] if e["rev"]}
        latest_major = major_revision(revision_history.get("latest_rev"))

    for cand in candidates:
        if cand in history_revs or (latest_major and major_revision(cand) == latest_major):
            return {
                "revisions": [cand], "revision": cand,
                "revision_source": "unlabelled page-header cell, corroborated by the revision history",
                "revision_confidence": "inferred-corroborated",
            }

    return {
        "revisions": candidates, "revision": None,
        "revision_source": "unlabelled page-header cell, not corroborated",
        "revision_confidence": "uncorroborated",
    }


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
        # Kept un-normalised so checks can see punctuation the cleanup removes.
        "raw_lines": [normalise(l) for l in header_text.splitlines() if normalise(l)],
    }


# --------------------------------------------------------------------------
# Body
# --------------------------------------------------------------------------

def build_sections(document):
    """Walk the body once, producing an ordered section tree."""
    # Rebuild the numbers Word renders from the Heading styles, since most
    # templates keep them out of the text entirely. Done in this single walk so
    # the numbers cannot drift onto the wrong paragraphs.
    numberer = HeadingNumberer(document)
    number_source = numberer.source
    # When the Heading styles already number the headings, the document has a
    # structure to trust: a numbered paragraph in a body style is a step, not a
    # section, so do not promote it from its text.
    allow_text_headings = number_source != "resolved"
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
            lvl = heading_level(block, allow_text_fallback=allow_text_headings)
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
                if m:
                    number, title, src = m.group(1), m.group(2), "text"
                else:
                    number = numberer.number_for(block, lvl)
                    title, src = text, (number_source if number else None)
                current = {
                    "number": number,
                    "heading": text,
                    "title": title,
                    "number_source": src,
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
        s.setdefault("number_source", None)
    return sections, tables, figure_numbers, title_paragraph, number_source


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

    # A document whose only self-naming is the page header has nothing to
    # compare against: no Title-styled paragraph and no Word title property.
    # Comparing the header against the first body heading ("PURPOSE AND SCOPE")
    # would fail every document of that kind, which is a false positive.
    own_title = title.get("core_properties_title") or title.get("title_paragraph")
    if not own_title:
        result = "REVIEW"
        note = (
            "The page header is the only place this document names itself: there is no "
            "title block on page 1 and no Word document title property, so there is "
            "nothing to compare the header against. Confirm the header title is correct "
            "for the document's scope by reading it."
        )
    elif header_titles:
        matched = any(
            squash(h) and squash(t) and (squash(h) in squash(t) or squash(t) in squash(h))
            for h in header_titles
            for t in title_values
        )
        result = "PASS" if matched else "FAIL"
        note = "Compare wording exactly; the automated match ignores punctuation and case."
    else:
        result = "REVIEW"
        note = "No title could be read from the page header."

    stray = [
        l for l in header_info.get("raw_lines", [])
        if re.search(r"[,;]\s*$", l.replace("|", "").strip())
    ]
    if stray:
        note += (
            f" Observation: the header title ends in stray punctuation ({stray!r});"
            " worth tidying."
        )

    checks.append(
        {
            "check": "1 Title vs header",
            "result": result,
            "header_title_candidates": header_titles,
            "document_title_candidates": title_values,
            "note": note,
        }
    )

    # Check 2 - header revision vs revision history.
    #
    # The header may legitimately carry only the major revision: a minor uprev
    # (8.01 - 8.04) deliberately leaves the header at 8. So compare the major
    # components, and report the exact match separately rather than failing a
    # document that follows that convention.
    hist_rev = revision_history["latest_rev"] if revision_history else None
    header_rev = header_info.get("revision")
    confidence = header_info.get("revision_confidence")
    exact = header_rev is not None and hist_rev is not None and header_rev == hist_rev
    major_match = (
        header_rev is not None and hist_rev is not None
        and major_revision(header_rev) == major_revision(hist_rev)
    )

    if header_rev is None or hist_rev is None or confidence == "uncorroborated":
        result = "REVIEW"
    elif exact or major_match:
        result = "PASS"
    else:
        result = "FAIL"

    note = None
    if result == "PASS" and not exact:
        note = (
            f"Header reads Rev {header_rev}; the latest revision history entry is "
            f"Rev {hist_rev}. The major revisions agree, which is the expected "
            "convention for a minor uprev - not a finding."
        )
    elif confidence == "uncorroborated":
        note = (
            f"A bare value {header_revs} was found in the page header but the revision "
            "history does not corroborate it as a revision. Confirm the header revision "
            "manually; do not report PASS or FAIL."
        )
    elif header_rev is None and not header_revs:
        note = "No revision could be found in the page header."

    checks.append(
        {
            "check": "2 Header revision vs revision history",
            "result": result,
            "header_revision": header_rev,
            "header_revisions": header_revs,
            "revision_source": header_info.get("revision_source"),
            "revision_confidence": confidence,
            "revision_history_latest": hist_rev,
            "exact_match": exact,
            "major_match": major_match,
            "revision_history_row_index": revision_history["latest_entry_row_index"] if revision_history else None,
            **({"note": note} if note else {}),
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

    # A footer whose cells hold only PAGE fields extracts as "| |" - present in
    # structure but carrying no controlled-copy statement or template ID. That
    # is not a pass.
    footer_words = " ".join(
        re.sub(r"[|\s\d]+", " ", f["text"]) for f in footers
    ).strip()
    if not footers:
        result, note = "FAIL", "No footer was found in any w:ftr part."
    elif len(footer_words) < 3:
        result, note = "REVIEW", (
            "A footer exists but carries no readable text - its cells appear to hold only "
            "page-number fields. Confirm by eye whether the controlled-copy statement and "
            "the template ID are present (QA-Q9 / QA-Q10)."
        )
    else:
        result, note = "PASS", None

    checks.append(
        {
            "check": "7b Footer present",
            "result": result,
            "footer_text": [f["text"] for f in footers],
            **({"note": note} if note else {}),
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
    # In a PDF the numbers are baked into the text, so what was parsed is what
    # Word rendered - no reconstruction needed, nothing to corroborate.
    pdf_numbering = validate_numbering(sections, rev_hist)
    pdf_numbering["source"] = "text"
    pdf_numbering["sections_numbered"] = sum(1 for s in sections if s.get("number"))
    header_info.setdefault("revision", header_info["revisions"][0] if header_info["revisions"] else None)
    header_info.setdefault(
        "revision_source", "labelled in the running header" if header_info["revisions"] else None
    )
    header_info.setdefault(
        "revision_confidence", "labelled" if header_info["revisions"] else "not found"
    )
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
        "numbering_validation": pdf_numbering,
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
    sections, tables, figure_numbers, title_paragraph, number_source = build_sections(document)
    title = document_title(document, sections, title_paragraph)
    rev_hist = find_revision_history(tables)
    page_fields = has_page_fields(document)

    # The revision may be unlabelled in a header table cell, so look for it
    # properly once the revision history is available to corroborate it.
    header_info.update(find_header_revision(document, header_text, rev_hist))

    numbering = validate_numbering(sections, rev_hist)
    # Reliability depends on where the numbers actually came from. A number
    # written in the heading text is as trustworthy as the document itself and
    # needs no corroboration; a number rebuilt from the Heading styles does.
    sources = {s["number_source"] for s in sections if s.get("number")}
    numbering["source"] = (
        "text" if sources == {"text"} else (number_source if sources else "none")
    )
    numbering["section_number_sources"] = sorted(s for s in sources if s)
    numbering["sections_numbered"] = sum(1 for s in sections if s.get("number"))
    reconstructed = bool(sources - {"text"})
    numbering_reliable = (not reconstructed) or numbering["reliable"]
    numbering["reliable"] = numbering_reliable

    warnings = []
    if rev_hist is None:
        warnings.append(
            "No revision history table was found. Checks 2, 3 and 4 cannot run and "
            "nothing may be annotated - report this and stop."
        )
    if not headers:
        warnings.append("No page header was found in any w:hdr part.")
    if header_info.get("revision_confidence") == "inferred-corroborated":
        warnings.append(
            f"The page header carries no 'Rev.' label; Rev {header_info['revision']} was read "
            "from a bare header cell and corroborated against the revision history."
        )
    elif header_info.get("revision_confidence") == "uncorroborated":
        warnings.append(
            "A bare value in the page header might be the revision but the revision history "
            "does not corroborate it. Confirm the revision by eye before reporting check 2."
        )
    if reconstructed and numbering["sections_numbered"]:
        if numbering["reliable"]:
            warnings.append(
                f"Section numbers are generated by the Heading styles and were rebuilt from "
                f"the document's numbering definitions; {len(numbering['agreed'])} of "
                f"{numbering['checked']} number/title pairs cited in the revision history "
                "confirm them."
            )
        else:
            warnings.append(
                "Section numbers are generated by the Heading styles and could not be "
                "corroborated against the revision history. Do not quote section numbers "
                "in findings - refer to sections by heading title instead."
            )

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
        "numbering_validation": numbering,
        "tables": [
            {"table_index": i, "rows": table_text_rows(t)} for i, t in enumerate(tables)
        ],
        "citations": harvest_citations(sections),
        "hygiene": {
            "placeholders": check_placeholders(sections),
            "numbering": check_numbering(sections) if numbering_reliable else [],
            "cross_references": check_cross_references(
                sections, len(tables), figure_numbers, numbering_reliable
            ),
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
    print(f"header rev    : {hi.get('revision') or ', '.join(hi['revisions']) or '(none found)'}"
          f"  [{hi.get('revision_confidence')}]")
    rh = data["revision_history"]
    print(f"rev history   : {len(rh['entries'])} entries, latest Rev {rh['latest_rev']}"
          if rh else "rev history   : NOT FOUND")
    nv = data.get("numbering_validation") or {}
    print(f"sections      : {len(data['sections'])} "
          f"({nv.get('sections_numbered', 0)} numbered, source={nv.get('source')}, "
          f"reliable={nv.get('reliable')}, corroborated {len(nv.get('agreed', []))}"
          f"/{nv.get('checked', 0)})")
    print(f"citations     : {len(data['citations'])}")
    for c in data["auto_checks"]:
        print(f"[{c['result']:6}] {c['check']}")
    for w in data["warnings"]:
        print(f"[WARN  ] {w}")
    print(f"wrote {target}")


if __name__ == "__main__":
    main()
