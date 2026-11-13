#!/usr/bin/env python3
"""Compare two revisions of a controlled document, section by section.

Usage:
    python3 diff_revisions.py <doc_a> <doc_b> --out <dir>
    python3 diff_revisions.py <old> <new> --out <dir> --force-order

Which file is the newer revision is decided from the revision in the page
header, not from the file name or the modification time. Pass --force-order to
take the arguments as given (old first) when the headers cannot be trusted.

Writes <dir>/diff.json: sections added / removed / modified, the changed
sentences inside each modified section, a separate numeric-value delta list
(a changed acceptance criterion must never be buried in prose diff noise),
table changes, and a flat list of change statements to compare against the
revision history.
"""
import argparse
import difflib
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from docx_common import normalise, revision_sort_key  # noqa: E402
from extract_doc import extract  # noqa: E402

#: A number with an optional unit. Used to pull quantitative changes out of
#: an otherwise ordinary sentence edit.
NUMBER_RE = re.compile(
    r"(?P<num>\d+(?:[.,]\d+)?)\s*"
    r"(?P<unit>%|percent|°\s*C|degrees?\s*C|deg\s*C|°\s*F|"
    r"mL|L\b|µL|uL|mg|µg|ug|ng|pg|mm|cm|nm|"
    r"hours?|hrs?|h\b|minutes?|mins?|days?|weeks?|months?|years?|"
    r"copies|reads|cycles|x\b)?",
    re.I,
)

SENTENCE_SPLIT_RE = re.compile(r"(?<=[.;:])\s+")

#: Numbers that are identifiers, not quantities. Masked before the numeric
#: delta scan so "VER-0045", "Section 9" and "ISO 13485:2003" are not reported
#: as changed values. A changed standard edition is check 6, not a numeric delta.
NOT_A_QUANTITY_RE = re.compile(
    r"\b[A-Z]{2,5}-\d{3,6}\b"
    r"|\b(?:Section|Sections|Table|Figure|Fig\.?|Appendix|Annex|Part|Clause|Step)\s+\d+(?:\.\d+)*"
    r"|\b(?:ISO|IEC|EN|CLSI)(?:/(?:TR|TS|IEEE))?\s*\d{3,5}(?:-\d+)*(?::\s*\d{4})?"
    r"|\b\d{1,2}\s*CFR\s*(?:Part\s*)?\d{1,4}(?:\.\d+)?"
    r"|\b\d{4}-\d{2}-\d{2}\b"
    r"|\b(?:EU\)?\s*)?20\d{2}/\d{3,4}\b",
    re.I,
)


def sentences(text):
    return [s.strip() for s in SENTENCE_SPLIT_RE.split(text) if s.strip()]


def numbers_in(text):
    """Every quantity in a string, with its unit and a little context.

    Identifiers (document numbers, cross-references, standard citations) are
    blanked out first, keeping the original offsets so context slices still
    line up with the untouched text.
    """
    masked = NOT_A_QUANTITY_RE.sub(lambda m: " " * len(m.group(0)), text)
    out = []
    for m in NUMBER_RE.finditer(masked):
        unit = normalise(m.group("unit") or "")
        out.append(
            {
                "value": m.group("num"),
                "unit": unit,
                "text": normalise(m.group(0)),
                "context": normalise(text[max(0, m.start() - 45):m.end() + 45]),
            }
        )
    return out


def word_diff(old, new):
    """Word-level replacements between two strings."""
    o, n = old.split(), new.split()
    changes = []
    for tag, i1, i2, j1, j2 in difflib.SequenceMatcher(None, o, n).get_opcodes():
        if tag == "equal":
            continue
        changes.append(
            {
                "type": tag,
                "old": " ".join(o[i1:i2]),
                "new": " ".join(n[j1:j2]),
            }
        )
    return changes


def numeric_deltas(old, new, where):
    """Numbers that changed value between two versions of the same sentence."""
    deltas = []
    for change in word_diff(old, new):
        if change["type"] not in ("replace", "delete", "insert"):
            continue
        old_nums = numbers_in(change["old"])
        new_nums = numbers_in(change["new"])
        if not old_nums and not new_nums:
            continue
        if old_nums and new_nums:
            for o, n in zip(old_nums, new_nums):
                if o["text"].lower() != n["text"].lower():
                    deltas.append(
                        {
                            "where": where,
                            "old": o["text"],
                            "new": n["text"],
                            "old_context": o["context"],
                            "new_context": n["context"],
                        }
                    )
            # Unpaired leftovers on either side.
            for o in old_nums[len(new_nums):]:
                deltas.append({"where": where, "old": o["text"], "new": None,
                               "old_context": o["context"], "new_context": None})
            for n in new_nums[len(old_nums):]:
                deltas.append({"where": where, "old": None, "new": n["text"],
                               "old_context": None, "new_context": n["context"]})
        elif old_nums:
            for o in old_nums:
                deltas.append({"where": where, "old": o["text"], "new": None,
                               "old_context": o["context"], "new_context": None})
        else:
            for n in new_nums:
                deltas.append({"where": where, "old": None, "new": n["text"],
                               "old_context": None, "new_context": n["context"]})
    return deltas


def section_key(sec):
    if sec.get("number"):
        return ("num", sec["number"])
    return ("title", normalise(sec.get("title") or sec.get("heading") or "").lower())


def label(sec):
    if sec.get("number"):
        return f"§{sec['number']} {normalise(sec.get('title') or '')}".strip()
    return normalise(sec.get("heading") or "(front matter)")


def diff_paragraphs(old_paras, new_paras, where):
    """Pair up paragraphs within a section and describe what changed."""
    added, removed, modified, deltas = [], [], [], []
    matcher = difflib.SequenceMatcher(None, old_paras, new_paras)
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            continue
        if tag == "insert":
            added.extend(new_paras[j1:j2])
            continue
        if tag == "delete":
            removed.extend(old_paras[i1:i2])
            continue
        # replace: pair the most similar paragraphs so an edited sentence is
        # reported as a modification rather than a delete plus an insert.
        olds, news = list(old_paras[i1:i2]), list(new_paras[j1:j2])
        used = set()
        for o in olds:
            best, best_ratio = None, 0.0
            for idx, n in enumerate(news):
                if idx in used:
                    continue
                ratio = difflib.SequenceMatcher(None, o, n).ratio()
                if ratio > best_ratio:
                    best, best_ratio = idx, ratio
            if best is not None and best_ratio >= 0.55:
                used.add(best)
                n = news[best]
                modified.append({"old": o, "new": n, "similarity": round(best_ratio, 3),
                                 "word_changes": word_diff(o, n)})
                deltas.extend(numeric_deltas(o, n, where))
            else:
                removed.append(o)
        added.extend(n for idx, n in enumerate(news) if idx not in used)
    return added, removed, modified, deltas


def diff_tables(old_tables, new_tables, skip_old, skip_new):
    """Cell-level table comparison, skipping the revision history table."""
    changes = []
    old_map = {t["table_index"]: t["rows"] for t in old_tables if t["table_index"] != skip_old}
    new_map = {t["table_index"]: t["rows"] for t in new_tables if t["table_index"] != skip_new}
    for idx in sorted(set(old_map) | set(new_map)):
        o = old_map.get(idx)
        n = new_map.get(idx)
        if o is None:
            changes.append({"table_index": idx, "type": "table_added", "rows": n})
            continue
        if n is None:
            changes.append({"table_index": idx, "type": "table_removed", "rows": o})
            continue
        cell_changes = []
        for r in range(max(len(o), len(n))):
            orow = o[r] if r < len(o) else []
            nrow = n[r] if r < len(n) else []
            for c in range(max(len(orow), len(nrow))):
                ov = normalise(orow[c]) if c < len(orow) else ""
                nv = normalise(nrow[c]) if c < len(nrow) else ""
                if ov != nv:
                    cell_changes.append({"row": r, "col": c, "old": ov, "new": nv})
        if cell_changes:
            changes.append({"table_index": idx, "type": "table_modified", "cells": cell_changes})
    return changes


def build_change_statements(result):
    """One short line per real difference, for matching against the history."""
    out = []
    meta = result["metadata_changes"]
    for field, change in meta.items():
        out.append(
            {
                "kind": "metadata",
                "where": field,
                "statement": f"{field}: '{change['old']}' -> '{change['new']}'",
            }
        )
    for sec in result["sections_added"]:
        out.append({"kind": "section_added", "where": sec,
                    "statement": f"Section added: {sec}"})
    for sec in result["sections_removed"]:
        out.append({"kind": "section_removed", "where": sec,
                    "statement": f"Section removed: {sec}"})
    for sec in result["sections_modified"]:
        where = sec["section"]
        for d in sec["numeric_deltas"]:
            out.append(
                {
                    "kind": "numeric",
                    "where": where,
                    "statement": f"{where}: {d['old']} -> {d['new']}",
                    "context": d.get("new_context") or d.get("old_context"),
                }
            )
        for m in sec["modified_paragraphs"]:
            words = "; ".join(
                f"'{c['old']}' -> '{c['new']}'"
                for c in m["word_changes"]
                if c["type"] == "replace"
            )
            out.append(
                {
                    "kind": "text",
                    "where": where,
                    "statement": f"{where}: text changed" + (f" ({words})" if words else ""),
                    "context": m["new"],
                }
            )
        for p in sec["added_paragraphs"]:
            out.append({"kind": "text_added", "where": where,
                        "statement": f"{where}: text added", "context": p})
        for p in sec["removed_paragraphs"]:
            out.append({"kind": "text_removed", "where": where,
                        "statement": f"{where}: text removed", "context": p})
    for t in result["table_changes"]:
        out.append({"kind": "table", "where": f"Table index {t['table_index']}",
                    "statement": f"Table index {t['table_index']}: {t['type']}"})
    return out


def header_revision(data):
    revs = data["header_info"]["revisions"]
    return revs[0] if len(revs) == 1 else (revs[0] if revs else None)


def order_documents(a_path, b_path, forced):
    a = extract(a_path)
    b = extract(b_path)
    if forced:
        return (a, a_path), (b, b_path), "forced by --force-order"
    ra, rb = header_revision(a), header_revision(b)
    if ra is None or rb is None:
        return None, None, (
            f"Could not read a revision from the page header of "
            f"{'both files' if ra is None and rb is None else (a_path.name if ra is None else b_path.name)}. "
            "Say so and ask which file is the newer revision, or re-run with --force-order."
        )
    ka, kb = revision_sort_key(ra), revision_sort_key(rb)
    if ka == kb:
        return None, None, (
            f"Both files report Rev {ra} in the page header. "
            "Ask which is the newer revision, or re-run with --force-order."
        )
    if ka > kb:
        return (b, b_path), (a, a_path), f"{b_path.name} is Rev {rb}, {a_path.name} is Rev {ra}"
    return (a, a_path), (b, b_path), f"{a_path.name} is Rev {ra}, {b_path.name} is Rev {rb}"


def run(old, new):
    old_secs = {section_key(s): s for s in old["sections"]}
    new_secs = {section_key(s): s for s in new["sections"]}

    result = {
        "old_file": old["file"],
        "new_file": new["file"],
        "old_revision": header_revision(old),
        "new_revision": header_revision(new),
        "metadata_changes": {},
        "sections_added": [],
        "sections_removed": [],
        "sections_modified": [],
        "table_changes": [],
    }

    # Metadata / header / title
    for field, o, n in (
        ("Document title", old["title"].get("core_properties_title"),
         new["title"].get("core_properties_title")),
        ("Title on page 1", old["title"].get("title_paragraph"),
         new["title"].get("title_paragraph")),
        ("First heading", old["title"].get("first_heading"),
         new["title"].get("first_heading")),
        ("Header title", " / ".join(old["header_info"]["title_candidates"]),
         " / ".join(new["header_info"]["title_candidates"])),
    ):
        if normalise(o or "") != normalise(n or ""):
            result["metadata_changes"][field] = {"old": o, "new": n}

    for key in sorted(set(new_secs) - set(old_secs)):
        result["sections_added"].append(label(new_secs[key]))
    for key in sorted(set(old_secs) - set(new_secs)):
        result["sections_removed"].append(label(old_secs[key]))

    for key in sorted(set(old_secs) & set(new_secs), key=lambda k: str(k)):
        o, n = old_secs[key], new_secs[key]
        where = label(n)
        added, removed, modified, deltas = diff_paragraphs(
            o["paragraphs"], n["paragraphs"], where
        )
        heading_changed = normalise(o["heading"]) != normalise(n["heading"])
        if added or removed or modified or heading_changed:
            result["sections_modified"].append(
                {
                    "section": where,
                    "number": n.get("number"),
                    "heading_old": o["heading"] if heading_changed else None,
                    "heading_new": n["heading"] if heading_changed else None,
                    "added_paragraphs": added,
                    "removed_paragraphs": removed,
                    "modified_paragraphs": modified,
                    "numeric_deltas": deltas,
                }
            )

    skip_old = old["revision_history"]["table_index"] if old.get("revision_history") else None
    skip_new = new["revision_history"]["table_index"] if new.get("revision_history") else None
    result["table_changes"] = diff_tables(
        old["tables"], new["tables"], skip_old, skip_new
    )

    result["change_statements"] = build_change_statements(result)
    result["all_numeric_deltas"] = [
        d for s in result["sections_modified"] for d in s["numeric_deltas"]
    ]
    result["new_revision_history_entries"] = []
    if new.get("revision_history"):
        old_revs = set()
        if old.get("revision_history"):
            old_revs = {e["rev"] for e in old["revision_history"]["entries"]}
        for e in new["revision_history"]["entries"]:
            if e["rev"] not in old_revs:
                result["new_revision_history_entries"].append(e)
    return result


def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("doc_a")
    ap.add_argument("doc_b")
    ap.add_argument("--out", required=True)
    ap.add_argument(
        "--force-order",
        action="store_true",
        help="treat doc_a as the old revision and doc_b as the new one",
    )
    a = ap.parse_args()

    a_path, b_path = Path(a.doc_a), Path(a.doc_b)
    for p in (a_path, b_path):
        if not p.exists():
            sys.exit(f"No such file: {p}")

    old_pair, new_pair, reason = order_documents(a_path, b_path, a.force_order)
    if old_pair is None:
        sys.exit(f"Cannot decide which revision is newer.\n{reason}")

    old, old_path = old_pair
    new, new_path = new_pair
    result = run(old, new)
    result["ordering_reason"] = reason

    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    target = out / "diff.json"
    target.write_text(json.dumps(result, indent=2, ensure_ascii=False))

    print(f"old           : {old_path.name} (Rev {result['old_revision']})")
    print(f"new           : {new_path.name} (Rev {result['new_revision']})")
    print(f"ordering      : {reason}")
    print(f"metadata      : {len(result['metadata_changes'])} changed")
    print(f"sections +/-  : {len(result['sections_added'])} added, "
          f"{len(result['sections_removed'])} removed")
    print(f"sections mod  : {len(result['sections_modified'])}")
    print(f"numeric deltas: {len(result['all_numeric_deltas'])}")
    for d in result["all_numeric_deltas"]:
        print(f"    {d['where']}: {d['old']} -> {d['new']}")
    print(f"statements    : {len(result['change_statements'])}")
    print(f"wrote {target}")


if __name__ == "__main__":
    main()
