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
    # Document and part numbers: DES-0028, OL-0116, IDX-PI-1015-96, HDA-LC-01
    r"\b[A-Z]{2,6}(?:-[A-Z0-9]{1,6}){1,5}\b"
    r"|\b(?:Section|Sections|Table|Figure|Fig\.?|Appendix|Annex|Part|Clause|Step|Steps)\s+\d+(?:\.\d+)*"
    r"|\b(?:ISO|IEC|EN|CLSI)(?:/(?:TR|TS|IEEE))?\s*\d{3,5}(?:-\d+)*(?::\s*\d{4})?"
    r"|\b\d{1,2}\s*CFR\s*(?:Part\s*)?\d{1,4}(?:\.\d+)?"
    r"|\b\d{4}-\d{2}-\d{2}\b"
    r"|\b(?:EU\)?\s*)?20\d{2}/\d{3,4}\b"
    # Step / section identifiers: three or more dotted components is never a
    # quantity, and a dotted number opening a block is a step label.
    r"|\b\d+\.\d+\.\d+(?:\.\d+)*\b"
    r"|^\s*\d+(?:\.\d+)+",
    re.I | re.M,
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
    """Quantities that changed value between two versions of the same sentence.

    Returns (changed, one_sided). Only a value present on *both* sides and
    different is a changed value - the high-risk edit a reviewer must see. A
    number that appears only in added text is just part of that added text, and
    reporting it as a "delta" with old=None buries the real changes in noise.
    """
    changed, one_sided = [], []
    for change in word_diff(old, new):
        if change["type"] not in ("replace", "delete", "insert"):
            continue
        old_nums = numbers_in(change["old"])
        new_nums = numbers_in(change["new"])
        if not old_nums and not new_nums:
            continue
        for o, n in zip(old_nums, new_nums):
            if o["text"].lower() != n["text"].lower():
                changed.append(
                    {
                        "where": where,
                        "old": o["text"],
                        "new": n["text"],
                        "old_context": o["context"],
                        "new_context": n["context"],
                    }
                )
        for o in old_nums[len(new_nums):]:
            one_sided.append({"where": where, "old": o["text"], "new": None,
                              "old_context": o["context"], "new_context": None})
        for n in new_nums[len(old_nums):]:
            one_sided.append({"where": where, "old": None, "new": n["text"],
                              "old_context": None, "new_context": n["context"]})
    return changed, one_sided


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
    added, removed, modified, deltas, one_sided = [], [], [], [], []
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
                ch, os_ = numeric_deltas(o, n, where)
                deltas.extend(ch)
                one_sided.extend(os_)
            else:
                removed.append(o)
        added.extend(n for idx, n in enumerate(news) if idx not in used)
    return added, removed, modified, deltas, one_sided


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


def _squash(s):
    return re.sub(r"[^a-z0-9]", "", (s or "").lower())


def cross_check_numbering(old, new):
    """Explain numbering gaps and stale citations using *both* revisions.

    A gap in the heading numbers, or a revision history entry citing a section
    that no longer exists, is not a defect on its own: the document may simply
    have had a section added or removed in this revision, which renumbers
    everything after it. Judging either from the new revision alone invents
    findings. With both revisions in hand each one can be explained, and what
    is left is the question actually worth asking - whether the add or the
    delete is recorded in the revision history.
    """
    from extract_doc import (
        CITED_SECTION_RE,
        _NOT_A_TITLE_START,
        check_numbering,
        find_section_by_title,
    )

    old_by_num = {s["number"]: s for s in old["sections"] if s.get("number")}
    new_by_num = {s["number"]: s for s in new["sections"] if s.get("number")}

    def in_old(title):
        return find_section_by_title(old["sections"], title)

    def in_new(title):
        return find_section_by_title(new["sections"], title)

    def missing_numbers(data):
        out = set()
        for issue in check_numbering(data["sections"]):
            if issue["type"] == "heading_number_gap":
                out.update(issue["missing"])
        return out

    new_missing = missing_numbers(new)
    old_missing = missing_numbers(old)

    gaps = []
    for number in sorted(new_missing):
        if number in old_missing:
            gaps.append({
                "number": number, "status": "pre-existing",
                "note": "The same gap exists in the previous revision, so this change did "
                        "not introduce it. Not a finding against this revision.",
            })
            continue
        was = old_by_num.get(number)
        if was is None:
            gaps.append({
                "number": number, "status": "unexplained",
                "note": "No section carried this number in the previous revision either, "
                        "so the gap is not explained by an addition or a removal.",
            })
            continue
        moved_to = in_new(was.get("title"))
        if moved_to is not None:
            gaps.append({
                "number": number, "status": "renumbered", "title": was.get("title"),
                "now": moved_to.get("number") or "unnumbered",
                "note": "The section still exists under a different number, so the gap is "
                        "the result of renumbering. Check the revision history describes "
                        "whatever insertion or removal caused it.",
            })
        else:
            gaps.append({
                "number": number, "status": "section_removed", "title": was.get("title"),
                "note": "This section existed in the previous revision and is gone from "
                        "this one, which explains the gap. The question is whether the "
                        "removal is recorded in the revision history, and whether the "
                        "remaining sections should have been renumbered to close it.",
            })

    citations = []
    rh = new.get("revision_history")
    if rh:
        for entry in rh["entries"]:
            for m in CITED_SECTION_RE.finditer(entry["description"] or ""):
                number, title = m.group(1), (m.group(2) or "").strip()
                if not title or len(title) < 4:
                    continue
                if title.split()[0].lower() in _NOT_A_TITLE_START:
                    continue
                here = new_by_num.get(number)
                if here is not None:
                    if in_new(title) is here and here is not None:
                        continue  # cites the right number for the right section
                now = in_new(title)
                record = {
                    "rev": entry["rev"], "row_index": entry["row_index"],
                    "cited_as": number, "cited_title": title,
                }
                if now is not None:
                    record.update({
                        "status": "moved",
                        "now": now.get("number") or "unnumbered",
                        "note": "The section still exists under a different number. The "
                                "entry was accurate when written; it is the numbering that "
                                "has moved since. Report it so the reader is not sent to "
                                "the wrong place, not as an error in the entry.",
                    })
                elif in_old(title) is not None:
                    record.update({
                        "status": "section_removed",
                        "note": "The section this entry cites existed in the previous "
                                "revision and is not in this one. Check the removal is "
                                "recorded in the revision history.",
                    })
                else:
                    record.update({
                        "status": "unresolved",
                        "note": "No section with this number or title exists in either "
                                "revision, so the citation looks wrong as written.",
                    })
                citations.append(record)

    return {
        "numbering_gaps": gaps,
        "revision_history_citations": citations,
        "note": "Gaps and citations judged against both revisions. A gap marked "
                "'pre-existing', 'renumbered' or 'section_removed' is explained by the "
                "document's history and is not by itself an error.",
    }


def header_revision(data):
    info = data["header_info"]
    if info.get("revision"):
        return info["revision"]
    revs = info.get("revisions") or []
    return revs[0] if len(revs) == 1 else None


def history_revision(data):
    rh = data.get("revision_history")
    return rh.get("latest_rev") if rh else None


def order_documents(a_path, b_path, forced):
    a = extract(a_path)
    b = extract(b_path)
    if forced:
        return (a, a_path), (b, b_path), "forced by --force-order"

    def decide(ra, rb, basis):
        ka, kb = revision_sort_key(ra), revision_sort_key(rb)
        if ka == kb:
            return None
        if ka > kb:
            return (b, b_path), (a, a_path), f"{b_path.name} is {basis} {rb}, {a_path.name} is {basis} {ra}"
        return (a, a_path), (b, b_path), f"{a_path.name} is {basis} {ra}, {b_path.name} is {basis} {rb}"

    ra, rb = header_revision(a), header_revision(b)
    if ra is not None and rb is not None:
        decided = decide(ra, rb, "Rev")
        if decided:
            return decided
        # Same header revision is not necessarily a tie: with a major-only
        # header convention, Rev 8 and Rev 8 can be 8.02 and 8.04 underneath.
        ha, hb = history_revision(a), history_revision(b)
        if ha is not None and hb is not None:
            decided = decide(ha, hb, "latest revision history entry")
            if decided:
                old, new, reason = decided
                return old, new, (
                    f"both headers read Rev {ra}, so ordered by the revision history: {reason}"
                )
        return None, None, (
            f"Both files report Rev {ra} in the page header and their revision histories do "
            "not distinguish them. Ask which is the newer revision, or re-run with "
            "--force-order."
        )

    # The header revision could not be read from one or both files. The revision
    # history is the next most reliable signal, so use it and say so.
    ha, hb = history_revision(a), history_revision(b)
    missing = (
        "both files" if ra is None and rb is None
        else (a_path.name if ra is None else b_path.name)
    )
    if ha is not None and hb is not None:
        decided = decide(ha, hb, "latest revision history entry")
        if decided:
            old, new, reason = decided
            return old, new, (
                f"no revision could be read from the page header of {missing}, so ordered by "
                f"the revision history instead: {reason}"
            )
        if len(a["revision_history"]["entries"]) != len(b["revision_history"]["entries"]):
            decided = decide(
                len(a["revision_history"]["entries"]),
                len(b["revision_history"]["entries"]),
                "revision history entry count",
            )
            if decided:
                old, new, reason = decided
                return old, new, (
                    f"no revision could be read from the page header of {missing}, and both "
                    f"revision histories end at the same revision, so ordered by entry "
                    f"count: {reason}"
                )

    return None, None, (
        f"Could not read a revision from the page header of {missing}, and the revision "
        "histories do not distinguish the two files. Say so and ask which file is the newer "
        "revision, or re-run with --force-order."
    )


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
        added, removed, modified, deltas, one_sided = diff_paragraphs(
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
                    "one_sided_values": one_sided,
                }
            )

    skip_old = old["revision_history"]["table_index"] if old.get("revision_history") else None
    skip_new = new["revision_history"]["table_index"] if new.get("revision_history") else None
    result["table_changes"] = diff_tables(
        old["tables"], new["tables"], skip_old, skip_new
    )

    result["numbering_cross_check"] = cross_check_numbering(old, new)
    result["change_statements"] = build_change_statements(result)
    result["all_numeric_deltas"] = [
        d for s in result["sections_modified"] for d in s["numeric_deltas"]
    ]
    # Numbers that appear on only one side. Lower signal than a changed value,
    # kept separate so they cannot crowd it out.
    result["all_one_sided_values"] = [
        d for s in result["sections_modified"] for d in s.get("one_sided_values", [])
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
    print(f"changed values: {len(result['all_numeric_deltas'])}")
    for d in result["all_numeric_deltas"]:
        print(f"    {d['where']}: {d['old']} -> {d['new']}")
    print(f"one-sided vals: {len(result['all_one_sided_values'])} (numbers only in added/removed text)")
    cc = result["numbering_cross_check"]
    if cc["numbering_gaps"]:
        print(f"numbering gaps: {len(cc['numbering_gaps'])}")
        for g in cc["numbering_gaps"]:
            print(f"    {g['number']}: {g['status']}"
                  + (f" ({g.get('title')})" if g.get("title") else ""))
    if cc["revision_history_citations"]:
        print(f"stale cites   : {len(cc['revision_history_citations'])}")
        for c in cc["revision_history_citations"]:
            print(f"    Rev {c['rev']} cites {c['cited_as']} \"{c['cited_title'][:40]}\""
                  f" -> {c['status']}" + (f" (now {c['now']})" if c.get("now") else ""))
    print(f"statements    : {len(result['change_statements'])}")
    print(f"wrote {target}")


if __name__ == "__main__":
    main()
