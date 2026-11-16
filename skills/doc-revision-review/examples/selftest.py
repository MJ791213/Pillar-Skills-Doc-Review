#!/usr/bin/env python3
"""Self-test: assert the scripts still find what they are supposed to find.

Usage:
    python3 examples/selftest.py              # run from the skill directory
    python3 examples/selftest.py --keep       # keep the generated files

Builds the Rev 1 / Rev 2 fixture pair, runs the whole pipeline, and asserts the
baseline recorded in examples/expected-findings.md. Exits 0 only if every
assertion holds. Run it after changing any script.

This tests the scripts, not Claude's judgement: it cannot tell you whether a
review reads well, only whether the mechanics still work.
"""
import argparse
import hashlib
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

SKILL_DIR = Path(__file__).resolve().parent.parent
SCRIPTS = SKILL_DIR / "scripts"

results = []


def check(name, condition, detail=""):
    results.append((bool(condition), name, detail))
    return bool(condition)


def equal(name, got, want):
    ok = got == want
    results.append((ok, name, "" if ok else f"got {got!r}, want {want!r}"))
    return ok


def run(*args, expect_rc=0):
    proc = subprocess.run(
        [sys.executable, *[str(a) for a in args]],
        capture_output=True, text=True,
    )
    if proc.returncode != expect_rc:
        print(f"--- command failed: {' '.join(str(a) for a in args)}")
        print(proc.stdout)
        print(proc.stderr)
    check(
        f"exit code {expect_rc} from {Path(args[0]).name}",
        proc.returncode == expect_rc,
        f"got {proc.returncode}",
    )
    return proc


def md5(path):
    return hashlib.md5(Path(path).read_bytes()).hexdigest()


def body_text(path):
    """Paragraph and table-row text of a .docx, for round-trip comparison."""
    import docx
    from docx.oxml.ns import qn

    d = docx.Document(str(path))
    out = []
    for el in d.element.body.iterchildren():
        if el.tag == qn("w:p"):
            t = "".join(x.text or "" for x in el.findall(f".//{qn('w:t')}")).strip()
            if t:
                out.append(("P", t))
        elif el.tag == qn("w:tbl"):
            for row in el.findall(f".//{qn('w:tr')}"):
                cells = [
                    "".join(x.text or "" for x in c.findall(f".//{qn('w:t')}")).strip()
                    for c in row.findall(qn("w:tc"))
                ]
                out.append(("R", " | ".join(cells)))
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--keep", action="store_true", help="keep the generated files")
    a = ap.parse_args()

    work = Path(tempfile.mkdtemp(prefix="doc-revision-review-selftest-"))
    print(f"working in {work}\n")

    try:
        # ---------------------------------------------------------------
        # Packaging limits. `claude plugin validate` does not enforce these,
        # so they only surface when the claude.ai uploader rejects the zip.
        # ---------------------------------------------------------------
        import re as _re

        skill_md = (SKILL_DIR / "SKILL.md").read_text()
        front = skill_md.split("---")[1]
        m = _re.search(r"^description:\s*(.*?)(?=^\w+:|\Z)", front, _re.S | _re.M)
        check("SKILL.md has a description", m is not None)
        if m:
            desc = " ".join(m.group(1).split())
            check(
                f"description is within the 1024-character upload limit ({len(desc)})",
                len(desc) <= 1024,
                f"{len(desc)} characters, {len(desc) - 1024} over",
            )
            missing = [
                p for p in ("document review", "review this document", "revision history",
                            "compare revisions", "QA review")
                if p not in desc
            ]
            check("description keeps its trigger phrases", not missing, f"missing {missing}")
        name = _re.search(r"^name:\s*(\S+)", front, _re.M)
        equal("skill name matches its directory name",
              name.group(1) if name else None, SKILL_DIR.name)
        # ---------------------------------------------------------------
        # Fixtures
        # ---------------------------------------------------------------
        run(SKILL_DIR / "examples" / "make_fixtures.py", "--out", work / "fx")
        rev1 = work / "fx" / "DES-9999_Rev1.docx"
        rev2 = work / "fx" / "DES-9999_Rev2.docx"
        check("fixtures exist", rev1.exists() and rev2.exists())
        rev2_before = md5(rev2)

        # ---------------------------------------------------------------
        # Rev 1 is clean: no false positives
        # ---------------------------------------------------------------
        run(SCRIPTS / "extract_doc.py", rev1, "--out", work / "r1")
        e1 = json.loads((work / "r1" / "extract.json").read_text())
        bad = [c["check"] for c in e1["auto_checks"] if c["result"] != "PASS"]
        check("Rev 1 passes every automated check (no false positives)", not bad, f"failing: {bad}")
        equal("Rev 1 has no warnings", e1["warnings"], [])
        equal("Rev 1 header revision", e1["header_info"]["revisions"], ["1"])
        equal("Rev 1 hygiene is clean", (
            len(e1["hygiene"]["placeholders"]),
            len(e1["hygiene"]["numbering"]),
            len(e1["hygiene"]["cross_references"]),
        ), (0, 0, 0))

        # ---------------------------------------------------------------
        # Rev 2 defects
        # ---------------------------------------------------------------
        run(SCRIPTS / "extract_doc.py", rev2, "--out", work / "r2")
        e2 = json.loads((work / "r2" / "extract.json").read_text())
        checks = {c["check"]: c["result"] for c in e2["auto_checks"]}
        equal("defect 1: title vs header FAILs", checks.get("1 Title vs header"), "FAIL")
        equal("defect 2: header rev vs history FAILs",
              checks.get("2 Header revision vs revision history"), "FAIL")
        equal("header reads Rev 2", e2["header_info"]["revisions"], ["2"])
        equal("history latest reads Rev 3", e2["revision_history"]["latest_rev"], "3")
        equal("latest history entry is row 2", e2["revision_history"]["latest_entry_row_index"], 2)
        equal("footer and page fields still PASS",
              (checks.get("7a Page x of y fields present"), checks.get("7b Footer present")),
              ("PASS", "PASS"))

        numbers = [s["number"] for s in e2["sections"] if s["number"]]
        equal("section tree is exact",
              numbers,
              ["1", "2", "3", "4", "4.1", "4.2", "4.4", "5", "5.1", "5.2", "6", "7", "8"])
        check("no phantom 'Section 21' from the 21 CFR reference line",
              "21" not in numbers, f"sections: {numbers}")
        check("title paragraph is not treated as a section",
              not any("Specification" in (s.get("title") or "") for s in e2["sections"] if s["number"]))

        equal("defect 7a: one TBD placeholder in section 6",
              [(p["section"], p["match"]) for p in e2["hygiene"]["placeholders"]],
              [("6", "TBD")])
        numbering = e2["hygiene"]["numbering"]
        equal("defect 7b: exactly one numbering gap", len(numbering), 1)
        if numbering:
            equal("the gap is 4.3", (numbering[0]["parent"], numbering[0]["missing"]), ("4", ["4.3"]))
        equal("defect 7c: one broken reference to Section 9",
              [(x["type"], x["target"]) for x in e2["hygiene"]["cross_references"]],
              [("broken_section_reference", "Section 9")])
        equal("acronym noise is filtered (no CLSI, no VER)",
              sorted(x["acronym"] for x in e2["hygiene"]["acronyms"]["undefined"]),
              ["DAISY", "NIPT"])

        cites = {(c["text"], c["section"]) for c in e2["citations"]}
        for want in (("ISO 13485:2003", "2"), ("ISO 15223-1:2016", "2"), ("21 CFR 820", "2")):
            check(f"defect 6: {want[0]} found in section {want[1]}", want in cites)

        # ---------------------------------------------------------------
        # Diff, with the arguments deliberately the wrong way round
        # ---------------------------------------------------------------
        run(SCRIPTS / "diff_revisions.py", rev2, rev1, "--out", work / "d")
        d = json.loads((work / "d" / "diff.json").read_text())
        equal("newer revision identified from the header, not argument order",
              (d["old_revision"], d["new_revision"]), ("1", "2"))
        check("old file is Rev 1", d["old_file"].endswith("Rev1.docx"), d["old_file"])

        deltas = {(x["where"], x["old"], x["new"]) for x in d["all_numeric_deltas"]}
        equal("defect 4: exactly three numeric deltas", len(deltas), 3)
        check("the 95% -> 98% acceptance criterion change is caught",
              ("§5.2 Acceptance Criteria", "95%", "98%") in deltas, f"got {sorted(deltas)}")
        check("the storage range change is caught",
              ("§3 Storage Conditions", "15", "2") in deltas, f"got {sorted(deltas)}")
        check("VER-0045 is not read as a quantity",
              not any("0045" in str(x) for t in deltas for x in t), f"got {sorted(deltas)}")
        check("the 'Section 9' cross-reference is not read as a quantity",
              not any(t[0].startswith("§1 ") for t in deltas), f"got {sorted(deltas)}")

        statements = [s["statement"] for s in d["change_statements"]]
        check("renumbering 4.3 -> 4.4 is reported",
              any("4.4" in s for s in statements) and any("4.3" in s for s in statements))
        check("the title change is reported", any("Document title" in s for s in statements))
        check("no phantom section added for the title paragraph",
              not any("Section added: NIPT" in s for s in statements), f"got {statements}")

        # ---------------------------------------------------------------
        # Annotate
        # ---------------------------------------------------------------
        annotated = work / "DES-9999_Rev2_CLAUDE-REVIEW.docx"
        run(SCRIPTS / "annotate_docx.py", rev2,
            "--findings", SKILL_DIR / "assets" / "findings-template.json",
            "--out", annotated)
        check("annotated file written", annotated.exists())
        equal("source file untouched by annotation", md5(rev2), rev2_before)

        run(SKILL_DIR / "examples" / "verify_annotation.py", annotated, "--source", rev2)

        import docx
        doc = docx.Document(str(annotated))
        cell = doc.tables[0].rows[2].cells[3]
        red = [p for p in cell.paragraphs
               if any(r.font.color and r.font.color.rgb and str(r.font.color.rgb) == "FF0000"
                      for r in p.runs)]
        equal("six red notes in the latest revision history row", len(red), 6)
        fonts = {(r.font.name, r.font.size.pt if r.font.size else None)
                 for p in red for r in p.runs}
        equal("red notes inherit the table font (Arial 10pt)", fonts, {("Arial", 10.0)})
        check("every red note carries the [Claude] marker",
              all(p.text.strip().startswith("[Claude]") for p in red))

        full = "\n".join(p.text for p in doc.paragraphs)
        check("orange block is labelled as not approved content",
              "not part of the approved document" in full)
        check("closing note states it is a Claude recommendation",
              "is a Claude recommendation" in full)

        # ---------------------------------------------------------------
        # Strip round-trip
        # ---------------------------------------------------------------
        stripped = work / "stripped.docx"
        run(SCRIPTS / "annotate_docx.py", annotated, "--strip-annotations", "--out", stripped)
        equal("stripping restores the original content exactly",
              body_text(stripped), body_text(rev2))

        # ---------------------------------------------------------------
        # Refusals
        # ---------------------------------------------------------------
        run(SCRIPTS / "diff_revisions.py", rev2, rev2, "--out", work / "same", expect_rc=1)
        run(SCRIPTS / "annotate_docx.py", rev2, "--findings", work / "nope.json",
            "--out", work / "x.docx", expect_rc=1)
        run(SCRIPTS / "annotate_docx.py", rev2,
            "--findings", SKILL_DIR / "assets" / "findings-template.json",
            "--out", rev2, expect_rc=1)
        equal("source still untouched after the refused in-place write", md5(rev2), rev2_before)

        # ---------------------------------------------------------------
        # Second pair: shaped like a real controlled-document template.
        # These are the traits the first pair cannot exercise.
        # ---------------------------------------------------------------
        tpl_old = work / "fx" / "TP-9999_Rev7.docx"
        tpl_new = work / "fx" / "TP-9999_Rev8.docx"
        check("template-shaped fixtures exist", tpl_old.exists() and tpl_new.exists())

        run(SCRIPTS / "extract_doc.py", tpl_new, "--out", work / "t2")
        t = json.loads((work / "t2" / "extract.json").read_text())
        hi = t["header_info"]
        tchecks = {c["check"]: c for c in t["auto_checks"]}

        # Revision as a bare, unlabelled number in a header table cell.
        equal("tpl: unlabelled header revision found", hi.get("revision"), "8")
        equal("tpl: and corroborated, not guessed",
              hi.get("revision_confidence"), "inferred-corroborated")
        c2 = tchecks.get("2 Header revision vs revision history", {})
        equal("tpl: check 2 PASSes under the major-only convention", c2.get("result"), "PASS")
        equal("tpl: header 8 is not an exact match for 8.04", c2.get("exact_match"), False)
        equal("tpl: but the majors agree", c2.get("major_match"), True)

        # Heading numbers generated by the styles, absent from the text.
        nv = t["numbering_validation"]
        equal("tpl: numbering came from the styles, not the text", nv["source"], "resolved")
        equal("tpl: numbering corroborated by the document's own citations",
              nv["reliable"], True)
        numbers = [s["number"] for s in t["sections"] if s["number"]]
        equal("tpl: reconstructed section numbers are exact",
              numbers, ["1", "2", "3", "4", "4.1", "4.2", "4.3", "5", "6"])
        appendix = [s for s in t["sections"] if "APPENDIX" in s["heading"].upper()]
        check("tpl: the appendix heading is unnumbered",
              appendix and appendix[0]["number"] is None)
        after = t["sections"][t["sections"].index(appendix[0]) + 1:] if appendix else []
        check("tpl: content under the appendix stays unnumbered",
              all(s["number"] is None for s in after),
              f"got {[(s['number'], s['heading'][:28]) for s in after]}")
        check("tpl: a numbered step typed in a body paragraph is not a section",
              not any(str(s.get("number") or "").startswith("4.3.") for s in t["sections"])
              and not any("escalate to QA" in (s.get("title") or "") for s in t["sections"]))

        # A stale citation is a document finding, not a numbering failure.
        stale = nv["stale_citations"]
        equal("tpl: the stale citation is classified as stale, not a mismatch",
              (len(stale), len(nv["mismatches"])), (1, 0))
        if stale:
            equal("tpl: and it reports where the section lives now",
                  (stale[0]["cited_as"], stale[0]["now"]), ("5.3", "4.3"))

        # No invented findings.
        equal("tpl: no false broken cross-references", t["hygiene"]["cross_references"], [])
        equal("tpl: definitions table read, so no false undefined acronyms",
              [x["acronym"] for x in t["hygiene"]["acronyms"]["undefined"]], [])
        equal("tpl: title check is REVIEW, not a false FAIL",
              tchecks.get("1 Title vs header", {}).get("result"), "REVIEW")
        equal("tpl: fields-only footer is REVIEW, not a false PASS",
              tchecks.get("7b Footer present", {}).get("result"), "REVIEW")
        equal("tpl: the Name column is read as the author",
              t["revision_history"]["entries"][0]["author"], "A. Reviewer")

        # Ordering and the diff.
        run(SCRIPTS / "diff_revisions.py", tpl_new, tpl_old, "--out", work / "t2d")
        td = json.loads((work / "t2d" / "diff.json").read_text())
        equal("tpl: ordered 7 -> 8 from the unlabelled header revision",
              (td["old_revision"], td["new_revision"]), ("7", "8"))
        tdeltas = {(x["where"], x["old"], x["new"]) for x in td["all_numeric_deltas"]}
        check("tpl: the hold-time change 15 -> 30 minutes is caught",
              any(o == "15" and n == "30" for _, o, n in tdeltas), f"got {sorted(tdeltas)}")

        # Annotating a style-driven document must keep the cell's paragraph style.
        tpl_findings = work / "tpl_findings.json"
        tpl_findings.write_text(json.dumps({
            "document": {"doc_id": "TP-9999", "revision": "8.04"},
            "revision_history_table_index": t["revision_history"]["table_index"],
            "revision_history_notes": [
                {"row_index": t["revision_history"]["latest_entry_row_index"],
                 "col_index": t["revision_history"]["columns"].get("description"),
                 "text": "Missing entry: the incoming inspection hold time changed from "
                         "15 to 30 minutes."}
            ],
            "summary": [{"number": 4, "check": "Revision history complete",
                         "result": "FAIL", "evidence": "hold time 15 -> 30 undocumented"}],
            "recommendations": [{"severity": "Major", "qa_q": "QA-Q6", "location": "§4.1",
                                 "finding": "Hold time changed without an entry.",
                                 "suggested_action": "Add a revision history entry."}],
        }))
        tpl_annotated = work / "TP-9999_Rev8_CLAUDE-REVIEW.docx"
        tpl_before = md5(tpl_new)
        run(SCRIPTS / "annotate_docx.py", tpl_new,
            "--findings", tpl_findings, "--out", tpl_annotated)
        run(SKILL_DIR / "examples" / "verify_annotation.py", tpl_annotated, "--source", tpl_new)
        equal("tpl: source untouched", md5(tpl_new), tpl_before)

    finally:
        if a.keep:
            print(f"\nkept {work}")
        else:
            shutil.rmtree(work, ignore_errors=True)

    # -------------------------------------------------------------------
    passed = sum(1 for ok, _, _ in results if ok)
    failed = [(n, d) for ok, n, d in results if not ok]
    print()
    for ok, name, detail in results:
        if not ok:
            print(f"[FAIL] {name}" + (f"  -- {detail}" if detail else ""))
    print(f"\n{passed}/{len(results)} assertions passed")
    if failed:
        print(f"{len(failed)} FAILED")
        return 1
    print("self-test OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())
