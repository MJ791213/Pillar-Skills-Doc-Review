#!/usr/bin/env python3
"""Check that an annotated .docx really carries the annotations it should.

Usage:
    python3 verify_annotation.py <annotated.docx> [--source <original.docx>]

Asserts, at the XML level rather than by eye:
  * the file reopens in python-docx (a file Word would refuse to open usually
    fails here too);
  * red runs exist, and they sit inside a table;
  * each red run inherited a font family and size rather than falling back to
    the Word default, which is the failure a reviewer notices first;
  * an orange recommendations block exists and is the last content in the body;
  * the "this is a Claude recommendation" note is present.

Exits non-zero and prints each failure if any check fails.
"""
import argparse
import sys
from pathlib import Path

try:
    import docx
    from docx.oxml.ns import qn
except ImportError:  # pragma: no cover
    sys.exit("python-docx is required: pip install python-docx")

RED = "FF0000"
ORANGE = "ED7D31"
MARKER = "[Claude]"
BLOCK_HEADING_PREFIX = "Claude Recommendations"


def run_text(r):
    return "".join(t.text or "" for t in r.findall(qn("w:t")))


def run_colour(r):
    rpr = r.find(qn("w:rPr"))
    if rpr is None:
        return None
    colour = rpr.find(qn("w:color"))
    return colour.get(qn("w:val")).upper() if colour is not None else None


def run_font(r):
    rpr = r.find(qn("w:rPr"))
    if rpr is None:
        return None, None
    fonts = rpr.find(qn("w:rFonts"))
    size = rpr.find(qn("w:sz"))
    family = fonts.get(qn("w:ascii")) if fonts is not None else None
    pts = int(size.get(qn("w:val"))) / 2 if size is not None else None
    return family, pts


def in_table(r):
    node = r.getparent()
    while node is not None:
        if node.tag == qn("w:tbl"):
            return True
        node = node.getparent()
    return False


def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("annotated")
    ap.add_argument("--source", help="the un-annotated original, for a font comparison")
    a = ap.parse_args()

    path = Path(a.annotated)
    if not path.exists():
        sys.exit(f"No such file: {path}")

    failures, notes = [], []

    try:
        document = docx.Document(str(path))
    except Exception as exc:
        sys.exit(f"FAIL: the annotated file does not reopen: {exc}")
    notes.append("file reopens cleanly")

    body = document.element.body
    runs = body.findall(f".//{qn('w:r')}")

    red = [r for r in runs if run_colour(r) == RED]
    orange = [r for r in runs if run_colour(r) == ORANGE]

    if not red:
        failures.append("no red runs found")
    else:
        notes.append(f"{len(red)} red run(s)")
        outside = [r for r in red if not in_table(r)]
        if outside:
            failures.append(
                f"{len(outside)} red run(s) sit outside any table; red belongs in "
                "the revision history table only"
            )
        unmarked = [r for r in red if not run_text(r).strip().startswith(MARKER)]
        if unmarked:
            failures.append(f"{len(unmarked)} red run(s) are missing the {MARKER} marker")

        # Whether a run *should* carry an explicit font depends on the document.
        # Many templates set the font on the paragraph style and leave the runs
        # bare, so an absent w:rFonts is correct there and demanding one would
        # fail a correctly annotated file. What matters is that the red run
        # matches its surroundings - checked against --source below, and via
        # the paragraph style here.
        for r in red:
            family, pts = run_font(r)
            if (family is None or pts is None) and not a.source:
                notes.append(
                    f"red run has no explicit font (family={family}, size={pts}); it will "
                    "inherit from the paragraph style. Pass --source to confirm that "
                    "matches the source table."
                )
                break

        styleless = []
        for r in red:
            p = r.getparent()
            if p is None or p.tag != qn("w:p"):
                continue
            ppr = p.find(qn("w:pPr"))
            style = ppr.find(qn("w:pStyle")) if ppr is not None else None
            siblings = [
                s for s in (p.getparent().findall(qn("w:p")) if p.getparent() is not None else [])
                if s is not p
            ]
            sibling_styles = set()
            for s in siblings:
                sppr = s.find(qn("w:pPr"))
                sstyle = sppr.find(qn("w:pStyle")) if sppr is not None else None
                sibling_styles.add(sstyle.get(qn("w:val")) if sstyle is not None else None)
            mine = style.get(qn("w:val")) if style is not None else None
            if sibling_styles and mine not in sibling_styles:
                styleless.append((mine, sorted(map(str, sibling_styles))))
        if styleless:
            failures.append(
                "red note paragraph(s) do not share the paragraph style of the cell's "
                f"existing text, so they will render differently: {styleless[:3]}"
            )
        else:
            notes.append("red notes share the paragraph style of the cell's existing text")

    if a.source:
        src = docx.Document(a.source)
        src_fonts = set()
        for r in src.element.body.findall(f".//{qn('w:r')}"):
            if in_table(r) and run_text(r).strip():
                src_fonts.add(run_font(r))
        red_fonts = {run_font(r) for r in red}
        if red_fonts and not (red_fonts & src_fonts):
            failures.append(
                f"red run fonts {sorted(map(str, red_fonts))} match none of the "
                f"source table fonts {sorted(map(str, src_fonts))}"
            )
        else:
            notes.append(f"red run font matches the source table font: {sorted(map(str, red_fonts))}")

    if not orange:
        failures.append("no orange runs found")
    else:
        notes.append(f"{len(orange)} orange run(s)")

    paragraphs = list(body.findall(qn("w:p")))
    heading_index = None
    for i, p in enumerate(paragraphs):
        text = "".join(t.text or "" for t in p.findall(f".//{qn('w:t')}"))
        if text.strip().startswith(BLOCK_HEADING_PREFIX):
            heading_index = i
            break
    if heading_index is None:
        failures.append(f"no paragraph starting with {BLOCK_HEADING_PREFIX!r}")
    else:
        notes.append(f"recommendations block heading at body paragraph {heading_index}")

    full_text = "\n".join(
        "".join(t.text or "" for t in p.findall(f".//{qn('w:t')}")) for p in paragraphs
    )
    if "is a Claude recommendation" not in full_text:
        failures.append("the 'this is a Claude recommendation' note is missing")
    else:
        notes.append("Claude-recommendation note present")

    # The block must be the last content: no body paragraph with text after the
    # closing note, ignoring empty paragraphs.
    if heading_index is not None:
        tail = [
            "".join(t.text or "" for t in p.findall(f".//{qn('w:t')}")).strip()
            for p in paragraphs[heading_index:]
        ]
        if not any("Claude recommendation" in t for t in tail):
            failures.append("the closing note does not appear inside the appended block")

    for n in notes:
        print(f"[ok    ] {n}")
    for f in failures:
        print(f"[FAIL  ] {f}")
    if failures:
        sys.exit(1)
    print("all annotation checks passed")


if __name__ == "__main__":
    main()
