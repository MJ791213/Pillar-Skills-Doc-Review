# Expected findings — DES-9999 Rev 1 → Rev 2 fixture

The regression baseline for the pair produced by `make_fixtures.py`. Rev 2
carries one deliberate defect per check. A correct review finds all of them and
invents nothing beyond them.

Regenerate and re-run the pipeline:

```bash
python3 examples/make_fixtures.py --out /tmp/fx
```

```bash
python3 scripts/extract_doc.py /tmp/fx/DES-9999_Rev2.docx --out /tmp/rev2
```

```bash
python3 scripts/diff_revisions.py /tmp/fx/DES-9999_Rev2.docx /tmp/fx/DES-9999_Rev1.docx --out /tmp/diff
```

```bash
python3 scripts/annotate_docx.py /tmp/fx/DES-9999_Rev2.docx --findings assets/findings-template.json --out /tmp/out/full.docx
```

```bash
python3 examples/verify_annotation.py /tmp/out/full.docx --source /tmp/fx/DES-9999_Rev2.docx
```

## Planted defects

| # | Defect | Catches | Where it shows up |
|---|---|---|---|
| 1 | Header title "NIPT Assay Specification" vs document title "NIPT Assay Validation Specification" | Check 1 | `auto_checks` check 1 = FAIL |
| 2 | Header reads `Rev. 2`; latest revision history row reads `3` | Check 2 | `auto_checks` check 2 = FAIL; `revision_history.latest_rev` = `3` |
| 3 | Rev 3 entry claims "Added Section 9 Environmental Monitoring"; no §9 exists | Check 3 | No §9 in the section tree |
| 4 | §5.2 acceptance criterion changed 95% → 98% with no revision history entry | Check 4 | `all_numeric_deltas` contains `§5.2 Acceptance Criteria: 95% → 98%` |
| 5 | §3 requires 2–8 °C; §7 requires room temperature (15–25 °C) | Check 5 | Read across sections; not automated |
| 6 | §2 cites ISO 13485:2003, ISO 15223-1:2016, and "21 CFR 820 Quality System Regulation" | Check 6 | `citations` with `edition_year` 2003 and 2016 |
| 7 | `TBD` in §6; heading numbering skips 4.3; §1 references a non-existent §9 | Check 7 | `hygiene.placeholders`, `hygiene.numbering`, `hygiene.cross_references` |

## Expected script output

`extract_doc.py` on Rev 2:

```
doc id        : DES-9999
header rev    : 2
rev history   : 2 entries, latest Rev 3
[FAIL  ] 1 Title vs header
[FAIL  ] 2 Header revision vs revision history
[PASS  ] 2b Document ID in header
[PASS  ] 7a Page x of y fields present
[PASS  ] 7b Footer present
```

Section tree must be exactly: 1, 2, 3, 4, 4.1, 4.2, **4.4**, 5, 5.1, 5.2, 6, 7, 8.
There must be no section `21` — the reference list entry "21 CFR 820 …" is a
list item, not a heading, and a regression that reintroduces "Section 21" means
the heading heuristic has broken.

`hygiene` must report:
- `placeholders`: one `TBD` in §6;
- `numbering`: one gap, `4.3` missing under parent `4` — and **no** spurious top-level gap;
- `cross_references`: one `broken_section_reference` to Section 9 from §1;
- `acronyms.undefined`: `DAISY` and `NIPT` only (not `CLSI`, not `VER`).

`diff_revisions.py` with the arguments in either order:

```
old           : DES-9999_Rev1.docx (Rev 1)
new           : DES-9999_Rev2.docx (Rev 2)
sections +/-  : 1 added, 1 removed      # §4.4 added, §4.3 removed
numeric deltas: 3
    §3 Storage Conditions: 15 -> 2
    §3 Storage Conditions: 25 -> 8
    §5.2 Acceptance Criteria: 95% -> 98%
```

`all_numeric_deltas` must contain **exactly** those three. A regression that adds
`§6: 0045 -> None` (the document ID `VER-0045` read as a quantity) or
`§1: None -> 9` (the cross-reference "Section 9" read as a quantity) means the
`NOT_A_QUANTITY_RE` mask has broken.

`change_statements` must not contain "Section added: NIPT Assay Validation
Specification" — the Title paragraph is not a section.

`verify_annotation.py` must report red runs at `('Arial', 10.0)`, matching the
source table font. A red run with no inherited font means the `w:rPr` deepcopy
has broken and Word will render the notes in the default font.

## Expected review

`assets/findings-template.json` is the worked findings file for this pair: all
seven checks FAIL, six red notes on row 2, and ten recommendations. A review of
this fixture should land close to it.

Results: checks 1–7 all **FAIL**. Six red notes, all on revision history row 2.
Recommendations: three Major (the §3/§7 storage conflict, the title
misalignment, the `TBD` verification input), five Minor (three superseded or
retitled citations, the broken §9 reference, the 4.3 numbering gap), two
Observation (undefined acronyms, template ID absent from the footer).

## Round-trip

Stripping an annotated file must restore the original content exactly:

```bash
python3 scripts/annotate_docx.py /tmp/out/full.docx --strip-annotations --out /tmp/out/stripped.docx
```

The paragraph and table-row text of `stripped.docx` must be identical to
`DES-9999_Rev2.docx`. The source file's checksum must be unchanged by any step.

## Caveats

These fixtures are synthetic, written for testing. They are not Pillar
documents, the product and values are invented, and the superseded citations are
deliberate. Do not reuse them as a template for a real document.

The fixtures do not yet cover: a revision history that lists newest first, a
revision history with a DCR column, alphabetic revisions (Rev A/B), a
first-page-only header, a merged-cell revision history, an initial release with
no previous revision, or a PDF input. Those are worth adding as the skill is
used on real documents.
