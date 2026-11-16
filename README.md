# Pillar Skills — Document Review

A Claude Code plugin for reviewing controlled medical device / IVD documents
before approval.

## What it does

`doc-revision-review` acts as a QA Document Specialist and reviews a controlled
document (specification, SOP, work instruction, user manual, protocol) against
seven checks:

| # | Check |
|---|---|
| 1 | The document title and the page header are aligned |
| 2 | The revision number in the header matches the revision history table |
| 3 | Every change listed in the revision history is actually present in the body |
| 4 | When the previous revision is provided, the revision history is a complete and accurate account of the differences between the two |
| 5 | No conflicts between sections of the same document |
| 6 | No superseded standards or regulations cited |
| 7 | Document hygiene — footer, page numbering, heading continuity, cross-references, placeholders, acronyms |

It reports the findings **in the reply, as two tables**:

| Section | Error |
|---|---|
| What is wrong: a wrong revision number, a revision history entry that is wrong or unsupported, a change that is undescribed or does not make sense, a conflict between sections, an outdated regulation, a broken reference. |

| Section | Recommendation |
|---|---|
| What could be better but is not wrong: clearer formatting, better wording, consistency, structural suggestions. |

No file is produced and the source document is never modified. A marked-up Word
copy is available on request (`scripts/annotate_docx.py`), but it is not the
default.

Which file is the newer revision is decided from the revision in the page
header, not from the file name or the upload order.

## Install

```bash
/plugin marketplace add MJ791213/Pillar-Skills-Doc-Review
```

```bash
/plugin install pillar-doc-review@pillar-skills
```

Then upload one or two revisions of a document and ask for a review, or invoke
it directly:

```
/pillar-doc-review:doc-revision-review
```

## Use

Upload the new revision, and the previous revision if there is one. Check 4 —
the comparison against the previous revision — is the most valuable part of the
review and only runs when the old revision is supplied.

A `.pdf` is reviewed the same way but cannot be annotated in place; the findings
come back as a separate `.docx` review record.

## Layout

```
.claude-plugin/
  plugin.json                      plugin manifest
  marketplace.json                 marketplace manifest
skills/doc-revision-review/
  SKILL.md                         the review workflow
  scripts/
    docx_common.py                 shared .docx traversal (headers, tables, headings)
    extract_doc.py                 header/footer/title/revision history/sections/citations
    diff_revisions.py              section-aware old-vs-new comparison
    annotate_docx.py               optional marked-up .docx; --strip-annotations
  references/
    review-checklist.md            the seven checks, line by line
    regulatory-currency.md         current standard editions, with per-row provenance
    output-format.md               how to word the two tables; the optional Word output
  assets/
    findings-template.json         worked findings file for the fixture pair
  examples/
    make_fixtures.py               generates two fixture pairs with planted defects
    selftest.py                    asserts the whole pipeline (96 assertions)
    expected-findings.md           the regression baseline
    verify_annotation.py           asserts the annotations are really in the output
```

## Develop

Load the plugin without installing it:

```bash
claude --plugin-dir ./
```

Validate the manifests:

```bash
claude plugin validate .
```

Run the self-test. It builds the fixture pair, runs the whole pipeline and
asserts the baseline in `examples/expected-findings.md`. Run it after changing
any script:

```bash
cd skills/doc-revision-review && python3 examples/selftest.py
```

It makes 96 assertions across two fixture pairs and exits non-zero if any fails. It covers: a clean Rev 1
raising no findings, each planted defect in Rev 2 being found, the newer
revision being identified from the header rather than the argument order, the
numeric-delta mask rejecting document IDs and cross-references, red notes
inheriting the table font, stripping restoring the original content exactly,
the source file never being modified, and the scripts refusing rather than
guessing on bad input.

To keep the generated files and look at them:

```bash
cd skills/doc-revision-review && python3 examples/selftest.py --keep
```

The self-test checks the scripts, not Claude's judgement. It cannot tell you
whether a review reads well — only whether the mechanics still work.

Requires `python-docx`; PDF input also needs `pdfplumber`.

## Caveats

- `references/regulatory-currency.md` is a convenience aid, not a controlled
  standards register. Every row carries a confidence mark and a verification
  date, the skill confirms a citation by web search before reporting it as
  outdated, and every regulatory finding is marked as needing confirmation
  against the company's controlled standards list.
- Revision history table layouts vary by template. Detection is heuristic; if no
  revision history table is found the skill reports that and annotates nothing
  rather than marking up the wrong table.
- The skill assists the reviewer. It does not replace the approval signature.

## Related

`doc-revision-review` is the document-side counterpart to the
`dcr-change-order-reviewer` skill: that one fills in the Change Order, this one
verifies the documents the Change Order covers. Findings are tagged with the
QA REVIEW question they answer (QA-Q6 to QA-Q10) so they can be pasted into the
eQMS.
