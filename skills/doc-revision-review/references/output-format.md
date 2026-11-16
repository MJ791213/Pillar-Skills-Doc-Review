# Output format

How to word the review. Keep to it so two reviews of two documents read the
same way.

## The deliverable is two tables in the reply

No file. No annotated copy of the document. Nothing is written back into the
source, and nothing is returned for download. The scripts write JSON into the
scratchpad for you to read; that is working material, not output.

The reply is, in order:

1. One line: which file you treated as the newer revision and which as the older, and why. Omit it when only one document was supplied.
2. The **Errors** table — `Section | Error`.
3. The **Recommendations** table — `Section | Recommendation`.
4. **Not checked** — a bullet per check you could not complete, when there are any.
5. One closing line: these are recommendations for the reviewer, not approved content, and every regulatory citation needs confirming against the controlled standards list.

Nothing else. No per-check walkthrough, no summary of the document, no
description of what the scripts did.

## Errors or recommendation?

This is the only judgement that matters for the output, so make it deliberately.

> **An error is something a reviewer is obliged to act on before approval.**
> **A recommendation is something they may act on, and the document is correct
> either way.**

If you cannot decide, ask whether the document would be *defective* if released
as it stands. If yes, it is an error.

**Errors** — not a closed list; anything wrong belongs here:

- the revision number is wrong, or the header and the revision history disagree
- an entry in the revision history is wrong: it claims a change that was not made, describes a change inaccurately, or cites a section that does not exist
- a real change to the body has no revision history entry
- a change in the body is wrong in itself: a value, limit, instruction or criterion that contradicts the rest of the document or the change's own stated purpose
- **an update that does not make sense** — the change was made, but the stated reason does not match what was done, or the change is internally inconsistent or incomplete (applied in one place and not in another that needed it)
- two sections of the document conflict
- a cited regulation or standard is superseded, retitled or wrong
- a cross-reference, table or figure reference resolves to nothing
- a placeholder (`TBD`, `XXX`) or draft markup survives into a released revision
- a required element is missing: footer statement, template ID, approval, effective date

**Recommendations** — improvements, where nothing is actually wrong:

- a clearer way to display something: a table instead of prose, a column added, consistent units
- a better description or wording: a revision history entry that is accurate but vague, a requirement phrased so it cannot be tested
- formatting and presentation consistency
- terms used without definition where meaning is still clear from context
- structural suggestions: add a definitions section, add a DCR column, split a long section
- process observations worth noting but requiring no document change

A finding you could not confirm is **not** an error. Either put it in
Recommendations worded as something to check, or list it under Not checked.

## Wording the revision history errors

These are the most common rows in the Errors table, so keep them to a form.
Put the revision history row in the Section column (`§7 Revision History,
row 2 (Rev 3)`) and the finding in the Error column.

**Missing entry** — a real difference between the revisions that no entry describes:

```
Missing entry: the call rate acceptance criterion in §5.2 changed from 95% to 98%.
```

**Claimed but not done** — an entry describing a change that is not in the document:

```
Claimed but not done: no §9 Environmental Monitoring exists in this revision.
```

**Incorrect** — the change happened, but the entry describes it wrongly. Give the correction:

```
Incorrect: §3 was changed to 2-8 degrees C, not to 2-8 degrees F.
```

**Revision number** — the header and the row disagree:

```
The revision column reads 3 but the page header reads Rev. 2. One of the two is wrong.
```

**Not evidenced** — a claim that cannot be found in the body (check 3, where no
old revision was supplied):

```
Not evidenced in the body: this entry claims the buffer reference was removed; §4.1 still names it.
```

**Stale reference** — the entry cites a section whose number has since changed.
Only after the cross-revision check says `moved`:

```
This entry cites "Section 5.2 Product Assessment Documentation". That section is §5.6 in this revision; it moved when §5 was reworked, so the reference no longer resolves.
```

Rules:

- One finding per row. Three findings on one entry means three rows.
- Quote the section number and the old → new value whenever both are known.
- Never assert what an entry *should* have said unless the comparison showed it.
- Never invent a document ID, date, name or value. Use `[TBD - author to confirm]`.
- If the revision history table was not found, report that and stop; do not guess at another table.

## The appended review block

Write them as markdown, exactly these columns:

| Section | Error |
|---|---|
| §3 and §7 | Conflicting storage requirements. §3 requires the kit at 2-8 °C; §7 requires samples at room temperature (15-25 °C). Both are written as requirements and cannot both be met. Decide which condition controls and align both sections. |

| Section | Recommendation |
|---|---|
| §7 Revision History | Describe changes by section number and old → new value rather than in general terms, so each entry is verifiable against the body. |

Order the Errors table with the most serious first — wrong or missing content
with regulatory, safety or product impact, then internal inconsistency without
product impact, then cosmetic. Do not add a severity column; the order carries
it.

If there are no errors, write `No errors found by the checks that were
performed.` and still give the Not-checked list.

### Writing the rows

- **Section** locates the finding: a section number (`§5.2`), two when it is a conflict (`§3 and §7`), a revision history row (`§7 Revision History, row 2 (Rev 3)`), a table caption (`Table 5`), or `Page header` / `Footer` / `Document-wide`.
- **Error** states what is wrong, quotes the evidence, and ends with what needs to happen. One finding per row — three problems in one section means three rows.
- **Recommendation** states the improvement and why it helps, in one or two sentences.
- Give the old → new value whenever both are known: `95% → 98%`.
- Regulatory errors end with `Confirm against the controlled standards list.`
- Never invent a document ID, revision, date, name or value — write `[TBD – author to confirm]`.
- Two columns only. Severity is carried by the row order, not a column.

## Optional: the same review as a marked-up Word file

Only when the user asks for a file. The default is the tables in the reply and
nothing else.

`scripts/annotate_docx.py` writes a `_CLAUDE-REVIEW.docx` copy: red notes inside
the revision history table in the row of the entry they concern, inheriting the
cell's font, plus the same two tables appended at the end. The source file is
never modified. Build `findings.json` first —
`assets/findings-template.json` is a worked example:

```json
{
  "document": {"doc_id": "DES-9999", "revision": "2", "title": "..."},
  "revision_history_table_index": 0,
  "revision_history_notes": [
    {"row_index": 2, "col_index": 3, "text": "Missing entry: ..."}
  ],
  "errors": [
    {"section": "§3 and §7", "severity": "Major", "qa_q": "QA-Q6", "error": "..."}
  ],
  "recommendations": [
    {"section": "Page header", "recommendation": "..."}
  ],
  "checks_not_performed": ["Check 5 — internal consistency: not completed"]
}
```

```bash
python3 "${CLAUDE_SKILL_DIR}/scripts/annotate_docx.py" <new.docx> --findings <scratch>/findings.json --out <dir>/<NAME>_CLAUDE-REVIEW.docx
```

`revision_history_table_index` and each `row_index` come straight from
`extract.json` → `revision_history`. Do not count rows by hand. `col_index` is
optional; omitted, the note goes in the last cell of the row — set it from
`revision_history.columns.description` when that is present. In this path
`severity` orders the Errors table and `qa_q` feeds the change order, though
neither is printed.

Confirm the file afterwards rather than assuming it:

```bash
python3 "${CLAUDE_SKILL_DIR}/examples/verify_annotation.py" <annotated.docx> --source <new.docx>
```

To re-review after the author responds, strip the previous round first with
`annotate_docx.py --strip-annotations`.
