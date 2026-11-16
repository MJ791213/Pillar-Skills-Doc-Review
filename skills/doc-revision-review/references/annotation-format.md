# Annotation format

The wording templates for what goes into the document. Keep to them so two
reviews of two documents read the same way.

## Where the output goes

| Colour | Where | What it carries |
|---|---|---|
| **Red** (`FF0000`) | Inside the revision history table, in the row of the entry it concerns | Only findings about the revision history: entries that are missing, claimed but not done, or described incorrectly, and a revision number that disagrees with the header. |
| **Red** (`FF0000`) | The **Errors** table in the appended block | Everything that is *wrong* anywhere in the document. |
| **Orange** (`ED7D31`) | The **Recommendations** table in the appended block | Everything that could be *better* but is not wrong. |

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
Recommendations worded as something to check, or list it under
`checks_not_performed`.

Nothing else in the document is touched. No body text is edited, no value is
corrected in place, no section is added. `annotate_docx.py` writes a new file and
leaves the source untouched — do not work around that.

Red runs inherit the font of the cell they are added to, so a note reads as part
of the table rather than as pasted-in text. Every inserted run begins with the
`[Claude]` marker so a later round can strip it with
`annotate_docx.py --strip-annotations`.

## Red: revision history notes

One note per finding, in the row of the entry it concerns. Start with the
finding type so a reviewer can scan the column.

**Missing entry** — a real difference between the revisions that no entry describes:

```
Missing entry: the call rate acceptance criterion in §5.2 changed from 95% to 98%.
```
```
Missing entry: §4.3 Equipment was renumbered to §4.4 when a section was inserted.
```

**Claimed but not done** — an entry describing a change that is not in the document:

```
Claimed but not done: no §9 Environmental Monitoring exists in this revision.
```

**Incorrect** — the change happened, but the entry describes it wrongly. Give the correction:

```
Incorrect: §3 was changed to 2-8 degrees C, not to 2-8 degrees F.
```
```
Incorrect: the storage change is in §3, not §5 as this entry states.
```

**Revision number** — the header and this row disagree:

```
Revision column reads 3 but the page header reads Rev. 2. One of the two is wrong.
```

**Not evidenced** — a claim from any revision that cannot be found in the body
(check 3, where no old revision was supplied):

```
Not evidenced in the body: this entry claims the buffer reference was removed; §4.1 still names it.
```

Rules for red notes:

- One finding per note. Three findings on one entry means three notes, not one paragraph.
- Quote the section number and the old → new value whenever both are known.
- Never write a note that asserts what the entry *should* say unless the comparison actually showed it.
- Never invent a document ID, a date, a name or a value. Use `[TBD - author to confirm]`.
- If the revision history table was not found, write no red notes at all and say so in the chat reply. Do not annotate another table.

## The appended review block

`annotate_docx.py` builds this from `findings.json`. It contains, in order:

1. The heading `Claude Review — not part of the approved document`.
2. A provenance line: the date, the source file, the document ID and revision, and a statement that the block is not approved content and must be removed before release.
3. **Errors** — red, two columns:

   | Section | Error |
   |---|---|
   | §3 and §7 | Conflicting storage requirements. §3 requires the kit at 2-8 °C; §7 requires samples at room temperature (15-25 °C). Both are written as requirements and cannot both be met. Decide which condition controls and align both sections. |

4. **Recommendations** — orange, two columns:

   | Section | Recommendation |
   |---|---|
   | §7 Revision History | Describe changes by section number and old → new value rather than in general terms, so each entry is verifiable against the body. |

5. **Not checked** — a bullet per check that could not be completed, when there are any. An empty Errors table must never be read as "nothing is wrong" when it means "not looked at".
6. The closing note: this is a Claude recommendation, not reviewed or approved by QA, and every regulatory citation must be confirmed against the controlled standards list.

### Writing the rows

- **Section** locates the finding: a section number (`§5.2`), two when it is a conflict (`§3 and §7`), a revision history row (`§7 Revision History, row 2 (Rev 3)`), a table caption (`Table 5`), or `Page header` / `Footer` / `Document-wide`.
- **Error** states what is wrong, quotes the evidence, and ends with what needs to happen. One finding per row — three problems in one section means three rows.
- **Recommendation** states the improvement and why it helps, in one or two sentences.
- Give the old → new value whenever both are known: `95% → 98%`.
- Regulatory errors end with `Confirm against the controlled standards list.`
- Never invent a document ID, revision, date, name or value — write `[TBD – author to confirm]`.
- The tables have exactly two columns. `severity` and `qa_q` stay in `findings.json` (severity orders the Errors table, Major first; both feed the change order) but are not printed.

## findings.json

`assets/findings-template.json` is a worked example. The shape:

```json
{
  "document": {"doc_id": "DES-9999", "revision": "2", "title": "..."},
  "revision_history_table_index": 0,
  "revision_history_notes": [
    {"row_index": 2, "col_index": 3, "text": "Missing entry: ..."}
  ],
  "errors": [
    {"section": "§3 and §7", "severity": "Major", "qa_q": "QA-Q6",
     "error": "..."}
  ],
  "recommendations": [
    {"section": "Page header", "recommendation": "..."}
  ],
  "checks_not_performed": ["Check 5 — internal consistency: not completed"]
}
```

`revision_history_table_index` and each `row_index` come straight from
`extract.json` → `revision_history` (`table_index`, and the `row_index` of the
entry). Do not count rows by hand. `col_index` is optional; omitted, the note
goes in the last cell of the row, which is the description column in the
templates seen so far — set it explicitly from
`revision_history.columns.description` when that is present.

## The chat reply

The deliverable is the annotated file, so keep the reply short:

- which file was treated as the newer revision, and which as the older, and why;
- a one-line result per check;
- the count of red notes and orange recommendations;
- the path to the annotated file;
- anything you could not check, and why.

Do not reproduce the findings tables in the chat; they are in the document.
