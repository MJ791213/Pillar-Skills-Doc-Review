# Annotation format

The wording templates for what goes into the document. Keep to them so two
reviews of two documents read the same way.

## The two colours, and nothing else

| Colour | Where | What it carries |
|---|---|---|
| **Red** (`FF0000`) | Inside the revision history table, in the row of the entry it concerns | Only findings about the revision history: entries that are missing, claimed but not done, or described incorrectly, and a revision number that disagrees with the header. |
| **Orange** (`ED7D31`) | A block appended at the very end of the document | Everything else: checks 1, 5, 6, 7, the review summary, and the statement that this is a Claude recommendation. |

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

## Orange: the recommendations block

`annotate_docx.py` builds this from `findings.json`. It contains, in order:

1. The heading `Claude Recommendations — not part of the approved document`.
2. A provenance line: the date, the source file, the document ID and revision, and a statement that the block is not approved content and must be removed before release.
3. **Review summary** — one row per check:

   | # | Check | Result | Evidence |
   |---|---|---|---|
   | 1 | Title vs header | FAIL | Header: "NIPT Assay Specification". Title: "NIPT Assay Validation Specification". |

   All seven checks appear, including the ones that passed and the ones that are
   N/A. A reviewer needs to see that a check ran.

4. **Recommended updates** — one row per finding, Major first:

   | Severity | QA-Q | Location | Finding | Suggested action |
   |---|---|---|---|---|
   | Major | QA-Q6 | §3 vs §7 | §3 requires 2-8 °C; §7 requires room temperature (15-25 °C). | Decide the controlling condition and align both sections. |

5. The closing note stating that the above is a Claude recommendation, has not
   been reviewed or approved by QA, and that every regulatory citation must be
   confirmed against the controlled standards list.

Writing rules for the orange rows:

- **Location** is a section number, or two when the finding is a conflict (`§3 vs §7`). For a header or footer finding, write `Page header` or `Footer`.
- **Finding** states the defect and quotes the evidence. No hedging, no "it appears that" where you saw it plainly.
- **Suggested action** is what the author should do, in one sentence. It is a suggestion, not an instruction — the reviewer decides.
- Regulatory findings end with `Confirm against the controlled standards list.`
- A finding you could not confirm goes in with severity **Observation** and the words "could not confirm" in the Finding column — not left out, and not stated as fact.

## findings.json

`assets/findings-template.json` is a worked example. The shape:

```json
{
  "document": {"doc_id": "DES-9999", "revision": "2", "title": "..."},
  "revision_history_table_index": 0,
  "revision_history_notes": [
    {"row_index": 2, "col_index": 3, "text": "Missing entry: ..."}
  ],
  "summary": [
    {"number": 1, "check": "Title vs header", "result": "FAIL", "evidence": "..."}
  ],
  "recommendations": [
    {"severity": "Major", "qa_q": "QA-Q6", "location": "§3 vs §7",
     "finding": "...", "suggested_action": "..."}
  ]
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
