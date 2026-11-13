# Review checklist

The seven checks, expanded. Work through every one. Do not skip an item because
it "looks fine": each line of the summary table must cite what you actually saw.

## Vocabulary

**Result** — one per check:

| Result | When |
|---|---|
| **PASS** | Verified correct, with evidence you can quote. |
| **FAIL** | A defect found, with evidence you can quote. |
| **REVIEW** | Cannot be determined: the value is unreadable, the section is absent, the old revision was not provided, or an external fact could not be confirmed. Say which. |
| **N/A** | The check does not apply (e.g. check 4 on an initial release). |

Never report PASS for something you could not read. Never report FAIL on a
regulatory citation you could not confirm — that is REVIEW.

**Severity** — one per finding:

| Severity | Meaning |
|---|---|
| **Major** | Wrong or missing content with regulatory, safety or product impact: a requirement contradicted, a revision history that misrepresents what changed, an undocumented change to an acceptance criterion, superseded legislation presented as current. |
| **Minor** | Internally inconsistent or incorrect, but without product impact: numbering gaps, a stale cross-reference, inconsistent citation formatting, a wrong date. |
| **Observation** | Cosmetic or a reviewer judgement call: wording that could be clearer, an acronym used before definition, a formatting inconsistency. |

## Mapping to the company QA REVIEW section

The findings are written so they can be pasted into the eQMS QA REVIEW fields
that `dcr-change-order-reviewer` documents. Tag each recommendation with the
QA-Q it answers.

| QA-Q | What QA must check | Checks that answer it |
|---|---|---|
| QA-Q4 | Action (Initial Release / Update / Replace / Obsolete) matches the Details table and revision history | 2, 4 |
| QA-Q6 | Required document updates + evidence; list any defects found | 3, 4, 5, 6, 7 |
| QA-Q7 | Revision history verified: rev number, date, description, author | 2, 3, 4 |
| QA-Q8 | Header verified: doc ID, title, rev, page x of y | 1, 2, 7 |
| QA-Q9 | Footer verified: controlled/uncontrolled statement, template ID | 7 |
| QA-Q10 | Correct template; template ID / rev current | 7 (partial — flag a mismatch, but you cannot know the current template revision; say so) |

---

## Check 1 — Title and header are aligned  *(requirement 2, QA-Q8)*

Compare every place the document names itself:

- the page header title, in **every** `w:hdr` part — default, first page and even page can differ, and a wrong title often hides on the first-page header only;
- the title on page 1 (the Title-styled paragraph or the first heading);
- the Word document property title (`docProps/core.xml`);
- the title in the footer, if the template puts one there;
- the file name, as supplied.

Report:
- a title that differs in wording, not just in case or punctuation;
- a document ID that appears in more than one form (`DES-9999` vs `DES-999`);
- a header that names a different document entirely;
- a title that no longer matches the document's scope after this revision — if the body now covers validation as well as specification, and the title does not, say so.

The automated match in `extract.json` ignores case and punctuation, so it can
pass a title that a reviewer would reject. Read the two strings yourself.

## Check 2 — Header revision matches the revision history  *(requirement 3, QA-Q7/Q8)*

- The revision in the page header must equal the revision of the **latest** entry in the revision history table. "Latest" means highest revision, not last row — some templates list newest first.
- The revision must be the same in every header part and on every page.
- The date on the latest entry must not be in the future, and must not predate the entry above it.
- An initial release must read "Initial release" (or the template's equivalent) and must be Rev 1 / Rev A / Rev 0 as the template requires.
- If the document ID appears in the revision history, it must match the header.
- Where a DCR / change order column exists, the latest entry should carry a change order reference. A blank one is a finding against QA-Q4.

A mismatch here gets **both** a red note in the revision history row **and** a
line in the orange block, because the reviewer needs to see it in context and in
the summary.

## Check 3 — Each revision history entry is evidenced in the body  *(requirement 4, QA-Q6)*

For **every** entry in the revision history table — not only the latest — take
each claim in its description and find it in the document:

- "Updated Section 5.2 acceptance criterion to 98%" → §5.2 must say 98%.
- "Added Section 9 Environmental Monitoring" → §9 must exist and must be about environmental monitoring.
- "Removed reference to the discontinued buffer" → the reference must be gone.
- "Updated to the current template" → the footer / template ID should show it.
- "Administrative changes only" → there should be no substantive change; if there is one, that is a Major finding.

Record, per entry: the claim, where you looked, what you found, and the verdict.
A claim you cannot place is **FAIL — claimed but not done**, not REVIEW, provided
the document was readable.

Entries for revisions **before** the current one describe changes already
incorporated; verify them against the body as it now stands, and treat an entry
whose change has since been reverted as a finding worth an Observation.

## Check 4 — The revision history is a complete and accurate account of the differences  *(requirements 5 and 6, QA-Q7)*

Runs only when the previous revision was provided. Otherwise **N/A — previous
revision not provided**, said plainly.

First state which file you treated as the newer revision and why (the header
revision). If the two headers give the same revision, or either is unreadable,
stop and ask.

Then run the comparison in **both** directions, using `diff.json`:

**(a) Done but not claimed.** Every entry in `change_statements` must be
described by the latest revision history entry. Pay particular attention to
`all_numeric_deltas`: a changed acceptance criterion, limit, tolerance,
temperature or volume that no entry mentions is a **Major** finding. Write the
missing statement into the table in red, phrased as the entry should have read:

> Missing entry: the call rate acceptance criterion in §5.2 changed from 95% to 98%.

**(b) Claimed but not done.** Any claim in the latest entry with no corresponding
difference. Red:

> Claimed but not done: no §9 Environmental Monitoring exists in this revision.

**(c) Claimed inaccurately.** The change happened but the entry describes it
wrongly — wrong section, wrong value, wrong direction. Red, giving the correction:

> Incorrect: §3 changed to 2–8 °C, not 2–8 °F.

Judgement to apply:
- Formatting-only differences (reflowed paragraphs, a changed page break, the revision history table itself, page numbers) are **not** changes that need an entry. Do not pad the findings with them.
- A reworded sentence with no change in meaning is an Observation at most.
- A change in a requirement, acceptance criterion, limit, instruction, warning, intended use, material or responsibility **always** needs an entry.
- Sections renumbered because one was inserted need a single entry about the insertion, not one per renumbered section.

## Check 5 — Internal consistency  *(requirement 7, QA-Q6)*

Conflicts **between sections of the same document** are the most common defect
this check finds. Look for:

- the same quantity stated differently in two places — storage temperature, hold time, sample volume, acceptance criterion, shelf life;
- a scope or intended-use statement that the body contradicts;
- a term defined one way in the definitions section and used another way later;
- the same concept under two names (e.g. "call rate" and "reportable rate") where it is unclear whether they are the same thing;
- a procedure whose steps conflict with a table or figure in another section;
- responsibilities assigned to two different roles;
- a referenced document whose own revision is pinned in one place and not in another;
- "shall" requirements that cannot all be met at once.

`extract.json` gives you the section tree and every table; read across sections,
not down one. Quote both conflicting statements with their section numbers.

## Check 6 — Regulatory and standards currency  *(requirement 7, QA-Q6)*

For every citation in `extract.json` → `citations`:

1. Look it up in `references/regulatory-currency.md`.
2. If that file says the cited edition is superseded, **confirm it by web search** before reporting.
3. Report confirmed findings as orange recommendations, each saying what to cite instead and ending with "confirm against the controlled standards list".
4. If a search cannot confirm, report **REVIEW** and say what you could not confirm.

Also report, without needing a lookup:
- the same standard cited with different edition years in different sections;
- inconsistent citation formatting (`ISO-13485-2016` vs `ISO 13485:2016`);
- a standard cited with no edition year where the document's purpose needs one;
- the wording problems listed at the end of `regulatory-currency.md`.

Never state an edition year you have not confirmed. "ISO 13485:2003 is cited;
the current edition appears to be 2016 — confirm against the controlled standards
list" is correct. Inventing "superseded by ISO 13485:2021" is not.

## Check 7 — Document hygiene  *(requirement 7, QA-Q9/Q10)*

From `extract.json` → `hygiene` and `auto_checks`:

- **Footer** — present, carries the controlled/uncontrolled statement and the template ID.
- **Page x of y** — uses `PAGE`/`NUMPAGES` fields, not typed text that will not renumber.
- **Placeholders** — `TBD`, `TBC`, `XXX`, `???`, `<insert …>`, highlighted runs.
- **Draft markup left in** — tracked insertions or deletions, comments still attached.
- **Numbering continuity** — gaps and duplicates in heading numbers, table and figure captions.
- **Cross-references** — every "see Section 7.3", "Table 4", "Figure 2" and referenced document ID resolves. Collect the referenced document IDs: they are the companion documents the change order will need.
- **Acronyms** — expanded at first use; defined but never used; used but never defined.
- **Orphan sections** — a heading with no content under it.
- **Ambiguous requirements** — "should", "may", "as appropriate", "if necessary" in a statement that verification will later have to test. Observation, but worth raising: it is what makes a specification unverifiable at V&V.

Numbering and cross-reference findings are Minor unless the broken reference
points at safety or regulatory content.
