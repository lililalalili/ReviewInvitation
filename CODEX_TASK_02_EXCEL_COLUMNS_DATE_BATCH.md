# Codex Task 02 — Excel schema, blank invitation date, PubMed link, Batch ID

## Goal
Implement the first functional upgrade in the workbook layer.

## Required changes
1. Add columns if missing:
   - `Affiliation of the First Author`
   - `Last Author Research`
   - `Last Author Web`
   - `Batch ID`
   - `Author Enrichment Status`
   - `Author Enrichment Evidence`
   - `Invitation Status`
   - `Invitation Error`
2. Populate `Affiliation of the First Author` from PubMed first-author affiliation.
3. Generate one `Batch ID` per run, e.g. `20260515_143012`, and write it for all newly appended rows.
4. Do not write `Date of Invitaion` in import/append flow. Leave blank.
5. Write `Date of Publication` as before.
6. Change `Pubmed Link` to `https://pubmed.ncbi.nlm.nih.gov/{PMID}/` when PMID exists.
7. Ensure `.xlsm` is opened with `keep_vba=True` and saved without losing macros.
8. Use header-name based lookup as much as possible.

## Acceptance tests
- Fixture workbook without new columns gets all required columns added.
- Newly appended rows have blank `Date of Invitaion`.
- `Batch ID` is same for all rows in one run and non-empty.
- `Affiliation of the First Author` is filled from parsed PubMed data.
- PubMed link equals `https://pubmed.ncbi.nlm.nih.gov/{PMID}/`.
- Existing formatting is preserved for added columns as far as practical.

## Regression checks
- Existing PMID duplicate logic still works.
- Existing `Manual Decision` remains blank for new rows.
- Existing run log update still happens only after successful run.
