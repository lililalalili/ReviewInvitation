# NB Review Invitation Agent v15 — Product Requirements

## Objective
Upgrade the existing v14 script into a stable downloadable Windows application with:
1. Excel schema upgrade and delayed invitation-date writing.
2. Author web/email/research enrichment through external search + DeepSeek.
3. GUI review workflow.
4. Outlook invitation workflow with manual confirmation.
5. Batch invitation support with logging, duplicate prevention, and dry-run tests.

## Source files expected in runtime directory
- `NB_Author_2026.xlsm`
- `ReviewInvitationLog.xlsx`
- `NB_Template_Review_Yes.docx`
- `NB_Template_Review_No.docx`
- `NB_Template_Insight.docx`

## Required Excel columns
Keep existing columns and add if missing:
- `Affiliation of the First Author`
- `Last Author Research`
- `Last Author Web`
- `Batch ID`
- `Author Enrichment Status`
- `Author Enrichment Evidence`
- `Invitation Status`
- `Invitation Error`

Important behavior:
- `Date of Invitaion` must remain blank after information collection/import.
- `Date of Invitaion` is written only after Outlook `Send()` succeeds.
- `Pubmed Link` should be a real PMID article URL when PMID exists: `https://pubmed.ncbi.nlm.nih.gov/{PMID}/`.

## Enrichment workflow
For each row:
- Last author inputs: full name, last author email, last author affiliation.
- Find correct personal/lab/profile page.
- Extract main research direction, preferably as a direct quote from the web page.
- Write URL to `Last Author Web` and research text to `Last Author Research`.
- If last author email is missing, search and fill `Email of the Last Author` only when confidence is high.
- If first author email is missing, search by first author full name + first author affiliation and fill `First Author Email` only when confidence is high.
- Low confidence results should be marked `Needs Review`, not silently written as fact.

Recommended JSON returned by DeepSeek extraction:
```json
{
  "is_correct_author": true,
  "confidence": 0.86,
  "personal_web_url": "https://...",
  "research_quote": "...",
  "last_author_email": "...",
  "first_author_email": "...",
  "evidence": "...",
  "reason": "..."
}
```

## GUI behavior
Open automatically after information collection.
Default view: first row of the latest `Batch ID`.
Batch selector: dropdown list.
Display these columns:
- Full Name of the Last Author
- Affiliation of the Last Author
- Country
- Overseas
- Last Author Web
- Email of the Last Author
- First Author Full name
- First Author Email
- Journal
- Title
- Title (CN)
- Pubmed Link
- Research field
- Macro Research Field
- Meso Research Field
- Micro Research Field
- Review Extension Potential
- Invited Review Angle
- Angle Rationale (CN)
- Editorial Bucket
- Manual Decision
- Neuro Decision Source
- Date of Invitaion

Editable fields:
- `Overseas`: dropdown Yes/No
- `Manual Decision`: dropdown Review/Insight/No
- `Research field`: editable text

Link buttons:
- `Last Author Web`: open URL in default browser
- `Pubmed Link`: open URL in default browser
- `Email of the Last Author`: open Bing search for the email address

Navigation/action buttons:
- Previous row
- Next row
- First row in batch
- Last row in batch
- Save all modifications
- Invite current author
- BatchInvitation

## Invitation workflow
Before creating/sending email:
1. Read `Manual Decision`.
   - If `No`, stop and show: `Manual Decision = No，请确认是否邀请`.
2. Read `Date of Invitaion`.
   - If non-empty, stop and show: `已邀请过`.
3. If date is blank, choose template:
   - Manual Decision = Review and Overseas = Yes: `NB_Template_Review_Yes.docx`; subject: `Neuroscience Bulletin Invites You to Submit a Review`.
   - Manual Decision = Review and Overseas = No: `NB_Template_Review_No.docx`; subject: `Neuroscience Bulletin Invites You to Submit a Review`.
   - Manual Decision = Insight: `NB_Template_Insight.docx`; subject: `Neuroscience Bulletin Invites You to Submit an Insight`.
4. Replace placeholders while preserving formatting:
   - `Aaaaa` = Family Name of the Last Author
   - `Jjjjj` = Journal, title case / proper journal capitalization
   - `Ttttt` = Title without trailing period
   - `Fffff` = Research field
   - `Pppppyes` = `For invited overseas authors, article publication charges will be covered by the journal, and NB will pay 1,000 USD remuneration for an accepted Review article.`
   - `Pppppno` = `For invited authors, article publication charges will be covered by the journal.`
   - `Dddddre` = current date + 6 months
   - `Dddddin` = current date + 125 days
5. To = `Email of the Last Author`.
6. CC = `First Author Email` if available.
7. SendUsingAccount = `nsb@ion.ac.cn`.
8. Display the generated email and show a confirmation dialog.
9. Only if the user confirms and Outlook send succeeds, write current date to `Date of Invitaion`.

## BatchInvitation behavior
- Iterate current batch only.
- Apply the same workflow as current-author invitation.
- Do not write invitation date for skipped/cancelled/failed rows.
- Record per-row status and error.
- Default to conservative behavior: stop or require confirmation for ambiguous rows.

## Release acceptance
A release is acceptable only if:
- Unit tests pass with fake search, fake DeepSeek, fake Outlook, and fixture workbook.
- No real secrets are committed.
- `.xlsm` macro preservation is tested or explicitly verified.
- Windows smoke-test instructions are present.
- Downloadable zip or release artifact includes source, `.bat`, requirements, README, and `.env.example`.
