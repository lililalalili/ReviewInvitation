# NB Review Invitation HTML Templates

These HTML templates are designed for the NB Review Invitation Agent and Outlook `HTMLBody` rendering.

Files:
- `NB_Template_Review_Yes.html`
- `NB_Template_Review_No.html`
- `NB_Template_Insight.html`

Required placeholders preserved exactly:
- `Aaaaa`
- `Jjjjj`
- `Ttttt`
- `Fffff`
- `Pppppyes`
- `Pppppno`
- `Dddddre`
- `Dddddin`

Notes:
- Dynamic placeholder values should be HTML-escaped by the renderer before replacement.
- The templates use simple Outlook-friendly HTML with inline CSS.
- The old `.docx` templates can remain in the repository as fallback/archive, but these `.html` files are intended to be the primary templates.
