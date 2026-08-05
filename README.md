# Merge — Agency Excel Consolidator

Merges Excel submissions from multiple agencies into one client-ready workbook, running a QA pass on every
row and reporting exactly what was corrected and what still needs a human.

Everything runs **inside the browser**. Files are never uploaded to a server — which matters here, because
the data contains participant names, emails, guardian contacts and birthdates.

---

## Running it

No build step, no npm, no Node. It is plain HTML, CSS and JavaScript with two vendored libraries.

**Locally:**

```bash
python3 -m http.server 8765 --directory /Users/davidk/Desktop/Projects/ExcelMerger
```

Then open <http://localhost:8765>.

**Hosted for the team:** upload the whole folder to any static host — Azure Static Web Apps, SharePoint, GitHub
Pages, Netlify. There is no backend to deploy.

Try it with the files in `sample-files/`, which contain deliberately messy data covering every QA rule.

---

## The column schema

Row 1 is the header row. All 21 columns, in order:

| Col | Header | Col | Header |
|-----|--------|-----|--------|
| A | Participant Name As Per ICF | L | Environment |
| B | Participant Email | **M** | **Head and Hair** |
| C | Name of Parents/Legal Guardian | **N** | **Facial Features** |
| D | Email of Parents/Legal Guardian | **O** | **Accessories and jewellery** |
| E | Expressions | **P** | **Others** |
| F | For Non Neutral, Please Select | **Q** | **None** |
| G | Age Group | R | Skintone |
| **H** | **Birthdate** | S | Device |
| I | Gender | T | Country of Collection |
| J | Main Ethnicity | U | State Abbreviation |
| K | Secondary Ethnicity | | |

Columns are matched **by header name, not by position**, so an agency that reorders its columns still merges
correctly. Header matching ignores case, spacing and punctuation, and tolerates a few common variants
(`DOB` → Birthdate, `jewelry` → `jewellery`). The header row is found automatically within the first 12 rows,
so a file with a title row above the headers still works.

A file whose headers cannot be matched is **skipped rather than merged**, and reported as an error.

---

## QA rules

Applied in this order — the order matters, because rule 3 can change the outcome of rule 4.

1. **Whitespace** — leading/trailing spaces trimmed, repeated spaces collapsed, non-breaking and zero-width
   characters removed. Cells left empty by this become genuinely empty.

2. **Birthdate (column H) → `YYYY/MM`.** Handles real Excel date cells, serial numbers, `YYYY-MM`, `MM/YYYY`,
   `YYYY/MM/DD`, `DD/MM/YYYY`, `MM/DD/YYYY`, `YYYYMM`, `YYYYMMDD`, and month names (`May 1992`).
   The column is written as **text format** so Excel cannot silently reinterpret `1990/05` as a date.

   Flagged for review rather than silently guessed:
   - Both day and month ≤ 12 (`05/06/1991` is genuinely ambiguous) — you choose day-first or month-first in
     Options, and every affected row is flagged either way.
   - Two-digit years, which get expanded but are worth confirming.
   - Year only (`1995`) or unrecognisable text — **the original value is kept**, never blanked.

3. **Column P (Others)** — a cell containing only a slash placeholder (`/`, `//`, ` / `) is emptied.
   A cell with real content is left completely alone, *including* content that happens to contain a slash:
   `Scar / birthmark` survives untouched.

4. **Head attribute block, M–Q** — evaluated per row, after rule 3:
   - If **any** of M, N, O, P has data → column Q is emptied.
   - If **all** of M–P are empty → column Q is set to `N/A-none apply`.

Also reported: missing columns, unrecognised extra columns (which are *not* carried into the output), and
possible duplicate participants across files (matched on email, falling back to name). Duplicates are
**flagged, never auto-removed** — both rows are kept for you to decide.

### Severity levels

| Level | Meaning |
|-------|---------|
| `AUTO-FIXED` | Corrected for you. Informational. |
| `NEEDS REVIEW` | Could not be fixed safely. Original value preserved. |
| `ERROR` | Structural problem with the file, e.g. a missing column. |

---

## Output

**Download merged Excel** produces a `Merged Data` sheet with exactly the 21 columns, bold frozen header and
autofilter — nothing extra, ready for the client. Ticking the option in step 2 appends `QA Report` and
`Sources` sheets to that same file; leave it off for the client copy.

**Download QA report** always gives you the full report separately, including a `Sources` sheet mapping each
agency file to the exact row range it occupies in the merged output. That is your traceability route back to
an agency without adding a column to the delivery sheet.

The on-screen table renders the first 500 issues for speed; the downloaded report always contains all of them.

---

## SharePoint links (optional)

The **Upload files** tab works immediately with no setup. The **SharePoint links** tab lets you paste
SharePoint/OneDrive URLs instead — including a folder link, which pulls in every `.xlsx` inside it.

That tab needs a one-time Azure app registration, because a browser cannot reach Microsoft 365 without one.
Send your Microsoft 365 admin the following:

> Please register a single-page application in Azure AD:
>
> 1. **Azure Portal → Microsoft Entra ID → App registrations → New registration**
> 2. Name: `Agency Excel Merger`
> 3. Supported account types: *Accounts in this organizational directory only*
> 4. **Redirect URI: platform `Single-page application (SPA)`** — the exact URL shown in the app's Settings
>    dialog. This must be the SPA platform, not Web.
> 5. After creating it: **API permissions → Add → Microsoft Graph → Delegated** →
>    `Files.Read.All` and `Sites.Read.All` → then **Grant admin consent**.
> 6. Send back the **Application (client) ID** and **Directory (tenant) ID**.

Paste those two IDs into the app's Settings dialog. To bake them into the deployment so nobody else has to,
put them in `config.js` instead — they are not secrets. A public-client registration has no client secret,
every user still signs in with their own account, and they can only reach files they already have permission
to open.

---

## Layout

```
index.html          markup
config.js           optional Azure client/tenant ID defaults
css/styles.css      Centific design system tokens and components
js/core.js          schema, QA rules, merge, workbook building  ← the logic lives here
js/sharepoint.js    MSAL sign-in and Microsoft Graph download
js/app.js           UI wiring
vendor/             ExcelJS and MSAL, vendored so there are no CDN calls
sample-files/       three agency files with deliberate data problems
```

To change a QA rule, edit `js/core.js` — the canonical header list, the column constants and each rule are all
at the top of that file.

Supported input formats are `.xlsx` and `.xlsm`. Older `.xls` and `.csv` files must be re-saved as `.xlsx`
first; the app reports this clearly rather than failing silently.
