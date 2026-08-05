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

Try it with the files in `sample-files/`, which contain deliberately messy data covering every QA rule.

**Hosted:** <https://dk-centific.github.io/excel-merger/> — published from `main` via GitHub Pages, open to
anyone with the link.

---

## Deploying

Live hosting is **GitHub Pages**, served from the root of `main`. Pushing to `main` republishes automatically;
there is nothing to configure.

Note that **GitHub Pages cannot send custom response headers**, so the `Content-Security-Policy` and
`Cache-Control: no-store` defined in `staticwebapp.config.json` are *not* applied while hosting is Pages. That
file only takes effect on Azure Static Web Apps.

### Alternative: Azure Static Web Apps with sign-in

Kept as an option if this ever needs to stop being publicly reachable. It puts the app behind Microsoft 365
sign-in so only invited colleagues can load it, and it activates the security headers above. The workflow at
`.github/workflows/azure-static-web-apps.yml` is already in place and no-ops until the token secret exists.

#### One-time setup

1. **Azure Portal → Create a resource → Static Web App.**
   - Plan type: **Free**
   - Deployment source: **GitHub** → authorise → org `DK-Centific`, repo `excel-merger`, branch `main`
   - Build presets: **Custom**, with app location `/`, api location empty, output location empty

   Azure will commit its own workflow file. Delete that one and keep
   `.github/workflows/azure-static-web-apps.yml` from this repo — it sets `skip_app_build: true`, which the
   generated one does not, and without it the deploy fails looking for a build step.

2. **Copy the deployment token** — Static Web App → *Manage deployment token*.

3. **Add it to GitHub** → repo *Settings → Secrets and variables → Actions → New repository secret*, named
   exactly `AZURE_STATIC_WEB_APPS_API_TOKEN`. Add it yourself; it is a credential and should not be pasted
   into chat.

4. **Invite each colleague** — Static Web App → *Role management* → *Invite*, pick **Microsoft Entra ID** as
   the provider, enter their work email, and assign the role **`merger`** (spelled exactly that way; it is
   what `staticwebapp.config.json` requires). Send them the generated invitation link.

Anyone who signs in without that role lands on `403.html` telling them to ask for access, rather than a bare
Azure error page.

#### Tenant-wide access instead of invitations

The Free plan's built-in Entra ID provider is a Microsoft-managed multi-tenant app, so *any* Microsoft account
can authenticate — the `merger` role is what actually restricts access, and that means inviting people one at
a time. The Free plan also caps how many custom-role users you can invite (25 at the time of writing; check
current Azure limits).

To let everyone at Centific in automatically, upgrade the Static Web App to the **Standard** plan and register
your own Entra ID application with a tenant-scoped issuer:

```jsonc
// staticwebapp.config.json — add alongside "routes"
"auth": {
  "identityProviders": {
    "azureActiveDirectory": {
      "registration": {
        "openIdIssuer": "https://login.microsoftonline.com/<YOUR_TENANT_ID>/v2.0",
        "clientIdSettingName": "AAD_CLIENT_ID",
        "clientSecretSettingName": "AAD_CLIENT_SECRET"
      }
    }
  }
}
```

Then change the `/*` route's `allowedRoles` from `["merger"]` to `["authenticated"]` — the tenant-scoped
issuer is already doing the restricting. Store `AAD_CLIENT_ID` and `AAD_CLIENT_SECRET` in the Static Web App's
*Configuration* settings, never in this repo. Custom identity providers require the Standard plan.

#### Security headers

`staticwebapp.config.json` defines a strict `Content-Security-Policy` limiting network access to Microsoft
Graph, Entra ID login and SharePoint, plus `no-store` caching so merged participant data is never written to a
shared cache. **These apply on Azure only** — Pages ignores the file entirely. If you add a library or a new
API, widen the policy there; the app has no inline scripts or styles, and keeping it that way is what lets the
policy stay strict.

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

The **Upload files** tab works immediately with no setup. The **SharePoint links** tab is a two-step flow:

1. **Sign in to Microsoft 365.** Until you do, step 2 stays locked — there is no point choosing a location
   before the app knows whose permissions apply.
2. **Paste SharePoint or OneDrive links**, one per line. A folder link pulls in every `.xlsx` inside it.

### Access is your account's access

The app uses **delegated** permissions, so it acts *as the signed-in user* and never has an identity of its
own. Effective access is the intersection of what the app may request and what that person can already reach:

- If your account can open a SharePoint location, the tool can read it.
- If it cannot, Graph returns 403/404 and that link is reported as failed.
- A colleague signing in sees only what *their* account can see.
- Revoking someone's SharePoint access revokes it here too, immediately. MFA and Conditional Access apply
  normally, because it is an ordinary interactive sign-in.

`Files.Read.All` sounds alarming but, as a *delegated* scope, it means "all files **the signed-in user can
access**" — not all files in the tenant. The `.All` only distinguishes it from `Files.Read`, which covers just
your own OneDrive and cannot reach SharePoint sites shared with you. This is usually the sticking point when
asking an admin to approve it. (`Sites.Read.All` may turn out to be unnecessary — both Graph calls the app
makes are likely satisfied by `Files.Read.All` alone. Worth testing once a registration exists, since it makes
the admin ask smaller.)

The tool only ever **reads**. It never writes back to SharePoint; merged output is saved through a normal
browser download. Saving results back to a SharePoint folder would need `Files.ReadWrite.All` and a code
change.

Tokens are held in `sessionStorage`, so closing the tab discards them — deliberate, for shared machines.

### One-time registration

This is set up **once per deployment** — end users never open the Settings dialog. Send your Microsoft 365
admin the following:

> Please register a single-page application in Azure AD:
>
> 1. **Azure Portal → Microsoft Entra ID → App registrations → New registration**
> 2. Name: `Agency Excel Merger`
> 3. Supported account types: *Accounts in this organizational directory only*
> 4. **Redirect URI: platform `Single-page application (SPA)`** — this must be the SPA platform, not Web.
>    Add `https://dk-centific.github.io/excel-merger/`, plus `http://localhost:8765/` if anyone runs it
>    locally. Several redirect URIs on one registration is fine.
> 5. After creating it: **API permissions → Add → Microsoft Graph → Delegated** →
>    `Files.Read.All` and `Sites.Read.All` → then **Grant admin consent**.
>    These are *delegated*, so the app can only ever reach files the signed-in user could already open.
> 6. Send back the **Application (client) ID** and **Directory (tenant) ID**.

Put the two IDs in **`config.js`** and commit — that bakes the connection into the deployment so nobody else
has to configure anything. The Settings dialog is a per-browser override, useful for testing a different
registration without redeploying.

Neither ID is a secret. A public-client registration has no client secret (it uses PKCE), every user still
signs in with their own account, and choosing *Accounts in this organizational directory only* means accounts
outside your tenant cannot sign in at all. The app's Settings dialog always displays the exact redirect URI
for wherever it is currently running, so you can copy it from there if the URL changes.

---

## Layout

```
index.html                  markup
403.html                    shown to signed-in users who lack the merger role
config.js                   optional Azure client/tenant ID defaults
staticwebapp.config.json    Azure auth rules, routing and security headers
css/styles.css              Centific design system tokens and components
js/core.js                  schema, QA rules, merge, workbook building  ← the logic lives here
js/sharepoint.js            MSAL sign-in and Microsoft Graph download
js/app.js                   UI wiring
vendor/                     ExcelJS and MSAL, vendored so there are no CDN calls
sample-files/               three agency files with deliberate data problems
.github/workflows/          Azure Static Web Apps deployment
```

To change a QA rule, edit `js/core.js` — the canonical header list, the column constants and each rule are all
at the top of that file.

Supported input formats are `.xlsx` and `.xlsm`. Older `.xls` and `.csv` files must be re-saved as `.xlsx`
first; the app reports this clearly rather than failing silently.
