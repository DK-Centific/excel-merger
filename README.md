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

**When you change a CSS or JS file, bump the `?v=` on its `<link>`/`<script>` tag in `index.html`.** Pages
serves assets with `Cache-Control: max-age=600`, so without it a returning visitor can get the new HTML against
a ten-minute-old stylesheet — which renders as missing or broken styling rather than an obvious error. There is
no build step to hash filenames automatically, so this one is manual. Keep the number the same across all the
tags so a single find-and-replace bumps them together.

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
   `Face/neck tattoos` survives untouched.

4. **Columns M–P accept only the client's vocabulary** (see below). Empty is always fine. A near miss is
   corrected automatically; anything genuinely off the list is **kept exactly as submitted** and flagged.

5. **Head attribute block, M–Q** — evaluated per row, and deliberately last, so that a cell cleared by rule 3
   correctly counts as empty here:
   - If **any** of M, N, O, P has data → column Q is emptied.
   - If **all** of M–P are empty → column Q is set to `N/A-none apply`.

Also reported: missing columns, unrecognised extra columns (which are *not* carried into the output), and
possible duplicate participants across files (matched on email, falling back to name). Duplicates are
**flagged, never auto-removed** — both rows are kept for you to decide.

### Accepted values for M–P

| Col | Header | Accepted values |
|-----|--------|-----------------|
| M | Head and Hair | Glasses · Religious headwear · Hat · Scarf |
| N | Facial Features | Mustache · Beard · Dimples · Facial scars · Facial moles · Acne · Face/neck tattoos |
| O | Accessories and jewellery | Makeup · Necklace · Earrings · Nose piercing · Lip piercing · Eyebrow piercing |
| P | Others | Freckles · Wrinkles · Bindi · Other tattoos · Other piercings · Other - not specified |

A cell may hold **several values separated by commas or semicolons** — each is checked independently and the
cell is rewritten with a consistent `, ` separator.

Matching ignores case, punctuation and spacing, then falls back to a Damerau–Levenshtein comparison so a
typo of one or two characters still resolves. Real examples from `sample-files/`:

| Submitted | Result |
|-----------|--------|
| `Glases` | → `Glasses` (auto-fixed) |
| `Dimple` | → `Dimples` (auto-fixed) |
| `Moustache` | → `Mustache` (auto-fixed — British spelling) |
| `Freckels` | → `Freckles` (auto-fixed — transposed letters) |
| `DIMPLES` | → `Dimples` (auto-fixed — case) |
| `Makeup; Earrings` | → `Makeup, Earrings` (auto-fixed — separator) |
| `Sombrero` | kept as-is, **flagged for review** |
| `Glasses; Sombrero` | → `Glasses, Sombrero` — the good half is fixed, the cell is still **flagged** |

The tolerance scales with word length (1 edit under 5 characters, 2 under 9, 3 above), and if two vocabulary
entries are *equally* close the value is flagged rather than guessed. Nothing off-list is ever silently
replaced or deleted.

To change the vocabulary, edit `ATTRIBUTE_VOCABULARY` near the top of `js/core.js`.

### Severity levels

| Level | Meaning |
|-------|---------|
| `AUTO-FIXED` | Corrected for you. Informational. |
| `NEEDS REVIEW` | Could not be fixed safely. Original value preserved. |
| `ERROR` | Structural problem with the file, e.g. a missing column. |

### Highlighted cells

Every `NEEDS REVIEW` issue also **shades the offending cell amber in the merged workbook** and attaches the
reason as a cell comment, so you can hover it in Excel instead of cross-referencing the QA report.

This means a clean run produces **no highlighting at all** — any colour in the delivery sheet is a signal it
is not ready to send yet. Fix the source files, re-run, and the highlights disappear.

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

## Metadata Link Filler

A second tool in the same app — switch with **Merge / Link filler** in the nav. It fills the three Dropbox
link columns on an already-merged Centaurus sheet by finding each participant's files and resolving a shared
link for each.

### It reads a different sheet from the merger

This is the wide Centaurus sheet where **F–AB are already populated upstream**, not the 21-column output of
the merger. Fixed positions, per the spec:

| Col | Role | |
|-----|------|---|
| D | Participant / folder name | read |
| **E** | Video URL — one per row | **written** |
| J | Expressions (`Neutral` / `Non-Neutral`) | read |
| K | For Non Neutral, Please Select | read, disambiguates |
| Q | Environment (`Indoor` / `Outdoor`) | read |
| **AC** | ICF URL — repeats per participant | **written** |
| **AD** | Assent URL — repeats per participant | **written** |

**F–AB are never written.** The config strip above the scan button shows the target sheet and the fill columns
so a wrong target is visible before you run anything.

### How it matches

Participant identity comes from the **folder name**, not the filename. Within a participant's folder, files are
classified by the agency preset (Powerling or Aqlama), then each video's Environment and Expression are read
off its filename and matched to the row wanting that combination.

- **Name join** normalizes case and separators, then matches **exactly**. `Ronald Okoth` and `Ronald Okothh`
  are different people and never merge.
- **Expression** tolerates typos — `N3UTRAL` still reads as Neutral — reusing the merger's edit-distance.
- **Column K breaks ties.** When two videos both map to (Indoor, Non-Neutral), K names which one this row
  wants. If K doesn't single one out, the row is flagged rather than guessed.

### The five guards

Each holds its rows out of the write set and lists them in the review panel:

| Guard | Behaviour |
|-------|-----------|
| Missing assent | AD left blank, flagged. **Never borrowed from another participant.** |
| Missing consent | AC left blank, flagged. |
| Ambiguous slot | 2+ videos map to one (Env, Expression) and K doesn't resolve it. |
| Unmatched name | A folder normalizes to no sheet row. Reported, never forced onto a near miss. |
| Duplicate file | 2+ consent or assent files (`ICF 1.pdf`, `ICF 2.pdf`). Held as ambiguous. |

### Nothing is created without confirmation

A scan is **read-only**: it lists files and reuses shared links that already exist. Creating a link publishes
an *anyone-with-the-link* URL to a participant video or consent form, so it happens only when you click
**Create N missing links** and confirm the count.

Row states are therefore distinct:

- **Ready** — matched and all three links exist. Written on download.
- **Needs link** — matched cleanly, but a file has no shared link yet. Not written, because writing would
  blank the cell. Resolved by the create step.
- **Review / Missing** — a guard tripped. Fix the source and re-run.

Re-running is safe: existing links are reused, so a second pass only fills what changed and never creates
duplicate shares. Links are stored in the durable form `scl/fi/<id>?rlkey=<key>&dl=0` — the ephemeral `st=`
token is dropped because it expires and isn't reproducible.

### Dropbox setup

Needs a one-time Dropbox app, same pattern as the Azure one — add its **App key** under Settings or in
`config.js`. Create the app at <https://www.dropbox.com/developers/apps>, choose **Scoped access → Full
Dropbox**, and add the redirect URI shown in Settings. It is a public client using PKCE, so there is **no app
secret**.

On the **Permissions** tab tick all four, then click **Submit**:

| Scope | Used by |
|---|---|
| `account_info.read` | `users/get_current_account` — the connection check, and `root_info` for addressing team folders |
| `files.metadata.read` | `files/list_folder` — browsing and enumerating |
| `files.content.read` | `files/download` — reading the agency metadata workbooks. Only `.xlsx` files are fetched; video and consent files are linked, never downloaded |
| `sharing.read` | `sharing/list_shared_links` — reusing links that already exist |
| `sharing.write` | `sharing/create_shared_link_with_settings` — only after you confirm |

⚠️ **A permission added after you connect does not apply to an existing sign-in.** An access token carries the
scopes it was issued with, and refreshing it does not upgrade them. If calls fail with
`missing_scope`, tick the permission, Submit, then **disconnect and connect again**. The app detects this
error, clears the stale token for you and offers a reconnect button.

`Browse Dropbox` always opens at `dropboxBrowseRoot` in `config.js` and will not navigate above it. That value
accepts either a plain path or a Dropbox web URL, so it can be pasted straight from the address bar.

**Each person signs in with their own Dropbox account, once per device.** The tool requests an offline refresh
token, so after the first connect it stays connected — no repeated sign-ins. There is deliberately no shared
credential: a token embedded in this app would be readable by anyone visiting the public URL, and Dropbox's
audit log would attribute every created link to one person instead of whoever actually made it. **Disconnect**
clears the stored token, which matters on a shared machine.

### Tests

The matching engine is pure and has a regression suite covering every edge case in the spec — uppercase
`.MOV`, near-duplicate names, lowercase leading names, duplicate `ICF 1.pdf`, a stray `.xlsx` beside the
media, expression typos, all five guards and both agency presets:

```bash
node tests/linkfiller.test.js
```

To change the classification rules, edit `AGENCY_PRESETS` at the top of `js/linkfiller.js`.

**Not yet exercised against the live Dropbox API** — no app key is registered. The matching, guards, writing
and UI are all verified; the HTTP calls in `js/dropbox.js` follow the v2 docs but expect to verify team-folder
path roots and pagination on the first real run.

---

## Layout

```
index.html                  markup for both tools
403.html                    shown to signed-in users who lack the merger role
config.js                   optional Azure client/tenant ID and Dropbox app key defaults
staticwebapp.config.json    Azure auth rules, routing and security headers
css/styles.css              Centific design system tokens and components
js/core.js                  schema, QA rules, merge, workbook building  ← merger logic
js/linkfiller.js            matching, guards, link normalization        ← link filler logic
js/sharepoint.js            MSAL sign-in and Microsoft Graph download
js/dropbox.js               Dropbox PKCE sign-in, listing, shared links
js/app.js                   merger UI wiring
js/linkfillerui.js          link filler UI wiring
tests/linkfiller.test.js    matching engine regression suite (node, no deps)
vendor/                     ExcelJS and MSAL, vendored so there are no CDN calls
sample-files/               three agency files with deliberate data problems
.github/workflows/          Azure Static Web Apps deployment
```

To change a QA rule, edit `js/core.js` — the canonical header list, the column constants and each rule are all
at the top of that file.

Supported input formats are `.xlsx` and `.xlsm`. Older `.xls` and `.csv` files must be re-saved as `.xlsx`
first; the app reports this clearly rather than failing silently.
