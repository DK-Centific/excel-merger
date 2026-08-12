# How to use Merge

This tool takes the agency spreadsheets sitting in Dropbox, combines them into one sheet,
checks every row, and fills in the video, consent and assent links.

**Open it:** <https://dk-centific.github.io/excel-merger/>

---

## First time only

Click **Connect Dropbox** and sign in.

That's it. It remembers you on that computer, so you won't be asked again.

You'll only ever see folders your own Dropbox account can already open — signing in here
doesn't give you access to anything new.

---

## Every time

**1. Pick your folders.**
Click **Browse Dropbox** and tick the batch folders you want. Tick as many as you like —
they don't have to be next to each other. (Prefer pasting links? Use the **Paste URLs** tab.)

**2. Say what you want.**
Under *Merge, check and map links*:

- **Merged sheet with mapped links** — the normal choice. Combines the spreadsheets and fills in the Dropbox links.
- **Merged sheet only** — just combines and checks. No links.

**3. Click the button** and give it a moment. It's reading every spreadsheet in those folders.

**4. Click Download filled sheet.**

Done. The file lands in your Downloads folder as `Result_Merged.xlsx`.

---

## Reading what came back

Four numbers appear at the top:

| | |
|---|---|
| **Rows merged** | How many rows came out, and from how many spreadsheets |
| **QA auto-fixes** | Things it tidied for you. No action needed |
| **Links** | How many were already there vs. how many are missing |
| **Need a human** | The only number you have to care about |

In the sheet itself, **every row tells you how it went**:

- **Column A** — a number per participant
- **Column B** — `OK`, `Auto-fixed`, `Review` or `Error`
- **Column C** — if something needs attention, this says exactly what

So you can filter column B by `Review` and see only the rows that need you. Anything
amber in the sheet is the same signal.

There's also a **QA Report** tab at the back of the workbook with the full detail.

---

## When a row says "Review"

It means the tool wasn't confident, so it **left your data alone rather than guessing**.
Column C tells you why. Usually one of:

- **No assent file found** — that participant's folder doesn't have one. It won't borrow someone else's.
- **Two videos, one slot** — two clips match the same Environment and Expression. Fill in
  the *For Non Neutral, Please Select* column and it'll pick the right one.
- **Folder name didn't match a sheet row** — the folder is spelled differently from the
  participant's name. Fix either one.
- **Birthdate needs a manual fix** — it couldn't read the date, so it kept what was there.

Fix the source, run it again. Nothing is ever silently changed or dropped.

---

## Creating missing links

If some files don't have a Dropbox link yet, you'll see a **Create N missing links** button.

Nothing is created until you click it and confirm. Worth knowing why: a Dropbox link is a
*anyone-with-the-link* URL to a participant's video or consent form, so you should see the
count before it happens.

Re-running later is safe — existing links are reused, never duplicated.

---

## Handy: grab the spreadsheets without hunting

Once you've picked folders, **Find the Excel files** lists every spreadsheet in them, and
**Download all as .zip** pulls them down in one go. Much faster than scrolling past a
thousand video files in Dropbox.

---

## If something goes wrong

| What you see | What to do |
|---|---|
| *"This app has reached its user limit"* | Not something you can fix — the Dropbox app needs approving for production. Tell whoever set it up |
| *"This sign-in is missing a permission"* | Someone needs to tick it in the Dropbox app settings, then you click **Disconnect** and connect again |
| *"No Excel files were found"* | Check you ticked the right folders — it looks inside subfolders too |
| *"None with recognisable headers"* | Click **Find the Excel files**, download one, and check its header row |
| Nothing downloads | Check your browser didn't block it — look for a bar at the top of the window |

Still stuck? Send the message it showed you — they're written to say what actually happened.
