# US Vacation Expense Tracker — Samit Reimbursement

A Google Sheets database + Google Apps Script web app for tracking US vacation
expenses that Samit reimburses monthly. Built for 3 people to use together.

**The spreadsheet:**
https://docs.google.com/spreadsheets/d/198cFEHC7fvbEZHlBBjXBSGdU9G5rI0t8YDkOqlKVaDI/edit

> **Where does this code actually run?** This GitHub folder is just the
> source-of-record/backup. The live copies of `Code.gs` and `App.html` live
> inside the Apps Script project attached to the spreadsheet (Extensions →
> Apps Script), and Google serves the app from a `script.google.com` URL.
> It cannot be hosted on GitHub Pages — the interface talks to the sheet
> through `google.script.run`, which only works when Google serves the page.

## What it tracks

Categories: Hotels, Plane Tickets, Disney Tickets, Universal Tickets,
Restaurants / Food, Las Vegas Show Tickets, Other.

Each expense records: date, category, description, amount (USD), who paid,
which month Samit reimburses it (defaults to the month **after** the expense),
status (Pending / Reimbursed), notes, an optional **receipt photo**, who
entered it, and timestamps.

Receipt photos: the app opens the camera on phones (or a file picker on
desktop for screenshots), compresses the image in the browser, saves it to a
Drive folder called **"Vacation Receipts (Samit Reimbursement)"**, and links
it on the expense row (📎 Receipt) and in the sheet's Receipt column.

**Automatic receipt reading (optional):** with an Anthropic API key set, the
app reads each photo with Claude's vision API and pre-fills the total, date,
merchant and category — you review and hit save. Everything it fills in stays
editable, and it flags low-confidence reads so you know when to double-check.
Without a key the photo still attaches; you just type the amount yourself.

To enable it: get a key at [console.anthropic.com](https://console.anthropic.com),
then in Apps Script go to **⚙️ Project Settings → Script Properties → Add
script property**, name it exactly `ANTHROPIC_API_KEY`, and paste the key as
the value. Reading a receipt costs roughly 1–3 cents.

Tabs:
- **Expenses** — the database, with dropdowns and color-coding by status
- **Monthly Summary** — auto-updating totals: owed by month, pending by month,
  by category, by payer, and grand totals (spent / reimbursed / still owed)
- **Settings** — edit the 3 people's names and the category list here

## One-time setup (~5 minutes)

1. Open the spreadsheet link above.
2. **Extensions → Apps Script**. Delete any starter code.
3. Create two files matching the ones in this folder:
   - Paste `Code.gs` into the default `Code.gs` file.
   - **File → New → HTML**, name it exactly `App`, paste in `App.html`.
4. Save, then run the `setupSpreadsheet` function once (select it in the
   toolbar dropdown and press ▶ Run). Approve the permissions prompt.
   This builds all three tabs, dropdowns, and the summary formulas.
5. Go to the **Settings** tab in the sheet and replace *Person 1/2/3* with
   your three real names.

## Deploy the shared web app

1. In Apps Script: **Deploy → New deployment → type: Web app**.
2. *Execute as:* **Me** (so the other two don't need edit access to the sheet).
3. *Who has access:* **Anyone with a Google account** — or share the sheet
   with your two companions and choose *Anyone with the link* / restrict as
   you prefer.
4. Click **Deploy** and share the web app URL with the other 2 people.

Each person picks their name from the "I am" menu in the header (remembered on
their device), then can add expenses with receipt photos, filter, mark
individual expenses paid, or hit **"💸 Samit paid this month"** to clear a
whole month at once.

Note the web app runs as you, so receipt reading uses your API key no matter
which of the three people snaps the photo — you don't need to hand the key out.

## Updating the code later

After editing code in Apps Script, use **Deploy → Manage deployments → ✏️ →
New version** so the existing URL picks up the changes.
