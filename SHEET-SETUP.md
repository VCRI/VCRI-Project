# Google Sheet — User Data Tracking (setup guide)

This logs every report generation to a Google Sheet so the team can see
**who generated a report, their mobile, date & time, the details they entered,
and how many times each user generated a report** — and search any user later.

Two tabs are created automatically the first time data arrives:

| Tab | What it holds |
|-----|----------------|
| **Reports** | One row **per generation** — Timestamp, Date, Time, Name, Mobile, Email, Generation #, Animal Type, Breed, No. of Animals, Project Cost, Loan, NPV, BCR, IRR, Avg DSCR, and the full **Details Entered (JSON)**. |
| **Users** | One row **per mobile number** — Name, Email, **Total Reports**, First Generated, Last Generated. (Quick per-user view.) |

---

## Step 1 — Create the Sheet
1. Go to <https://sheets.google.com> and create a **blank spreadsheet**.
2. Name it e.g. **"Dairy Report — User Log"**. (You don't need to add any columns; the script makes them.)

## Step 2 — Add the script
1. In the sheet, open **Extensions ▸ Apps Script**.
2. Delete any sample code, then **paste the entire contents of `google-apps-script.gs`**.
3. Click the **Save** icon (💾).

## Step 3 — Deploy as a Web App
1. Click **Deploy ▸ New deployment**.
2. Click the gear ⚙ next to "Select type" → choose **Web app**.
3. Set:
   - **Description:** Dairy report logger
   - **Execute as:** **Me**
   - **Who has access:** **Anyone**  *(required so the website can post to it)*
4. Click **Deploy**, then **Authorize access** and allow the permissions (it's your own script).
5. Copy the **Web app URL** — it looks like:
   `https://script.google.com/macros/s/AKfy................/exec`

## Step 4 — Connect the website
1. Open **`report.html`** in a text editor.
2. Find this line near the bottom:
   ```js
   const TRACK_URL = "";
   ```
3. Paste your Web app URL inside the quotes:
   ```js
   const TRACK_URL = "https://script.google.com/macros/s/AKfy...../exec";
   ```
4. Save. Done — every time someone generates a report, a row is added.

---

## How the team uses it
- **See all activity:** open the **Reports** tab — newest rows at the bottom.
- **Find a specific user:** use Google Sheets **Find** (Ctrl/Cmd + F) or a column filter on
  **Name** or **Mobile**. Every report that user generated will be listed, with date, time
  and the exact details they entered.
- **How many times a user generated:** the **Generation #** column shows the running count for
  that mobile; the **Users** tab shows the **Total Reports** per person at a glance.
- **What they entered:** the **Details Entered (JSON)** column holds every parameter from the form
  for that generation.

## Notes
- **Privacy:** data is sent only to *your* Google Sheet; nothing else leaves the browser.
- **Re-deploying after edits:** if you change the `.gs` code later, use **Deploy ▸ Manage deployments ▸
  Edit ▸ Version: New version** so the URL stays the same.
- **Test quickly:** pasting the Web app URL straight into a browser runs `doGet` and should
  return `{"ok":true,...}` and add a (mostly blank) test row.
- **Counting:** a generation = one click of **Generate Report**. Refreshing the report page does
  **not** add a duplicate (it's de-duplicated per generation token).
