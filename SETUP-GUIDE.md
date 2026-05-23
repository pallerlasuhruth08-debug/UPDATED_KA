# Isha Karnataka — Backend setup (about 5 minutes)

This connects the app to a Google Sheet so everyone shares the same data, and adds per-user logins. You only do this once.

You have two files:
- `centre-os-prototype.html` — the app
- `Code.gs` — the backend code you'll paste into Google

---

## Step 1 — Create the Google Sheet
1. Go to https://sheets.google.com and create a new blank spreadsheet.
2. Name it something like **Isha Karnataka — Data**.
   (You don't need to add any tabs or headings — the script builds them for you.)

## Step 2 — Open the script editor
1. In that Sheet, click **Extensions → Apps Script**.
2. A code editor opens in a new tab with a file called `Code.gs` and an empty `function myFunction() {}`.
3. Delete everything in that editor.
4. Open the `Code.gs` file I gave you, copy **all** of it, and paste it into the editor.
5. Click the **Save** icon (or press Ctrl/Cmd + S).

## Step 3 — Deploy it as a web app
1. Top right, click **Deploy → New deployment**.
2. Click the gear icon next to "Select type" and choose **Web app**.
3. Fill in:
   - **Description:** Isha Karnataka backend (anything is fine)
   - **Execute as:** **Me** (your Google account)
   - **Who has access:** **Anyone**   ← important, so teammates don't each need a Google login
4. Click **Deploy**.
5. Google will ask you to **Authorize access** — approve it with your account.
   (You may see a "Google hasn't verified this app" screen — click **Advanced → Go to … (unsafe)**. This is normal for your own scripts.)
6. Copy the **Web app URL**. It looks like:
   `https://script.google.com/macros/s/AKfy....../exec`

## Step 4 — Connect the app
1. Open `centre-os-prototype.html` in your browser.
2. On the first screen, paste the **Web app URL** and click **Connect**.
3. Log in with the starter account:
   - **Username:** `admin`
   - **Password:** `isha@2026`
4. You're in. The app pushes the current Karnataka data into your Sheet automatically on first login.

---

## Add your team's logins
1. Go back to the Google Sheet — you'll now see a **Users** tab.
2. Each row is one login: `username | password | name | active`
   - Example: `priya | priya123 | Priya N. | yes`
   - Set `active` to `no` to disable someone without deleting them.
3. Change the **admin** password from `isha@2026` to something private.
4. New users can log in immediately — no redeploy needed.

## Where the data lives
- The whole centre tree is stored as one JSON cell in the **Data** tab (`A1`). You normally won't touch it — the app reads and writes it. (If you ever want a clean reset, clear `A1` and the app re-seeds on next login.)
- The **Sessions** tab is managed automatically (login tokens). You can ignore it.

## Updating the app later
If I send you a new version of the HTML, just replace the file — your Sheet, users, and data are untouched. If I change `Code.gs`, paste the new version into the editor and do **Deploy → Manage deployments → Edit (pencil) → New version → Deploy** (this keeps the same URL).

---

## A note on security (please read)
This is a lightweight access gate suitable for an internal team tool, not a bank-grade login:
- Passwords are stored as plain text in the Users tab, so **don't reuse important passwords** here, and limit who can open the Sheet (it's private to people you share it with).
- "Who has access: Anyone" means anyone with the **web app URL** can reach the login screen, but they still need a valid username + password to get any data.
- The password is checked on Google's server, not inside the HTML, so it isn't exposed in the page source.

When this graduates from prototype to a real product, the natural next step is a proper backend with hashed passwords. I can help with that when you're ready.

## If something doesn't work
- **"Couldn't reach the web app"** → re-check Step 3: access must be **Anyone**, and you must use the `/exec` URL (not the `/dev` one).
- **Login fails for a new user** → check the Users tab spelling and that `active` is `yes`.
- **Changed Code.gs but nothing changed** → you must redeploy a **New version** (Manage deployments → Edit → New version).
