# Contact Form → Google Sheet + Email setup

Follow these steps once. After this, every contact form submission on the site will:
1. Add a row to a Google Sheet
2. Email **you** (info@tourtripsdubai.com) the enquiry
3. Email the **customer** an auto-reply confirmation

## Steps

1. Go to [sheets.google.com](https://sheets.google.com) and create a new blank sheet. Name it e.g. "TourTripsDubai Enquiries".
2. In the sheet, go to **Extensions → Apps Script**.
3. Delete the placeholder code in `Code.gs`, and paste the entire contents of [`Code.gs`](Code.gs) from this folder instead.
4. At the top of the script, confirm `ADMIN_EMAIL` is the email you want enquiries sent to (currently set to `info@tourtripsdubai.com`).
5. Click **Deploy → New deployment**.
   - Click the gear icon next to "Select type" → choose **Web app**.
   - Description: anything, e.g. "Contact form handler".
   - Execute as: **Me**.
   - Who has access: **Anyone**.
   - Click **Deploy**.
6. Google will ask you to authorize the script (it needs permission to send email and edit the sheet) — click through **Authorize access**, choose your account, click **Advanced → Go to (project name)** if it warns it's unverified, then **Allow**.
7. Copy the **Web app URL** it gives you (ends in `/exec`).
8. Open `index.html` in the site, find this line near the bottom (search for `CONTACT_SCRIPT_URL`):
   ```js
   const CONTACT_SCRIPT_URL = 'PASTE_YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE';
   ```
   Replace the placeholder text with the URL you copied, keeping the quotes.
9. Save and re-upload `index.html`. Test the form — check that a row appears in the sheet and both emails arrive (they may land in Spam the first time; mark as "not spam").

## If you ever edit the contact form

The script matches form fields by their `<label>` text, so it already understands both forms on the site (homepage quick form and the Contact page form). If you add a brand-new field with a new label, it will still land in the Sheet automatically — you just won't have a dedicated column for it unless you add one in `appendToSheet` in `Code.gs`.

## Re-deploying after editing Code.gs

If you ever change `Code.gs`, you must create a **New deployment** again (or use "Manage deployments → Edit → New version") for the changes to take effect — editing the script alone does not update a live deployment.
