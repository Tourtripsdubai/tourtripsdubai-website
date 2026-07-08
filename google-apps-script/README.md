# Contact Form → Google Sheet + Email setup

Follow these steps once. After this, every contact form submission on the site will:
1. Add a row to a Google Sheet
2. Email **you** (`ADMIN_EMAIL`, currently `tourtripsdubai@gmail.com`) the customer's details
3. Email the **customer** an auto-reply confirmation, sent from `FROM_EMAIL` (currently `info@tourtripsdubai.com`)

## Steps

1. Go to [sheets.google.com](https://sheets.google.com) and create a new blank sheet. Name it e.g. "TourTripsDubai Enquiries".
2. In the sheet, go to **Extensions → Apps Script**.
3. Delete the placeholder code in `Code.gs`, and paste the entire contents of [`Code.gs`](Code.gs) from this folder instead.
4. At the top of the script, confirm `ADMIN_EMAIL` and `FROM_EMAIL` are set the way you want.
5. **Add `info@tourtripsdubai.com` as a verified "Send mail as" alias** on the Google account you're about to authorize this script with (the one you'll click "Allow" with in step 7) — otherwise sending "from" that address will fail:
   - In that Google account, open Gmail → ⚙️ **Settings → See all settings → Accounts and Import → Send mail as → Add another email address**.
   - Name: `TourTripsDubai`, Email: `info@tourtripsdubai.com`. Uncheck "Treat as an alias".
   - Choose **Send through SMTP servers** and enter the SMTP host/port/username(`info@tourtripsdubai.com`)/password from your hosting's email settings (in cPanel: Email Accounts → Connect Devices/Configure Mail Client, for the `info@` mailbox). This keeps deliverability good since it authenticates as the real mailbox.
   - Gmail will verify by sending a confirmation to `info@tourtripsdubai.com` — open that inbox and confirm.
6. Click **Deploy → New deployment**.
   - Click the gear icon next to "Select type" → choose **Web app**.
   - Description: anything, e.g. "Contact form handler".
   - Execute as: **Me**.
   - Who has access: **Anyone**.
   - Click **Deploy**.
7. Google will ask you to authorize the script (it needs permission to send email as `info@tourtripsdubai.com` and edit the sheet) — click through **Authorize access**, choose your account, click **Advanced → Go to (project name)** if it warns it's unverified, then **Allow**.
8. Copy the **Web app URL** it gives you (ends in `/exec`).
9. Open `index.html` in the site, find this line near the bottom (search for `CONTACT_SCRIPT_URL`):
   ```js
   const CONTACT_SCRIPT_URL = 'PASTE_YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE';
   ```
   Replace the placeholder text with the URL you copied, keeping the quotes.
10. Save and re-upload `index.html`. Test the form — check that a row appears in the sheet and both emails arrive (they may land in Spam the first time; mark as "not spam").

## If you ever edit the contact form

The script matches form fields by their `<label>` text, so it already understands both forms on the site (homepage quick form and the Contact page form). If you add a brand-new field with a new label, it will still land in the Sheet automatically — you just won't have a dedicated column for it unless you add one in `appendToSheet` in `Code.gs`.

## Re-deploying after editing Code.gs

If you ever change `Code.gs`, you must create a **New deployment** again (or use "Manage deployments → Edit → New version") for the changes to take effect — editing the script alone does not update a live deployment.
