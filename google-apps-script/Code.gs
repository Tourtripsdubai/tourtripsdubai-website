/**
 * TourTripsDubai — Contact Form backend.
 * Deploy this bound to a Google Sheet (Extensions > Apps Script), as a Web App.
 * See README.md in this folder for step-by-step setup.
 */

const SHEET_NAME = 'Enquiries';
const ADMIN_EMAIL = 'tourtripsdubai@gmail.com';
const FROM_EMAIL = 'info@tourtripsdubai.com';
const SITE_NAME = 'TourTripsDubai';
const SITE_URL = 'https://www.tourtripsdubai.com';
const ADMIN_PHONE_DISPLAY = '+971 58 914 7828';
const TIMEZONE = 'Asia/Dubai';

const ICON_INSTAGRAM = SITE_URL + '/assets/images/icons8-instagram-logo.svg';
const ICON_FACEBOOK = SITE_URL + '/assets/images/icons8-facebook.svg';
const ICON_YOUTUBE = SITE_URL + '/assets/images/icons8-youtube.svg';
const ICON_TRUSTPILOT = SITE_URL + '/assets/images/icons8-trustpilot.svg';

const NAVY = '#081120';
const NAVY_2 = '#0F3A68';
const ORANGE = '#FE5A08';
const BG = '#F5F7FA';
const INK = '#1A2333';
const INK_SOFT = '#6B7688';
const BORDER = '#E7EAF0';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const row = appendToSheet(data);
    const leadId = 'TTD-' + Utilities.formatDate(new Date(), TIMEZONE, 'yyyy') + '-' + String(row - 1).padStart(5, '0');
    notifyAdmin(data, leadId);
    notifyCustomer(data, leadId);
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Timestamp', 'Page', 'Full Name', 'Email Address', 'WhatsApp Number',
      'Nationality', 'Travel Date', 'Package / Service', 'Number of Adults',
      'Number of Children', 'Special Requests / Message']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function findValue(data, keys) {
  for (const k of keys) {
    if (data[k]) return data[k];
  }
  return '';
}

function appendToSheet(data) {
  const sheet = getSheet();
  sheet.appendRow([
    new Date(),
    data['Page'] || '',
    findValue(data, ['Full Name', 'Full Name (As Per Passport)']),
    findValue(data, ['Email Address', 'Email']),
    findValue(data, ['WhatsApp Number']),
    findValue(data, ['Nationality']),
    findValue(data, ['Travel Date']),
    findValue(data, ['Package / Service']),
    findValue(data, ['Number of Adults']),
    findValue(data, ['Number of Children']),
    findValue(data, ['Special Requests', 'Special Requests / Message', 'Message'])
  ]);
  return sheet.getLastRow();
}

function formatSummary(data) {
  const skip = ['Page', 'Submitted At'];
  return Object.keys(data)
    .filter(k => !skip.includes(k) && data[k])
    .map(k => k + ': ' + data[k])
    .join('\n');
}

function escapeHtml(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function notifyAdmin(data, leadId) {
  const name = findValue(data, ['Full Name', 'Full Name (As Per Passport)']) || 'Website Visitor';
  GmailApp.sendEmail(ADMIN_EMAIL, 'New Enquiry — ' + name + ' (' + leadId + ')',
    'You have a new enquiry from ' + SITE_NAME + ':\n\n' + formatSummary(data),
    { from: FROM_EMAIL, name: SITE_NAME, htmlBody: buildAdminEmailHtml(data, leadId) }
  );
}

function notifyCustomer(data, leadId) {
  const email = findValue(data, ['Email Address', 'Email']);
  if (!email) return;
  const name = findValue(data, ['Full Name', 'Full Name (As Per Passport)']);
  GmailApp.sendEmail(email, 'We received your enquiry — ' + SITE_NAME,
    'Hi ' + (name || 'there') + ',\n\n' +
      'Thank you for reaching out to ' + SITE_NAME + '. We have received your enquiry ' +
      'and one of our travel experts will get back to you shortly.\n\n' +
      'Here is a summary of what you submitted:\n\n' + formatSummary(data) + '\n\n' +
      'Warm regards,\n' + SITE_NAME + ' Team',
    { from: FROM_EMAIL, name: SITE_NAME, htmlBody: buildCustomerEmailHtml(data, leadId) }
  );
}

/* ===================== Shared email building blocks ===================== */

function row(label, value) {
  if (!value) return '';
  return '<td width="50%" valign="top" style="padding:0 10px 22px 0;">' +
    '<div style="font-family:\'Work Sans\',Arial,sans-serif;font-size:11px;letter-spacing:.08em;text-transform:uppercase;color:' + INK_SOFT + ';font-weight:600;margin-bottom:5px;">' + label + '</div>' +
    '<div style="font-family:\'Work Sans\',Arial,sans-serif;font-size:15px;color:' + INK + ';font-weight:600;">' + value + '</div>' +
    '</td>';
}

function fullRow(label, value) {
  if (!value) return '';
  return '<tr><td colspan="2" style="padding:0 0 22px 0;">' +
    '<div style="font-family:\'Work Sans\',Arial,sans-serif;font-size:11px;letter-spacing:.08em;text-transform:uppercase;color:' + INK_SOFT + ';font-weight:600;margin-bottom:5px;">' + label + '</div>' +
    '<div style="font-family:\'Work Sans\',Arial,sans-serif;font-size:15px;color:' + INK + ';font-weight:500;line-height:1.55;">' + value + '</div>' +
    '</td></tr>';
}

function ctaButton(href, bg, color, label) {
  if (!href) return '';
  return '<td style="padding:0 6px 12px 6px;" align="center">' +
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0"><tr>' +
    '<td align="center" style="border-radius:8px;background:' + bg + ';">' +
    '<a href="' + href + '" target="_blank" style="display:inline-block;padding:13px 20px;font-family:\'Work Sans\',Arial,sans-serif;font-size:13px;font-weight:700;letter-spacing:.03em;color:' + color + ';text-decoration:none;">' + label + '</a>' +
    '</td></tr></table></td>';
}

function socialIcon(href, iconUrl) {
  return '<td style="padding:0 6px;">' +
    '<a href="' + href + '" target="_blank" style="display:inline-block;width:36px;height:36px;">' +
    '<img src="' + iconUrl + '" width="36" height="36" alt="" style="display:block;border-radius:50%;">' +
    '</a></td>';
}

function socialIconRow() {
  return socialIcon('https://www.instagram.com/tourtripsdubai', ICON_INSTAGRAM) +
    socialIcon('https://www.facebook.com/share/19GzNfyX7z/', ICON_FACEBOOK) +
    socialIcon('https://www.youtube.com/@TourTripsDubai', ICON_YOUTUBE) +
    socialIcon('https://www.trustpilot.com/evaluate/tourtripsdubai.com', ICON_TRUSTPILOT);
}

/* ===================== Admin notification email (HTML) ===================== */

function buildAdminEmailHtml(data, leadId) {
  const fullName = escapeHtml(findValue(data, ['Full Name', 'Full Name (As Per Passport)'])) || 'Website Visitor';
  const email = escapeHtml(findValue(data, ['Email Address', 'Email']));
  const whatsapp = findValue(data, ['WhatsApp Number']);
  const whatsappSafe = escapeHtml(whatsapp);
  const nationality = escapeHtml(findValue(data, ['Nationality']));
  const travelDate = escapeHtml(findValue(data, ['Travel Date']));
  const pkg = escapeHtml(findValue(data, ['Package / Service']));
  const adults = escapeHtml(findValue(data, ['Number of Adults']));
  const children = escapeHtml(findValue(data, ['Number of Children']));
  const message = escapeHtml(findValue(data, ['Special Requests', 'Special Requests / Message', 'Message']));
  const page = escapeHtml(data['Page'] || 'website');
  const submittedAt = Utilities.formatDate(new Date(), TIMEZONE, 'dd MMM yyyy, hh:mm a') + ' (GST)';
  const preferredContact = whatsapp ? 'WhatsApp' : (email ? 'Email' : '—');

  const phoneDigits = whatsapp.replace(/[^\d]/g, '');
  const telHref = phoneDigits ? 'tel:+' + phoneDigits : '';
  const waHref = phoneDigits ? 'https://wa.me/' + phoneDigits + '?text=' + encodeURIComponent('Hi ' + fullName + ', thank you for your enquiry with TourTripsDubai! ') : '';
  const mailHref = email ? 'mailto:' + email + '?subject=' + encodeURIComponent('Re: Your Dubai Trip Enquiry — ' + SITE_NAME) : '';

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta name="color-scheme" content="light">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Playfair+Display:ital,wght@0,400;0,500;0,600;0,700;0,800;1,500&family=Work+Sans:wght@300;400;500;600;700&display=swap" rel="stylesheet">
<title>New Enquiry — ${SITE_NAME}</title>
<style>
  @media only screen and (max-width:640px){
    .ttd-wrap{width:100% !important;}
    .ttd-pad{padding-left:20px !important;padding-right:20px !important;}
    .ttd-stack td{display:block !important;width:100% !important;padding-right:0 !important;}
    .ttd-btn td{display:block !important;width:100% !important;padding:0 0 10px 0 !important;}
    .ttd-btn td a{display:block !important;}
  }
</style>
</head>
<body style="margin:0;padding:0;background:${BG};">
<div style="display:none;max-height:0;overflow:hidden;opacity:0;">New booking enquiry from ${fullName} — ${pkg || 'Dubai trip'} — ${leadId}</div>
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background:${BG};">
<tr><td align="center" style="padding:30px 16px;">

<table role="presentation" class="ttd-wrap" width="650" cellpadding="0" cellspacing="0" border="0" style="width:650px;max-width:650px;background:#ffffff;border-radius:15px;overflow:hidden;box-shadow:0 20px 50px -20px rgba(8,17,32,.25);">

  <!-- HEADER -->
  <tr><td class="ttd-pad" style="background:${NAVY};background:linear-gradient(135deg,${NAVY},${NAVY_2});padding:40px 40px 34px 40px;text-align:center;">
    <div style="font-family:'Playfair Display',Georgia,serif;font-size:24px;font-weight:700;color:#ffffff;letter-spacing:.02em;">
      Tour<span style="color:${ORANGE};">Trips</span>Dubai
    </div>
    <div style="font-family:'Work Sans',Arial,sans-serif;font-size:10px;letter-spacing:.35em;text-transform:uppercase;color:${ORANGE};font-weight:600;margin-top:6px;">Premium Dubai Experiences</div>
    <div style="height:1px;background:rgba(255,255,255,.15);margin:26px 0;"></div>
    <div style="font-family:'Playfair Display',Georgia,serif;font-size:26px;line-height:1.3;color:#ffffff;font-weight:700;">&#128640; New Booking Enquiry Received</div>
    <div style="font-family:'Work Sans',Arial,sans-serif;font-size:14px;color:rgba(255,255,255,.7);margin-top:10px;">A new customer has submitted an enquiry from the website.</div>
  </td></tr>

  <!-- BODY -->
  <tr><td class="ttd-pad" style="padding:34px 40px 10px 40px;background:${BG};">

    <!-- Status badges -->
    <table role="presentation" cellpadding="0" cellspacing="0" border="0" align="center" style="margin:0 auto 24px auto;">
      <tr>
        <td style="padding:0 5px;"><div style="background:${ORANGE};color:#ffffff;font-family:'Work Sans',Arial,sans-serif;font-size:11px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;padding:8px 16px;border-radius:20px;">&#9679; New Lead</div></td>
        <td style="padding:0 5px;"><div style="background:#FEECE3;color:#C23B00;font-family:'Work Sans',Arial,sans-serif;font-size:11px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;padding:8px 16px;border-radius:20px;">High Priority</div></td>
        <td style="padding:0 5px;"><div style="background:#ffffff;border:1px solid ${BORDER};color:${NAVY_2};font-family:'Work Sans',Arial,sans-serif;font-size:11px;font-weight:700;letter-spacing:.03em;padding:8px 16px;border-radius:20px;">Website Contact Form · ${page}</div></td>
      </tr>
    </table>

    <!-- Lead ID bar -->
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background:${NAVY};border-radius:10px;margin-bottom:24px;">
      <tr>
        <td style="padding:14px 20px;font-family:'Work Sans',Arial,sans-serif;font-size:13px;color:#ffffff;">
          <strong style="color:${ORANGE};">Lead ID:</strong> ${leadId}
        </td>
        <td align="right" style="padding:14px 20px;font-family:'Work Sans',Arial,sans-serif;font-size:13px;color:rgba(255,255,255,.65);">
          Submitted: ${submittedAt}
        </td>
      </tr>
    </table>

    <!-- Customer Info Card -->
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#ffffff;border:1px solid ${BORDER};border-radius:15px;box-shadow:0 8px 24px -12px rgba(8,17,32,.12);margin-bottom:22px;">
      <tr><td style="padding:28px 28px 6px 28px;">
        <div style="font-family:'Playfair Display',Georgia,serif;font-size:18px;font-weight:700;color:${NAVY};margin-bottom:20px;">&#128100; Customer Information</div>
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" class="ttd-stack">
          <tr>
            ${row('Full Name', fullName)}
            ${row('Email Address', email || '—')}
          </tr>
          <tr>
            ${row('WhatsApp Number', whatsappSafe || '—')}
            ${row('Nationality', nationality || '—')}
          </tr>
          <tr>
            ${row('Travel Date', travelDate || '—')}
            ${row('Selected Package', pkg || '—')}
          </tr>
          <tr>
            ${row('Number of Adults', adults || '—')}
            ${row('Number of Children', children || '0')}
          </tr>
          ${fullRow('Special Request', message)}
        </table>
      </td></tr>
    </table>

    <!-- Quick Summary -->
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background:linear-gradient(135deg,${NAVY},${NAVY_2});border-radius:15px;margin-bottom:28px;">
      <tr><td style="padding:22px 24px;">
        <div style="font-family:'Work Sans',Arial,sans-serif;font-size:11px;letter-spacing:.1em;text-transform:uppercase;color:${ORANGE};font-weight:700;margin-bottom:16px;">Quick Summary</div>
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" class="ttd-stack">
          <tr>
            <td width="25%" style="padding:0 8px 0 0;">
              <div style="font-family:'Work Sans',Arial,sans-serif;font-size:10px;color:rgba(255,255,255,.55);text-transform:uppercase;letter-spacing:.05em;margin-bottom:4px;">Package</div>
              <div style="font-family:'Playfair Display',Georgia,serif;font-size:15px;color:#ffffff;font-weight:700;">${pkg || '—'}</div>
            </td>
            <td width="25%" style="padding:0 8px;">
              <div style="font-family:'Work Sans',Arial,sans-serif;font-size:10px;color:rgba(255,255,255,.55);text-transform:uppercase;letter-spacing:.05em;margin-bottom:4px;">Travel Date</div>
              <div style="font-family:'Playfair Display',Georgia,serif;font-size:15px;color:#ffffff;font-weight:700;">${travelDate || '—'}</div>
            </td>
            <td width="25%" style="padding:0 8px;">
              <div style="font-family:'Work Sans',Arial,sans-serif;font-size:10px;color:rgba(255,255,255,.55);text-transform:uppercase;letter-spacing:.05em;margin-bottom:4px;">Passengers</div>
              <div style="font-family:'Playfair Display',Georgia,serif;font-size:15px;color:#ffffff;font-weight:700;">${adults || '0'} Adult${adults === '1' ? '' : 's'}${children ? ', ' + children + ' Child' + (children === '1' ? '' : 'ren') : ''}</div>
            </td>
            <td width="25%" style="padding:0 0 0 8px;">
              <div style="font-family:'Work Sans',Arial,sans-serif;font-size:10px;color:rgba(255,255,255,.55);text-transform:uppercase;letter-spacing:.05em;margin-bottom:4px;">Preferred Contact</div>
              <div style="font-family:'Playfair Display',Georgia,serif;font-size:15px;color:#ffffff;font-weight:700;">${preferredContact}</div>
            </td>
          </tr>
        </table>
      </td></tr>
    </table>

    <!-- CTA Buttons -->
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" class="ttd-btn" style="margin-bottom:10px;">
      <tr>
        ${ctaButton(telHref, ORANGE, '#ffffff', '&#128222; Call Customer')}
        ${ctaButton(waHref, '#25D366', '#ffffff', '&#128172; WhatsApp Customer')}
      </tr>
      <tr>
        ${ctaButton(mailHref, NAVY, '#ffffff', '&#9993; Reply via Email')}
        ${ctaButton(SITE_URL, '#ffffff', NAVY, '&#127760; Open Website')}
      </tr>
    </table>

  </td></tr>

  <!-- FOOTER -->
  <tr><td class="ttd-pad" style="background:${NAVY};padding:36px 40px;text-align:center;">
    <div style="font-family:'Playfair Display',Georgia,serif;font-size:18px;font-weight:700;color:#ffffff;">TourTripsDubai</div>
    <div style="font-family:'Work Sans',Arial,sans-serif;font-size:12px;color:rgba(255,255,255,.55);margin-top:4px;">Premium Dubai Travel Company</div>

    <div style="height:1px;background:rgba(255,255,255,.12);margin:22px 0;"></div>

    <div style="font-family:'Work Sans',Arial,sans-serif;font-size:12px;color:rgba(255,255,255,.75);line-height:2;">
      <a href="${SITE_URL}" style="color:${ORANGE};text-decoration:none;">www.tourtripsdubai.com</a> &nbsp;·&nbsp;
      <a href="mailto:info@tourtripsdubai.com" style="color:rgba(255,255,255,.75);text-decoration:none;">info@tourtripsdubai.com</a> &nbsp;·&nbsp;
      <a href="tel:+971589147828" style="color:rgba(255,255,255,.75);text-decoration:none;">${ADMIN_PHONE_DISPLAY}</a>
    </div>

    <table role="presentation" cellpadding="0" cellspacing="0" border="0" align="center" style="margin:20px auto 0 auto;">
      <tr>
        ${socialIconRow()}
      </tr>
    </table>

    <div style="height:1px;background:rgba(255,255,255,.12);margin:22px 0;"></div>

    <div style="font-family:'Work Sans',Arial,sans-serif;font-size:11px;color:rgba(255,255,255,.4);line-height:1.6;">
      © ${new Date().getFullYear()} TourTripsDubai. All rights reserved.<br>
      This email was automatically generated by the TourTripsDubai website.
    </div>
  </td></tr>

</table>
</td></tr>
</table>
</body>
</html>`;
}

/* ===================== Customer confirmation email (HTML) ===================== */

function buildCustomerEmailHtml(data, leadId) {
  const fullName = escapeHtml(findValue(data, ['Full Name', 'Full Name (As Per Passport)'])) || 'Traveller';
  const firstName = fullName.split(' ')[0];
  const email = escapeHtml(findValue(data, ['Email Address', 'Email']));
  const whatsapp = findValue(data, ['WhatsApp Number']);
  const whatsappSafe = escapeHtml(whatsapp);
  const travelDate = escapeHtml(findValue(data, ['Travel Date']));
  const pkg = escapeHtml(findValue(data, ['Package / Service']));
  const adults = escapeHtml(findValue(data, ['Number of Adults']));
  const children = escapeHtml(findValue(data, ['Number of Children']));
  const message = escapeHtml(findValue(data, ['Special Requests', 'Special Requests / Message', 'Message']));

  const phoneDigits = whatsapp.replace(/[^\d]/g, '');
  const telHref = phoneDigits ? 'tel:+' + phoneDigits : 'tel:' + ADMIN_PHONE_DISPLAY.replace(/[^\d+]/g, '');
  const waHref = 'https://wa.me/' + ADMIN_PHONE_DISPLAY.replace(/[^\d]/g, '') + '?text=' + encodeURIComponent('Hi TourTripsDubai, my enquiry reference is ' + leadId + '.');

  function whyCard(icon, title, desc) {
    return `<td width="50%" valign="top" style="padding:0 10px 20px 0;">` +
      `<div style="background:#ffffff;border:1px solid ${BORDER};border-radius:12px;padding:20px;height:100%;">` +
      `<div style="font-size:26px;margin-bottom:10px;">${icon}</div>` +
      `<div style="font-family:'Playfair Display',Georgia,serif;font-size:15px;font-weight:700;color:${NAVY};margin-bottom:6px;">${title}</div>` +
      `<div style="font-family:'Work Sans',Arial,sans-serif;font-size:12.5px;color:${INK_SOFT};line-height:1.5;">${desc}</div>` +
      `</div></td>`;
  }

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta name="color-scheme" content="light">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Playfair+Display:ital,wght@0,400;0,500;0,600;0,700;0,800;1,500&family=Work+Sans:wght@300;400;500;600;700&display=swap" rel="stylesheet">
<title>Thank You — ${SITE_NAME}</title>
<style>
  @media only screen and (max-width:640px){
    .ttd-wrap{width:100% !important;}
    .ttd-pad{padding-left:20px !important;padding-right:20px !important;}
    .ttd-stack td{display:block !important;width:100% !important;padding-right:0 !important;}
    .ttd-btn td{display:block !important;width:100% !important;padding:0 0 10px 0 !important;}
    .ttd-btn td a{display:block !important;}
  }
</style>
</head>
<body style="margin:0;padding:0;background:${BG};">
<div style="display:none;max-height:0;overflow:hidden;opacity:0;">Thank you for your enquiry, ${firstName} — our Dubai travel expert will reach out shortly. Reference ${leadId}.</div>
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background:${BG};">
<tr><td align="center" style="padding:30px 16px;">

<table role="presentation" class="ttd-wrap" width="650" cellpadding="0" cellspacing="0" border="0" style="width:650px;max-width:650px;background:#ffffff;border-radius:15px;overflow:hidden;box-shadow:0 20px 50px -20px rgba(8,17,32,.25);">

  <!-- HEADER -->
  <tr><td class="ttd-pad" style="background:linear-gradient(135deg,${NAVY},${NAVY_2});padding:40px 40px 30px 40px;text-align:center;">
    <div style="font-family:'Playfair Display',Georgia,serif;font-size:24px;font-weight:700;color:#ffffff;letter-spacing:.02em;">
      Tour<span style="color:${ORANGE};">Trips</span>Dubai
    </div>
    <div style="font-family:'Work Sans',Arial,sans-serif;font-size:10px;letter-spacing:.35em;text-transform:uppercase;color:${ORANGE};font-weight:600;margin-top:6px;">Premium Dubai Experiences</div>
    <div style="height:1px;background:rgba(255,255,255,.15);margin:26px 0;"></div>
    <div style="font-family:'Playfair Display',Georgia,serif;font-size:27px;line-height:1.3;color:#ffffff;font-weight:700;">Thank You for Contacting TourTripsDubai</div>
    <div style="font-family:'Work Sans',Arial,sans-serif;font-size:14px;color:rgba(255,255,255,.7);margin-top:10px;">Your enquiry has been successfully received.</div>
  </td></tr>

  <!-- BODY -->
  <tr><td class="ttd-pad" style="padding:34px 40px 10px 40px;background:${BG};">

    <!-- Greeting -->
    <div style="font-family:'Playfair Display',Georgia,serif;font-size:19px;font-weight:700;color:${NAVY};margin-bottom:12px;">Dear ${fullName},</div>
    <p style="font-family:'Work Sans',Arial,sans-serif;font-size:14.5px;color:${INK};line-height:1.7;margin:0 0 10px 0;">Thank you for choosing <strong>TourTripsDubai</strong>. We have successfully received your enquiry.</p>
    <p style="font-family:'Work Sans',Arial,sans-serif;font-size:14.5px;color:${INK};line-height:1.7;margin:0 0 26px 0;">Our Dubai travel expert will contact you shortly through <strong style="color:${ORANGE};">WhatsApp</strong> or <strong style="color:${ORANGE};">Email</strong>.</p>

    <!-- Booking Summary Card -->
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#ffffff;border:1px solid ${BORDER};border-radius:15px;box-shadow:0 8px 24px -12px rgba(8,17,32,.12);margin-bottom:28px;">
      <tr><td style="padding:28px 28px 6px 28px;">
        <div style="font-family:'Playfair Display',Georgia,serif;font-size:18px;font-weight:700;color:${NAVY};margin-bottom:20px;">&#128203; Booking Summary</div>
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" class="ttd-stack">
          <tr>
            ${row('Booking Reference', '<span style="color:' + ORANGE + '">' + leadId + '</span>')}
            ${row('Selected Package', pkg || '—')}
          </tr>
          <tr>
            ${row('Travel Date', travelDate || '—')}
            ${row('Adults / Children', (adults || '0') + ' Adult' + (adults === '1' ? '' : 's') + ' · ' + (children || '0') + ' Child' + (children === '1' ? '' : 'ren'))}
          </tr>
          <tr>
            ${row('WhatsApp Number', whatsappSafe || '—')}
            ${row('Email Address', email || '—')}
          </tr>
          ${fullRow('Special Request', message)}
        </table>
      </td></tr>
    </table>

    <!-- Why Book With Us -->
    <div style="font-family:'Playfair Display',Georgia,serif;font-size:18px;font-weight:700;color:${NAVY};margin-bottom:16px;text-align:center;">Why Book With Us</div>
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" class="ttd-stack" style="margin-bottom:12px;">
      <tr>
        ${whyCard('&#128176;', 'Best Price Guarantee', "Unbeatable rates on Dubai's best experiences, guaranteed.")}
        ${whyCard('&#128737;', 'Trusted Dubai Experts', 'Handpicked local experts who know Dubai inside out.')}
      </tr>
      <tr>
        ${whyCard('&#128172;', '24/7 WhatsApp Support', 'Round-the-clock assistance, whenever you need us.')}
        ${whyCard('&#128274;', 'Secure Booking', 'Safe, hassle-free booking from enquiry to arrival.')}
      </tr>
    </table>

    <!-- Need Immediate Help -->
    <div style="font-family:'Playfair Display',Georgia,serif;font-size:18px;font-weight:700;color:${NAVY};margin:28px 0 16px 0;text-align:center;">Need Immediate Help?</div>
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom:10px;">
      <tr><td align="center" style="border-radius:10px;background:#25D366;">
        <a href="${waHref}" target="_blank" style="display:block;padding:18px;font-family:'Work Sans',Arial,sans-serif;font-size:16px;font-weight:700;color:#ffffff;text-decoration:none;">&#128172; Chat with Travel Expert</a>
      </td></tr>
    </table>
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" class="ttd-btn" style="margin-bottom:10px;">
      <tr>
        ${ctaButton(telHref, ORANGE, '#ffffff', '&#128222; Call Now')}
        ${ctaButton(SITE_URL, '#ffffff', NAVY, '&#127760; Visit Website')}
      </tr>
    </table>

  </td></tr>

  <!-- FOOTER -->
  <tr><td class="ttd-pad" style="background:${NAVY};padding:36px 40px;text-align:center;">
    <div style="font-family:'Playfair Display',Georgia,serif;font-size:18px;font-weight:700;color:#ffffff;">TourTripsDubai</div>
    <div style="font-family:'Work Sans',Arial,sans-serif;font-size:12px;color:rgba(255,255,255,.55);margin-top:4px;">Premium Dubai Travel Company</div>

    <div style="height:1px;background:rgba(255,255,255,.12);margin:22px 0;"></div>

    <div style="font-family:'Work Sans',Arial,sans-serif;font-size:12px;color:rgba(255,255,255,.75);line-height:2;">
      <a href="${SITE_URL}" style="color:${ORANGE};text-decoration:none;">www.tourtripsdubai.com</a> &nbsp;·&nbsp;
      <a href="mailto:info@tourtripsdubai.com" style="color:rgba(255,255,255,.75);text-decoration:none;">info@tourtripsdubai.com</a><br>
      <a href="tel:+971589147828" style="color:rgba(255,255,255,.75);text-decoration:none;">WhatsApp: ${ADMIN_PHONE_DISPLAY}</a> &nbsp;·&nbsp;
      <span style="color:rgba(255,255,255,.55);">Dubai, UAE</span>
    </div>

    <table role="presentation" cellpadding="0" cellspacing="0" border="0" align="center" style="margin:20px auto 0 auto;">
      <tr>
        ${socialIconRow()}
      </tr>
    </table>

    <div style="height:1px;background:rgba(255,255,255,.12);margin:22px 0;"></div>

    <div style="font-family:'Work Sans',Arial,sans-serif;font-size:11px;color:rgba(255,255,255,.4);line-height:1.6;">
      © ${new Date().getFullYear()} TourTripsDubai. All rights reserved.<br>
      This is an automated confirmation email. Prices and availability are subject to change until final booking confirmation.
    </div>
  </td></tr>

</table>
</td></tr>
</table>
</body>
</html>`;
}
