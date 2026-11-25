// Begin ChatGPT code block 2025-11-12 22:40 EST — Self-Registration (US phone only)
/**
 * MFM Self-Registration Web App (Public)
 *
 *  - Serves a public signup form (MyRegistrationMarkup.html).
 *  - Stores submissions in PendingSelfRegistrations.
 *  - Emails a confirmation link to the requester.
 *  - On confirmation, creates/updates the roster record and sets QTC = TRUE.
 *
 * Security:
 *  - No roster data is ever shown in the web UI.
 *  - Only the spreadsheet's owner/managers (Editors in Drive) can open or edit the roster.
 *  - Web app should be deployed "Execute as: clerk/admin" and "Anyone".
 */

const BUILD_ID = 'MyRegistrationV6 / 2025-11-12T22:40:00Z';
const SR_CONFIRM_SLUG_BASE = 'https://midlothianfriends.freeddns.org/confirm?token=';
function getBuildId() {
  return BUILD_ID;
}

/** CONFIG: set these to match your environment */
const SR_ROSTER_FILE_ID     = '1e17lPBAz12xFghE0nqJ76cKr8MHv2WIbFlY_du18axM';
const SR_ROSTER_SHEET_NAME  = 'Midlothian Friends Meeting Roster';
const SR_PENDING_SHEET_NAME = 'PendingSelfRegistrations';
const SR_FROM_NAME          = 'Midlothian Friends Meeting';
const SR_ADMIN_EMAIL        = 'clerk@midlothianfriends.org';  // optional but recommended

/**
 * Entry point:
 *  - GET with no params → show signup form.
 *  - GET with ?confirm=TOKEN → process confirmation.
 */
function doGet(e) {
  try {
    const token = e && e.parameter && e.parameter.confirm;
    if (token) {
      return srHandleConfirm_(String(token).trim());
    }
    return srRenderForm_();
  } catch (err) {
    const msg = (err && err.message) ? err.message : String(err);
    return HtmlService.createHtmlOutput(
      '<p><b>doGet failed</b></p><pre style="white-space:pre-wrap">' + msg + '</pre>'
    );
  }
}

function srIsAlreadySubscribed_(email) {
  const ss = SpreadsheetApp.openById(SR_ROSTER_FILE_ID);
  const sh = ss.getSheetByName('Roster');
  if (!sh) throw new Error('Roster sheet not found.');

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return false;

  const headers = data[0];
  const iEmail = headers.indexOf('Email');
  const iSub   = headers.indexOf('Subscribed?');

  if (iEmail < 0 || iSub < 0) throw new Error('Roster headers missing Email or Subscribed?.');

  const target = String(email).trim().toLowerCase();

  for (let r = 1; r < data.length; r++) {
    const rowEmail = String(data[r][iEmail] || '').trim().toLowerCase();
    const sub = data[r][iSub];
    if (rowEmail && rowEmail === target && (sub === true || String(sub).toLowerCase() === 'true')) {
      return true;
    }
  }
  return false;
}


/** Render the public signup form HTML */
function srRenderForm_() {
  const t = HtmlService.createTemplateFromFile('MyRegistrationMarkup');
  return t
    .evaluate()
    .setTitle('Midlothian Friends Newsletter Sign-up')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* =========================
   Phone: U.S. NANP only
   ========================= */

/**
 * Normalize/validate U.S. phone to "(AAA) BBB-CCCC".
 *
 * Rules:
 *  - Strip non-digits; if empty → '' (optional field).
 *  - Must be exactly 10 digits.
 *  - 1st digit (area code) 2–9; not N11.
 *  - 4th digit (central office) 2–9; not N11.
 *  - Returns formatted string "(AAA) BBB-CCCC"; throws Error on invalid.
 *
 * If you later want stricter checks, you can add a NANPA area-code
 * whitelist here before returning.
 */
function normalizeUSPhoneNANP_(raw) {
  var s = String(raw || '').trim();
  if (!s) return '';

  var d = s.replace(/\D+/g, '');
  if (d.length !== 10) throw new Error('Not a valid phone number');

  var a = d.slice(0, 3);
  var b = d.slice(3, 6);
  var c = d.slice(6);

  if (!/^[2-9][0-9]{2}$/.test(a)) throw new Error('Not a valid phone number');
  if (/^[2-9]11$/.test(a))        throw new Error('Not a valid phone number');

  if (!/^[2-9][0-9]{2}$/.test(b)) throw new Error('Not a valid phone number');
  if (/^[2-9]11$/.test(b))        throw new Error('Not a valid phone number');

  return '(' + a + ') ' + b + '-' + c;  // 14 chars
}

/**
 * Called from MyRegistrationMarkup.html via google.script.run.selfRegSubmit(form).
 * Stores pending request, sends confirmation email, notifies admin.
 */
function selfRegSubmit(form) {
  if (!form) throw new Error('Missing form data');

  const email = String(form.email || '').trim().toLowerCase();

  if (srIsAlreadySubscribed_(email)) {
    return 'That email address is already subscribed.';
  }

  const first = String(form.firstName || '').trim();
  const last  = String(form.lastName || '').trim();
  let   phone = String(form.phone || '').trim();
  const addr  = String(form.address || '').trim();

  if (!email) throw new Error('Email is required.');
  if (!first && !last) throw new Error('First or last name is required.');

  // Phone: optional; if present must be valid U.S. NANP
  if (phone) {
    try {
      phone = normalizeUSPhoneNANP_(phone);
    } catch (err) {
      // Do not block submission; just drop invalid phone
      phone = '';
    }
  }

  const pending = srGetOrCreatePendingSheet_();
  const headers = pending.getRange(1, 1, 1, pending.getLastColumn()).getValues()[0];

  const col = function(name) { return headers.indexOf(name) + 1; };

  const token = srGenerateToken_();
  const now   = new Date();

  const row = [];
  row[col('Token') - 1]        = token;
  row[col('Requested at') - 1] = now;
  row[col('Email') - 1]        = email;
  row[col('First name') - 1]   = first;
  row[col('Last name') - 1]    = last;
  row[col('Phone') - 1]        = phone;
  row[col('Address') - 1]      = addr;
  row[col('Confirmed at') - 1] = '';

  pending.appendRow(row);

  srSendConfirmEmail_(email, token);
  srNotifyAdminPending_(email, first, last, phone, addr, token, now);

  return 'Thank you. Please check your email to confirm your subscription.';
}

function srHtmlError_(label, err) {
  const msg = (err && err.message) ? err.message : String(err);
  return HtmlService.createHtmlOutput(
    '<p><b>' + label + '</b></p>' +
    '<pre style="white-space:pre-wrap">' + msg + '</pre>'
  );
}


/**
 * Handle confirmation link: /exec?confirm=TOKEN
 * If valid and unused, applies to roster and marks QTC = TRUE.
 */

function srHandleConfirm_(token) {
  try {
    if (!token) {
      return HtmlService.createHtmlOutput('<p>Invalid confirmation link.</p>');
    }

    const pending = srGetOrCreatePendingSheet_();
    const data    = pending.getDataRange().getValues();
    if (data.length < 2) {
      return HtmlService.createHtmlOutput('<p>Confirmation link not found or already used.</p>');
    }

    const headers = data[0];
    const iToken  = headers.indexOf('Token');
    const iEmail  = headers.indexOf('Email');
    const iFirst  = headers.indexOf('First name');
    const iLast   = headers.indexOf('Last name');
    const iPhone  = headers.indexOf('Phone');
    const iAddr   = headers.indexOf('Address');
    const iConf   = headers.indexOf('Confirmed at');

    let rowIdx = -1;
    for (let r = 1; r < data.length; r++) {
      if (String(data[r][iToken]) === token) { rowIdx = r; break; }
    }

    if (rowIdx === -1) {
      return HtmlService.createHtmlOutput('<p>Confirmation link not found or already used.</p>');
    }
    if (data[rowIdx][iConf]) {
      return HtmlService.createHtmlOutput('<p>This subscription has already been confirmed.</p>');
    }

    const email = String(data[rowIdx][iEmail] || '').trim().toLowerCase();
    const first = String(data[rowIdx][iFirst] || '').trim();
    const last  = String(data[rowIdx][iLast]  || '').trim();
    const phone = String(data[rowIdx][iPhone] || '').trim();
    const addr  = String(data[rowIdx][iAddr]  || '').trim();
    
    srApplyToRoster_(email, { first:first, last:last, phone:phone, addr:addr });
    pending.getRange(rowIdx + 1, iConf + 1).setValue(new Date());
    srSendRegistrantDisposition_(email, true);
    srNotifyAdminConfirmed_(email, first, last);
    return HtmlService.createHtmlOutput('<p>Thank you. Your subscription has been confirmed.</p>');
  } catch (err) {
    return srHtmlError_('Confirmation failed', err);
  }
}
/**
function srHandleConfirm_(token) {
  if (!token) {
    return HtmlService.createHtmlOutput('<p>Invalid confirmation link.</p>');
  }

  const pending = srGetOrCreatePendingSheet_();
  const data    = pending.getDataRange().getValues();
  if (data.length < 2) {
    return HtmlService.createHtmlOutput('<p>Confirmation link not found or already used.</p>');
  }

  const headers = data[0];
  const iToken  = headers.indexOf('Token');
  const iEmail  = headers.indexOf('Email');
  const iFirst  = headers.indexOf('First name');
  const iLast   = headers.indexOf('Last name');
  const iPhone  = headers.indexOf('Phone');
  const iAddr   = headers.indexOf('Address');
  const iConf   = headers.indexOf('Confirmed at');

  let rowIdx = -1;
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][iToken]) === token) {
      rowIdx = r;
      break;
    }
  }

  if (rowIdx === -1) {
    return HtmlService.createHtmlOutput('<p>Confirmation link not found or already used.</p>');
  }

  if (data[rowIdx][iConf]) {
    return HtmlService.createHtmlOutput('<p>This subscription has already been confirmed.</p>');
  }

  const email = String(data[rowIdx][iEmail] || '').trim().toLowerCase();
  const first = String(data[rowIdx][iFirst] || '').trim();
  const last  = String(data[rowIdx][iLast]  || '').trim();
  const phone = String(data[rowIdx][iPhone] || '').trim();
  const addr  = String(data[rowIdx][iAddr]  || '').trim();

  srApplyToRoster_(email, { first:first, last:last, phone:phone, addr:addr });

  pending.getRange(rowIdx + 1, iConf + 1).setValue(new Date());

  srNotifyAdminConfirmed_(email, first, last);

  return HtmlService.createHtmlOutput('<p>Thank you. Your subscription has been confirmed.</p>');
}
**/
function testPing() {
  return 'pong';
}

/** Create or update roster row; set QTC = TRUE. */
function srApplyToRoster_(email, info) {
  const ss = SpreadsheetApp.openById(SR_ROSTER_FILE_ID);
  const sh = ss.getSheetByName(SR_ROSTER_SHEET_NAME);
  if (!sh) throw new Error('Roster sheet not found.');

  const vals    = sh.getDataRange().getValues();
  const headers = vals[0].map(String);
  const h = {};
  headers.forEach(function(name, i){ if (name) h[name] = i; });

  const iEmail = h['Email'];
  const iFirst = h['First name'];
  const iLast  = h['Last name'];
  const iPhone = h['Phone'];
  const iAddr  = h['Address'];
  const iQTC   = h['QTC'];

  if (iEmail == null || iQTC == null) {
    throw new Error('Roster must have "Email" and "QTC" columns.');
  }

  let rowIdx = -1;
  for (let r = 1; r < vals.length; r++) {
    const em = String(vals[r][iEmail] || '').trim().toLowerCase();
    if (em === email) {
      rowIdx = r;
      break;
    }
  }

  if (rowIdx === -1) {
    const newRow = new Array(headers.length).fill('');
    newRow[iEmail] = email;
    if (iFirst != null) newRow[iFirst] = info.first || '';
    if (iLast  != null) newRow[iLast]  = info.last  || '';
    if (iPhone != null) newRow[iPhone] = info.phone || '';
    if (iAddr  != null) newRow[iAddr]  = info.addr  || '';
    newRow[iQTC] = true;
    sh.appendRow(newRow);
  } else {
    if (iFirst != null && info.first) vals[rowIdx][iFirst] = info.first;
    if (iLast  != null && info.last)  vals[rowIdx][iLast]  = info.last;
    if (iPhone != null)               vals[rowIdx][iPhone] = info.phone || vals[rowIdx][iPhone];
    if (iAddr  != null)               vals[rowIdx][iAddr]  = info.addr  || vals[rowIdx][iAddr];
    vals[rowIdx][iQTC] = true;
    sh.getRange(rowIdx + 1, 1, 1, headers.length).setValues([vals[rowIdx]]);
  }

  // Sort roster: Last name, First name, Email (if present)
  const sortSpecs = [];
  if (iLast  != null) sortSpecs.push({ column: iLast  + 1, ascending: true });
  if (iFirst != null) sortSpecs.push({ column: iFirst + 1, ascending: true });
  sortSpecs.push({ column: iEmail + 1, ascending: true });

  const lastRow = sh.getLastRow();
  if (sortSpecs.length && lastRow > 1) {
    sh.getRange(2, 1, lastRow - 1, headers.length).sort(sortSpecs);
  }
}

/** Ensure PendingSelfRegistrations sheet exists */
function srGetOrCreatePendingSheet_() {
  const ss = SpreadsheetApp.openById(SR_ROSTER_FILE_ID);
  let sh   = ss.getSheetByName(SR_PENDING_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SR_PENDING_SHEET_NAME);
    sh.appendRow([
      'Token',
      'Requested at',
      'Email',
      'First name',
      'Last name',
      'Phone',
      'Address',
      'Confirmed at'
    ]);
  }
  return sh;
}

function srGenerateToken_() {
  return Utilities.getUuid().replace(/-/g, ''); // 32 chars, URL-safe
}

/** Send confirmation email to requester */
function srSendConfirmEmail_(email, token) {
  const url  = ScriptApp.getService().getUrl();
  const link = url + '?confirm=' + encodeURIComponent(token);

  const textBody =
    'Dear Friend,\n\n' +
    'Thank you for your interest in receiving news from Midlothian Friends Meeting.\n\n' +
    'Please confirm your subscription by clicking the link below:\n' +
    link + '\n\n' +
    'After you confirm, you will receive a second email with the results for your records.\n\n' +
    'If you did not request this, you can ignore this message.\n\n' +
    'In Friendship,\n' +
    SR_FROM_NAME + '\n';

  const htmlBody =
    '<p>Dear Friend,</p>' +
    '<p>Thank you for your interest in receiving news from Midlothian Friends Meeting.</p>' +
    '<p>Please confirm your subscription by clicking the button below:</p>' +

    '<p style="text-align:center;margin:22px 0;">' +
    '<a href="' + link + '" ' +
    'style="display:inline-block;padding:14px 22px;font-size:16px;' +
    'background:#2f6fed;color:#ffffff;text-decoration:none;' +
    'border-radius:10px;">' +
    'Confirm Subscription</a></p>' +

    '<p style="font-size:14px;color:#444;">' +
    'If the button does not work, you may copy and paste this link into your browser:</p>' +

    '<p style="font-size:14px;word-break:break-all;">' +
    '<a href="' + link + '">' + link + '</a></p>' +

    '<p>After you confirm, you will receive a second email with the results for your records.</p>' +
    '<p>If you did not request this, you can ignore this message.</p>' +
    '<p>In Friendship,<br>' + SR_FROM_NAME + '</p>';

  GmailApp.sendEmail(
    email,
    'Confirm your subscription to Midlothian Friends news',
    textBody,
    { name: SR_FROM_NAME, htmlBody: htmlBody }
  );
}


/** Notify clerk/admin of pending request */
function srNotifyAdminPending_(email, first, last, phone, addr, token, when) {
  if (!SR_ADMIN_EMAIL) return;

  const payload = {
    type: 'pending_subscription',
    email: email,
    first_name: first,
    last_name: last,
    phone: phone,
    address: addr,
    requested_at: when,
    token: token
  };

  const body =
    'A new newsletter subscription request has been received.\n\n' +
    JSON.stringify(payload, null, 2) + '\n';

  GmailApp.sendEmail(
    SR_ADMIN_EMAIL,
    '[MFM-News] Pending subscription',
    body,
    { name: SR_FROM_NAME }
  );
}

/** Notify clerk/admin when confirmed & applied */
function srNotifyAdminConfirmed_(email, first, last) {
  if (!SR_ADMIN_EMAIL) return;

  const payload = {
    type: 'confirmed_subscription',
    email: email,
    first_name: first,
    last_name: last
  };

  const body =
    'A subscription has been confirmed and applied to the roster.\n\n' +
    JSON.stringify(payload, null, 2) + '\n';

  GmailApp.sendEmail(
    SR_ADMIN_EMAIL,
    '[MFM-News] Subscription confirmed',
    body,
    { name: SR_FROM_NAME }
  );
}
// End ChatGPT code block 2025-11-12 22:40 EST
