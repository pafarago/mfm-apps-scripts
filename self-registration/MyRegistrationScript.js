/**
 * MFM Self-Registration Web App (Public)
 *
 * This project:
 *  - Serves a public signup form.
 *  - Stores submissions in PendingSelfRegistrations.
 *  - Emails a confirmation link to the requester.
 *  - On confirmation, creates/updates the roster record and sets QTC = TRUE.
 *
 * Security:
 *  - No roster data is ever shown in the web UI.
 *  - Only the spreadsheet's owner/managers, set as Editors in Drive sharing,
 *    can open or edit the roster directly.
 *  - Web app should be deployed "Execute as: clerk/admin" and "Anyone".
 */

const BUILD_ID = 'MyRegistrationV5 / Version12 / ' + new Date().toISOString();

/** For debugging: confirm which code is running in the web app */
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
  const token = e && e.parameter && e.parameter.confirm;

  if (token) {
    return srHandleConfirm_(String(token).trim());
  }

  return srRenderForm_();
}

/** Render the public vertical signup form (MyRegistrationMarkup.html) */
function srRenderForm_() {
  const t = HtmlService.createTemplateFromFile('MyRegistrationMarkup');
  return t.evaluate()
    .setTitle('Midlothian Friends Newsletter Sign-up')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Called from MyRegistrationMarkup.html via google.script.run.selfRegSubmit_(form).
 * Stores pending request, sends confirmation email, notifies admin.
 */
function selfRegSubmit(form) {
  if (!form) throw new Error('Missing form data');

  const email = String(form.email || '').trim().toLowerCase();
  const first = String(form.firstName || '').trim();
  const last  = String(form.lastName || '').trim();
  const phone = String(form.phone || '').trim();
  const addr  = String(form.address || '').trim();

  if (!email) throw new Error('Email is required.');
  if (!first && !last) throw new Error('First or last name is required.');

  const pending = srGetOrCreatePendingSheet_();
  const headers = pending.getRange(1, 1, 1, pending.getLastColumn()).getValues()[0];

  const col = (name) => headers.indexOf(name) + 1;

  const token = srGenerateToken_();
  const now = new Date();

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

/**
 * Handle confirmation link: /exec?confirm=TOKEN
 * If valid and unused, applies to roster and marks QTC = TRUE.
 */
function srHandleConfirm_(token) {
  if (!token) {
    return HtmlService.createHtmlOutput(
      '<p>Invalid confirmation link.</p>'
    );
  }

  const pending = srGetOrCreatePendingSheet_();
  const data = pending.getDataRange().getValues();
  if (data.length < 2) {
    return HtmlService.createHtmlOutput(
      '<p>Confirmation link not found or already used.</p>'
    );
  }

  const headers = data[0];
  const iToken   = headers.indexOf('Token');
  const iEmail   = headers.indexOf('Email');
  const iFirst   = headers.indexOf('First name');
  const iLast    = headers.indexOf('Last name');
  const iPhone   = headers.indexOf('Phone');
  const iAddr    = headers.indexOf('Address');
  const iConf    = headers.indexOf('Confirmed at');

  let rowIdx = -1;
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][iToken]) === token) {
      rowIdx = r;
      break;
    }
  }

  if (rowIdx === -1) {
    return HtmlService.createHtmlOutput(
      '<p>Confirmation link not found or already used.</p>'
    );
  }

  if (data[rowIdx][iConf]) {
    return HtmlService.createHtmlOutput(
      '<p>This subscription has already been confirmed.</p>'
    );
  }

  const email = String(data[rowIdx][iEmail] || '').trim().toLowerCase();
  const first = String(data[rowIdx][iFirst] || '').trim();
  const last  = String(data[rowIdx][iLast] || '').trim();
  const phone = String(data[rowIdx][iPhone] || '').trim();
  const addr  = String(data[rowIdx][iAddr] || '').trim();

  // Apply confirmed subscription to roster (QTC = TRUE)
  srApplyToRoster_(email, { first, last, phone, addr });

  // Mark as confirmed
  pending.getRange(rowIdx + 1, iConf + 1).setValue(new Date());

  srNotifyAdminConfirmed_(email, first, last);

  return HtmlService.createHtmlOutput(
    '<p>Thank you. Your subscription has been confirmed.</p>'
  );
}
function testPing() {
  return 'pong';
}
/** Create or update roster row; set QTC = TRUE. No data ever exposed to public UI. */
function srApplyToRoster_(email, info) {
  const ss = SpreadsheetApp.openById(SR_ROSTER_FILE_ID);
  const sh = ss.getSheetByName(SR_ROSTER_SHEET_NAME);
  if (!sh) throw new Error('Roster sheet not found.');

  const vals = sh.getDataRange().getValues();
  const headers = vals[0].map(String);
  const h = {};
  headers.forEach((name, i) => { if (name) h[name] = i; });

  const iEmail = h['Email'];
  const iFirst = h['First name'];
  const iLast  = h['Last name'];
  const iPhone = h['Phone'];
  const iAddr  = h['Address'];
  const iQTC   = h['QTC'];

  if (iEmail == null || iQTC == null) {
    throw new Error('Roster must have "Email" and "QTC" columns.');
  }

  // find existing row
  let rowIdx = -1;
  for (let r = 1; r < vals.length; r++) {
    const em = String(vals[r][iEmail] || '').trim().toLowerCase();
    if (em === email) {
      rowIdx = r;
      break;
    }
  }

  if (rowIdx === -1) {
    // new record
    const newRow = new Array(headers.length).fill('');
    newRow[iEmail] = email;
    if (iFirst != null) newRow[iFirst] = info.first || '';
    if (iLast  != null) newRow[iLast]  = info.last  || '';
    if (iPhone != null) newRow[iPhone] = info.phone || '';
    if (iAddr  != null) newRow[iAddr]  = info.addr  || '';
    newRow[iQTC] = true;
    sh.appendRow(newRow);
  } else {
    // update existing
    if (iFirst != null && info.first) vals[rowIdx][iFirst] = info.first;
    if (iLast  != null && info.last)  vals[rowIdx][iLast]  = info.last;
    if (iPhone != null && info.phone) vals[rowIdx][iPhone] = info.phone;
    if (iAddr  != null && info.addr)  vals[rowIdx][iAddr]  = info.addr;
    vals[rowIdx][iQTC] = true;
    sh.getRange(rowIdx + 1, 1, 1, headers.length).setValues([vals[rowIdx]]);
  }

  // === Sort roster: Last name, First name, Email ===
  const sortSpecs = [];
  if (iLast  != null) sortSpecs.push({ column: iLast + 1,  ascending: true });
  if (iFirst != null) sortSpecs.push({ column: iFirst + 1, ascending: true });
  sortSpecs.push({ column: iEmail + 1, ascending: true });

  const lastRow = sh.getLastRow();
  if (sortSpecs.length && lastRow > 1) {
    sh.getRange(2, 1, lastRow - 1, headers.length).sort(sortSpecs);
  }
}

/** Ensure PendingSelfRegistrations sheet exists in the roster file */
function srGetOrCreatePendingSheet_() {
  const ss = SpreadsheetApp.openById(SR_ROSTER_FILE_ID);
  let sh = ss.getSheetByName(SR_PENDING_SHEET_NAME);
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

/** Generate a random token */
function srGenerateToken_() {
  const raw = Utilities.getUuid() + Math.random();
  return Utilities.base64Encode(raw)
    .replace(/[^A-Za-z0-9]/g, '')
    .substring(0, 32);
}

/** Send confirmation email to requester */
function srSendConfirmEmail_(email, token) {
  const url = ScriptApp.getService().getUrl();
  const link = url + '?confirm=' + encodeURIComponent(token);

  const body =
    'Dear Friend,\n\n' +
    'Thank you for your interest in receiving news from Midlothian Friends Meeting.\n\n' +
    'Please confirm your subscription by clicking the link below:\n' +
    link + '\n\n' +
    'If you did not request this, you can ignore this message.\n\n' +
    'In Friendship,\n' +
    SR_FROM_NAME + '\n';

  GmailApp.sendEmail(email, 'Confirm your subscription to Midlothian Friends news', body, {
    name: SR_FROM_NAME
  });
}

/** Notify clerk/admin of pending request (structured JSON in body) */
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

  GmailApp.sendEmail(SR_ADMIN_EMAIL, '[MFM-News] Pending subscription', body, {
    name: SR_FROM_NAME
  });
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

  GmailApp.sendEmail(SR_ADMIN_EMAIL, '[MFM-News] Subscription confirmed', body, {
    name: SR_FROM_NAME
  });
}