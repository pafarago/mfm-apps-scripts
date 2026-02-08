/**
 * MFM Self-Registration Webhook (POST only)
 * - Receives cross-origin POSTs from your standalone website form using a hidden iframe.
 * - Replies with small HTML that postMessage()s the JSON result to the parent window.
 *
 * Deploy as Web App:
 *   Execute as: you (clerk/admin)
 *   Who has access: Anyone
 */

const BUILD_ID = 'MFM-SR-WEBHOOK / 2025-12-25';
const SR_ROSTER_FILE_ID     = '1e17lPBAz12xFghE0nqJ76cKr8MHv2WIbFlY_du18axM';
const SR_ROSTER_SHEET_NAME  = 'Midlothian Friends Meeting Roster';
const SR_PENDING_SHEET_NAME = 'PendingSelfRegistrations';
const SR_FROM_NAME          = 'Midlothian Friends Meeting';
const SR_ADMIN_EMAIL        = 'clerk@midlothianfriends.org'; // optional

function getBuildId(){ return BUILD_ID; }

/**
 * POST endpoint
 * Expects application/x-www-form-urlencoded fields:
 *   action=submit|verify
 *   payload=<json string>
 */
function doPost(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) ? String(e.parameter.action) : '';
    const payloadStr = (e && e.parameter && e.parameter.payload) ? String(e.parameter.payload) : '{}';
    const payload = JSON.parse(payloadStr || '{}');

    if (action === 'submit') {
      const msg = selfRegSubmit(payload);
      return srPostMessageHtml_({ ok: true, action: 'submit', message: msg });
    }

    if (action === 'verify') {
      const msg = selfRegVerify(payload);
      // selfRegVerify returns a string; treat “confirmed” as ok unless you prefer strictness
      // (We’ll use ok=true when it contains “confirmed”, else ok=false.)
      const ok = /confirmed/i.test(String(msg || ''));
      return srPostMessageHtml_({ ok: ok, action: 'verify', message: msg });
    }

    return srPostMessageHtml_({ ok: false, action: action || 'unknown', message: 'Unknown action.' });
  } catch (err) {
    return srPostMessageHtml_({ ok: false, action: 'error', message: String(err && err.message ? err.message : err) });
  }
}

/**
 * Returns a tiny HTML page that can run inside an iframe and send data to the parent.
 */
function srPostMessageHtml_(obj) {
  const safeJson = JSON.stringify({
    type: 'mfm_sr',
    ok: !!obj.ok,
    action: obj.action || '',
    message: obj.message || ''
  });

  // Note: HtmlOutput can be iframed; we explicitly allow all.
  return HtmlService
    .createHtmlOutput(
      '<!doctype html><html><head><meta charset="utf-8"></head><body>' +
      '<script>' +
      '  try {' +
      '    var data = ' + safeJson + ';' +
      '    window.top && window.top.postMessage(data, "*");' +
      '  } catch(e) {}' +
      '</script>' +
      '</body></html>'
    )
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* =========================
   Your existing logic (kept)
   ========================= */

function normalizeUSPhoneNANP_(raw) {
  var s = String(raw || '').trim();
  if (!s) return '';

  var d = s.replace(/\D+/g, '');
  if (d.length === 11 && d.slice(0,1) === '1') d = d.slice(1);
  if (d.length !== 10) throw new Error('Not a valid phone number');

  var a = d.slice(0, 3);
  var b = d.slice(3, 6);
  var c = d.slice(6);

  if (!/^[2-9][0-9]{2}$/.test(a)) throw new Error('Not a valid phone number');
  if (!/^[2-9][0-9]{2}$/.test(b)) throw new Error('Not a valid phone number');

  return '(' + a + ') ' + b + '-' + c;
}

function srEmailAlreadyUsed_(email) {
  const ss = SpreadsheetApp.openById(SR_ROSTER_FILE_ID);
  const sh = ss.getSheetByName(SR_ROSTER_SHEET_NAME);
  if (!sh) throw new Error('Roster sheet not found.');

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return false;

  const headers = data[0];
  const iEmail = headers.indexOf('Email');
  if (iEmail < 0) throw new Error('Roster headers missing Email.');

  const target = String(email).trim().toLowerCase();
  for (let r = 1; r < data.length; r++) {
    const rowEmail = String(data[r][iEmail] || '').trim().toLowerCase();
    if (rowEmail && rowEmail === target) return true;
  }
  return false;
}

function selfRegSubmit(form) {
  if (!form) throw new Error('Missing form data');

  const email = String(form.email || '').trim().toLowerCase();
  if (srEmailAlreadyUsed_(email)) {
    return email + ' is already subscribed.';
  }

  const first = String(form.firstName || '').trim();
  const last  = String(form.lastName || '').trim();
  let   phone = String(form.phone || '').trim();
  const addr  = String(form.address || '').trim();

  if (!email) throw new Error('Email is required.');
  if (!first || !last) throw new Error('First and last name are required.');

  if (phone) {
    try { phone = normalizeUSPhoneNANP_(phone); }
    catch (err) { phone = ''; }
  }

  const pending = srGetOrCreatePendingSheet_();
  const headers = pending.getRange(1, 1, 1, pending.getLastColumn()).getValues()[0];
  const col = (name) => headers.indexOf(name) + 1;

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
  const n = Math.floor(Math.random() * 1000000);
  return String(n).padStart(6, '0');
}

function srSendConfirmEmail_(email, token) {
  const body =
    'Dear Friend,\n\n' +
    'Thank you for your interest in receiving news from Midlothian Friends Meeting.\n\n' +
    'Your verification code is:\n\n' +
    token + '\n\n' +
    'Please return to the signup page, enter the code, and click Verify.\n\n' +
    'If you did not request this, you can ignore this message.\n\n' +
    'In Friendship,\n' +
    SR_FROM_NAME + '\n';

  GmailApp.sendEmail(
    email,
    'Your verification code for Midlothian Friends news',
    body,
    { name: SR_FROM_NAME }
  );
}

function selfRegVerify(payload) {
  if (!payload) throw new Error('Missing verification data');

  const email = String(payload.email || '').trim().toLowerCase();
  const code  = String(payload.code  || '').replace(/\D+/g, '').trim();

  if (!email) throw new Error('Email is required.');
  if (!/^\d{6}$/.test(code)) throw new Error('Verification code must be 6 digits.');

  const pending = srGetOrCreatePendingSheet_();
  const data    = pending.getDataRange().getValues();
  if (data.length < 2) return 'That code was not found (or it has already been used).';

  const headers = data[0];
  const iToken  = headers.indexOf('Token');
  const iEmail  = headers.indexOf('Email');
  const iFirst  = headers.indexOf('First name');
  const iLast   = headers.indexOf('Last name');
  const iPhone  = headers.indexOf('Phone');
  const iAddr   = headers.indexOf('Address');
  const iConf   = headers.indexOf('Confirmed at');

  if ([iToken,iEmail,iFirst,iLast,iPhone,iAddr,iConf].some(i => i < 0)) {
    throw new Error('Pending sheet headers are missing required columns.');
  }

  let rowIdx = -1;
  for (let r = 1; r < data.length; r++) {
    const rowEmail = String(data[r][iEmail] || '').trim().toLowerCase();
    const rowToken = String(data[r][iToken] || '').trim();
    const confirmed = String(data[r][iConf] || '').trim();
    if (!confirmed && rowEmail === email && rowToken === code) {
      rowIdx = r;
      break;
    }
  }

  if (rowIdx < 0) {
    return 'That code was not found (or it has already been used). Please double-check and try again.';
  }

  const first = String(data[rowIdx][iFirst] || '').trim();
  const last  = String(data[rowIdx][iLast]  || '').trim();
  const phone = String(data[rowIdx][iPhone] || '').trim();
  const addr  = String(data[rowIdx][iAddr]  || '').trim();

  srApplyToRoster_(email, { first:first, last:last, phone:phone, addr:addr });
  pending.getRange(rowIdx + 1, iConf + 1).setValue(new Date());

  srNotifyAdminConfirmed_(email, first, last);

  return 'Thank you — your subscription is confirmed.';
}

function srApplyToRoster_(email, info) {
  const ss = SpreadsheetApp.openById(SR_ROSTER_FILE_ID);
  const sh = ss.getSheetByName(SR_ROSTER_SHEET_NAME);
  if (!sh) throw new Error('Roster sheet not found.');

  const vals    = sh.getDataRange().getValues();
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

  let rowIdx = -1;
  for (let r = 1; r < vals.length; r++) {
    const em = String(vals[r][iEmail] || '').trim().toLowerCase();
    if (em === email) { rowIdx = r; break; }
  }

  if (rowIdx === -1) {
    const newRow = new Array(headers.length).fill('');
    newRow[iEmail] = email;
    if (iFirst != null) newRow[iFirst] = info.first || '';
    if (iLast  != null) newRow[iLast]  = info.last  || '';
    if (iPhone != null) newRow[iPhone] = info.phone || '';
    if (iAddr  != null) newRow[iAddr]  = info.addr  || '';
    newRow[iQTC] = 'TRUE';
    sh.appendRow(newRow);
  } else {
    if (iFirst != null && info.first) vals[rowIdx][iFirst] = info.first;
    if (iLast  != null && info.last)  vals[rowIdx][iLast]  = info.last;
    if (iPhone != null)               vals[rowIdx][iPhone] = info.phone || vals[rowIdx][iPhone];
    if (iAddr  != null)               vals[rowIdx][iAddr]  = info.addr  || vals[rowIdx][iAddr];
    vals[rowIdx][iQTC] = 'TRUE';
    sh.getRange(rowIdx + 1, 1, 1, headers.length).setValues([vals[rowIdx]]);
  }

  const sortSpecs = [];
  if (iLast  != null) sortSpecs.push({ column: iLast  + 1, ascending: true });
  if (iFirst != null) sortSpecs.push({ column: iFirst + 1, ascending: true });
  sortSpecs.push({ column: iEmail + 1, ascending: true });

  const lastRow = sh.getLastRow();
  if (sortSpecs.length && lastRow > 1) {
    sh.getRange(2, 1, lastRow - 1, headers.length).sort(sortSpecs);
  }
}

function srNotifyAdminPending_(email, first, last, phone, addr, token, when) {
  if (!SR_ADMIN_EMAIL) return;

  const payload = {
    type: 'pending_subscription',
    email, first_name: first, last_name: last,
    phone, address: addr, requested_at: when, token
  };

  GmailApp.sendEmail(
    SR_ADMIN_EMAIL,
    '[MFM-News] Pending subscription',
    'A new newsletter subscription request has been received.\n\n' + JSON.stringify(payload, null, 2) + '\n',
    { name: SR_FROM_NAME }
  );
}

function srNotifyAdminConfirmed_(email, first, last) {
  if (!SR_ADMIN_EMAIL) return;

  const payload = { type: 'confirmed_subscription', email, first_name: first, last_name: last };

  GmailApp.sendEmail(
    SR_ADMIN_EMAIL,
    '[MFM-News] Subscription confirmed',
    'A subscription has been confirmed and applied to the roster.\n\n' + JSON.stringify(payload, null, 2) + '\n',
    { name: SR_FROM_NAME }
  );
}

