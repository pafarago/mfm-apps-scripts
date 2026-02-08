/**
 * MFM Self-Registration Webhook (POST only)
 * - Receives cross-origin POSTs from your standalone website form using a hidden iframe.
 * - Replies with small HTML that postMessage()s the JSON result to the parent window.
 *
 * Deploy as Web App:
 *   Execute as: you (clerk/admin)
 *   Who has access: Anyone
 */

const BUILD_ID = 'MFM-SR-WEBHOOK / 2025-12-25 (updates+cards)';
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
      // CHANGED: return { message, isUpdate } so the client can display the right splash
      const out = selfRegVerify(payload); // { message, isUpdate }
      const msg = (out && out.message) ? out.message : String(out || '');
      const ok = /confirmed/i.test(String(msg || ''));
      return srPostMessageHtml_({ ok: ok, action: 'verify', message: msg, isUpdate: !!(out && out.isUpdate) });
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
    message: obj.message || '',
    // NEW: optional fields, only used when present
    isUpdate: (obj.isUpdate === true)
  });

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

/**
 * Lookup roster info by email. Returns:
 *   { exists:boolean, rowIndex0:number, info:{first,last,phone,addr} }
 */
function srLookupRosterByEmail_(email) {
  const ss = SpreadsheetApp.openById(SR_ROSTER_FILE_ID);
  const sh = ss.getSheetByName(SR_ROSTER_SHEET_NAME);
  if (!sh) throw new Error('Roster sheet not found.');

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return { exists: false, rowIndex0: -1, info: { first:'', last:'', phone:'', addr:'' } };

  const headers = data[0].map(v => String(v || ''));
  const idx = (name) => headers.indexOf(name);

  const iEmail = idx('Email');
  if (iEmail < 0) throw new Error('Roster headers missing Email.');

  const iFirst = idx('First name');
  const iLast  = idx('Last name');
  const iPhone = idx('Phone');
  const iAddr  = idx('Address');

  const target = String(email || '').trim().toLowerCase();

  for (let r = 1; r < data.length; r++) {
    const rowEmail = String(data[r][iEmail] || '').trim().toLowerCase();
    if (rowEmail && rowEmail === target) {
      return {
        exists: true,
        rowIndex0: r,
        info: {
          first: iFirst >= 0 ? String(data[r][iFirst] || '').trim() : '',
          last:  iLast  >= 0 ? String(data[r][iLast]  || '').trim() : '',
          phone: iPhone >= 0 ? String(data[r][iPhone] || '').trim() : '',
          addr:  iAddr  >= 0 ? String(data[r][iAddr]  || '').trim() : ''
        }
      };
    }
  }

  return { exists: false, rowIndex0: -1, info: { first:'', last:'', phone:'', addr:'' } };
}

function selfRegSubmit(form) {
  if (!form) throw new Error('Missing form data');

  const email = String(form.email || '').trim().toLowerCase();

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

  // if email exists, treat as an update (do not reject)
  const rosterHit = srLookupRosterByEmail_(email);
  const isUpdate  = !!rosterHit.exists;
  const orig      = rosterHit.info || { first:'', last:'', phone:'', addr:'' };

  const pending = srGetOrCreatePendingSheet_();
  srEnsurePendingHeaders_(pending);

  const headers = pending.getRange(1, 1, 1, pending.getLastColumn()).getValues()[0].map(String);
  const col = (name) => headers.indexOf(name) + 1;

  const token = srGenerateToken_();
  const now   = new Date();

  const row = new Array(headers.length).fill('');

  row[col('Token') - 1]                = token;
  row[col('Requested at') - 1]         = now;
  row[col('Email') - 1]                = email;

  row[col('First name') - 1]           = first;
  row[col('Last name') - 1]            = last;
  row[col('Phone') - 1]                = phone;
  row[col('Address') - 1]              = addr;

  row[col('Is update') - 1]            = isUpdate ? 'TRUE' : 'FALSE';
  row[col('Original first name') - 1]  = isUpdate ? (orig.first || '') : '';
  row[col('Original last name') - 1]   = isUpdate ? (orig.last  || '') : '';
  row[col('Original phone') - 1]       = isUpdate ? (orig.phone || '') : '';
  row[col('Original address') - 1]     = isUpdate ? (orig.addr  || '') : '';

  row[col('Confirmed at') - 1]         = '';

  pending.appendRow(row);

  srSendConfirmEmail_(email, token, {
    isUpdate,
    requested: { first, last, phone, addr },
    original:  orig
  });

  srNotifyAdminPending_(email, first, last, phone, addr, token, now, isUpdate, orig);

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
      'Is update',
      'Original first name',
      'Original last name',
      'Original phone',
      'Original address',
      'Confirmed at'
    ]);
  }
  return sh;
}

/**
 * Ensures the Pending sheet has the columns needed for update flow and cards.
 * Adds any missing columns at the end (keeps existing data intact).
 */
function srEnsurePendingHeaders_(sh) {
  const want = [
    'Token',
    'Requested at',
    'Email',
    'First name',
    'Last name',
    'Phone',
    'Address',
    'Is update',
    'Original first name',
    'Original last name',
    'Original phone',
    'Original address',
    'Confirmed at'
  ];

  const lastCol = Math.max(1, sh.getLastColumn());
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim());
  const existing = new Set(headers.filter(Boolean));

  const toAdd = want.filter(h => !existing.has(h));
  if (toAdd.length) {
    sh.getRange(1, lastCol + 1, 1, toAdd.length).setValues([toAdd]);
  }
}

function srGenerateToken_() {
  const n = Math.floor(Math.random() * 1000000);
  return String(n).padStart(6, '0');
}

/* =========================
   Email: formatted cards
   ========================= */

function srEscHtml_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function srFieldRowsHtml_(obj) {
  const rows = [
    ['Email', obj.email || ''],
    ['First name', obj.first || ''],
    ['Last name', obj.last || ''],
    ['Phone', obj.phone || ''],
    ['Address', obj.addr || '']
  ];

  return rows.map(([k, v]) => {
    const vv = (k === 'Address')
      ? '<div style="white-space:pre-wrap;line-height:1.35;">' + srEscHtml_(v || '') + '</div>'
      : srEscHtml_(v || '');
    return (
      '<tr>' +
      '<td style="padding:6px 8px;color:#444;font-weight:600;vertical-align:top;width:140px;">' + srEscHtml_(k) + '</td>' +
      '<td style="padding:6px 8px;color:#111;vertical-align:top;">' + vv + '</td>' +
      '</tr>'
    );
  }).join('');
}

function srCardHtml_(title, obj) {
  return (
    '<div style="border:1px solid #ddd;border-radius:14px;background:#fff;padding:14px 14px 10px;box-shadow:0 1px 3px rgba(0,0,0,.08);">' +
      '<div style="font-size:16px;font-weight:750;color:#0a3a66;margin:0 0 10px;">' + srEscHtml_(title) + '</div>' +
      '<table role="presentation" cellpadding="0" cellspacing="0" style="border-collapse:collapse;width:100%;">' +
        srFieldRowsHtml_(obj) +
      '</table>' +
    '</div>'
  );
}

function srSendConfirmEmail_(email, token, ctx) {
  const isUpdate = !!(ctx && ctx.isUpdate);
  const requested = (ctx && ctx.requested) ? ctx.requested : { first:'', last:'', phone:'', addr:'' };
  const original  = (ctx && ctx.original)  ? ctx.original  : { first:'', last:'', phone:'', addr:'' };

  const subject = isUpdate
    ? 'Confirm your update for Midlothian Friends news'
    : 'Your verification code for Midlothian Friends news';

  const intro = isUpdate
    ? 'We received a request to update an existing Midlothian Friends Meeting news subscription.'
    : 'Thank you for your interest in receiving news from Midlothian Friends Meeting.';

  const actionLine = isUpdate
    ? 'If you want to proceed with this update, please verify using the code below.'
    : 'Please verify using the code below to confirm your subscription.';

  const warning = isUpdate
    ? 'If you did not request this update, please ignore this email.'
    : 'If you did not request this, you can ignore this message.';

  const html =
    '<div style="font-family:system-ui,-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,Arial,sans-serif;line-height:1.45;color:#111;background:#f5f5f7;padding:18px;">' +
      '<div style="max-width:720px;margin:0 auto;">' +
        '<div style="font-size:18px;font-weight:750;margin:0 0 10px;">Dear Friend,</div>' +
        '<div style="margin:0 0 14px;">' + srEscHtml_(intro) + '</div>' +

        (isUpdate
          ? (
              '<div style="display:block;margin:0 0 14px;">' +
                '<div style="margin:0 0 10px;">Please review the information below.</div>' +
                '<div style="display:block;">' +
                  srCardHtml_('Current information on file', { email, first: original.first, last: original.last, phone: original.phone, addr: original.addr }) +
                  '<div style="height:12px;"></div>' +
                  srCardHtml_('Requested update', { email, first: requested.first, last: requested.last, phone: requested.phone, addr: requested.addr }) +
                '</div>' +
              '</div>'
            )
          : (
              '<div style="margin:0 0 14px;">' +
                srCardHtml_('Information to be entered', { email, first: requested.first, last: requested.last, phone: requested.phone, addr: requested.addr }) +
              '</div>'
            )
        ) +

        '<div style="margin:0 0 10px;">' + srEscHtml_(actionLine) + '</div>' +

        '<div style="margin:12px 0 16px;padding:14px;border-radius:14px;background:#ffffff;border:1px solid #ddd;">' +
          '<div style="font-size:14px;color:#444;margin:0 0 6px;">Your verification code is:</div>' +
          '<div style="font-size:28px;font-weight:850;letter-spacing:2px;color:#111;">' + srEscHtml_(token) + '</div>' +
          '<div style="font-size:13px;color:#555;margin-top:10px;">Return to the signup page, enter the code, and click <b>Verify</b>.</div>' +
        '</div>' +

        '<div style="margin:0 0 14px;color:#333;">' + srEscHtml_(warning) + '</div>' +

        '<div style="margin-top:18px;color:#111;">In Friendship,<br>' + srEscHtml_(SR_FROM_NAME) + '</div>' +
      '</div>' +
    '</div>';

  const textBody =
    'Dear Friend,\n\n' +
    intro + '\n\n' +
    (isUpdate
      ? 'Current on file:\n' +
        'Email: ' + email + '\n' +
        'First name: ' + (original.first || '') + '\n' +
        'Last name: ' + (original.last || '') + '\n' +
        'Phone: ' + (original.phone || '') + '\n' +
        'Address: ' + (original.addr || '') + '\n\n' +
        'Requested update:\n' +
        'Email: ' + email + '\n' +
        'First name: ' + (requested.first || '') + '\n' +
        'Last name: ' + (requested.last || '') + '\n' +
        'Phone: ' + (requested.phone || '') + '\n' +
        'Address: ' + (requested.addr || '') + '\n\n'
      : 'Information to be entered:\n' +
        'Email: ' + email + '\n' +
        'First name: ' + (requested.first || '') + '\n' +
        'Last name: ' + (requested.last || '') + '\n' +
        'Phone: ' + (requested.phone || '') + '\n' +
        'Address: ' + (requested.addr || '') + '\n\n'
    ) +
    actionLine + '\n\n' +
    'Your verification code is:\n\n' +
    token + '\n\n' +
    'Please return to the signup page, enter the code, and click Verify.\n\n' +
    warning + '\n\n' +
    'In Friendship,\n' +
    SR_FROM_NAME + '\n';

  GmailApp.sendEmail(
    email,
    subject,
    textBody,
    { name: SR_FROM_NAME, htmlBody: html }
  );
}

/**
 * CHANGED: returns { message, isUpdate }
 */
function selfRegVerify(payload) {
  if (!payload) throw new Error('Missing verification data');

  const email = String(payload.email || '').trim().toLowerCase();
  const code  = String(payload.code  || '').replace(/\D+/g, '').trim();

  if (!email) throw new Error('Email is required.');
  if (!/^\d{6}$/.test(code)) throw new Error('Verification code must be 6 digits.');

  const pending = srGetOrCreatePendingSheet_();
  srEnsurePendingHeaders_(pending);

  const data    = pending.getDataRange().getValues();
  if (data.length < 2) return { message: 'That code was not found (or it has already been used).', isUpdate: false };

  const headers = data[0].map(String);
  const iToken  = headers.indexOf('Token');
  const iEmail  = headers.indexOf('Email');
  const iFirst  = headers.indexOf('First name');
  const iLast   = headers.indexOf('Last name');
  const iPhone  = headers.indexOf('Phone');
  const iAddr   = headers.indexOf('Address');
  const iConf   = headers.indexOf('Confirmed at');
  const iIsUpd  = headers.indexOf('Is update');

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
    return { message: 'That code was not found (or it has already been used). Please double-check and try again.', isUpdate: false };
  }

  const first = String(data[rowIdx][iFirst] || '').trim();
  const last  = String(data[rowIdx][iLast]  || '').trim();
  const phone = String(data[rowIdx][iPhone] || '').trim();
  const addr  = String(data[rowIdx][iAddr]  || '').trim();

  const isUpdate = (iIsUpd >= 0) ? /true/i.test(String(data[rowIdx][iIsUpd] || '')) : false;

  srApplyToRoster_(email, { first:first, last:last, phone:phone, addr:addr });
  pending.getRange(rowIdx + 1, iConf + 1).setValue(new Date());

  srNotifyAdminConfirmed_(email, first, last, isUpdate);

  return { message: 'Thank you — your subscription is confirmed.', isUpdate: isUpdate };
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

function srNotifyAdminPending_(email, first, last, phone, addr, token, when, isUpdate, orig) {
  if (!SR_ADMIN_EMAIL) return;

  const payload = {
    type: isUpdate ? 'pending_update' : 'pending_subscription',
    email,
    requested: { first_name: first, last_name: last, phone, address: addr },
    original: isUpdate ? { first_name: orig.first, last_name: orig.last, phone: orig.phone, address: orig.addr } : null,
    requested_at: when,
    token
  };

  GmailApp.sendEmail(
    SR_ADMIN_EMAIL,
    isUpdate ? '[MFM-News] Pending subscription update' : '[MFM-News] Pending subscription',
    'A newsletter request has been received.\n\n' + JSON.stringify(payload, null, 2) + '\n',
    { name: SR_FROM_NAME }
  );
}

function srNotifyAdminConfirmed_(email, first, last, isUpdate) {
  if (!SR_ADMIN_EMAIL) return;

  const payload = {
    type: isUpdate ? 'confirmed_update' : 'confirmed_subscription',
    email,
    first_name: first,
    last_name: last
  };

  GmailApp.sendEmail(
    SR_ADMIN_EMAIL,
    isUpdate ? '[MFM-News] Subscription update confirmed' : '[MFM-News] Subscription confirmed',
    'A request has been confirmed and applied to the roster.\n\n' + JSON.stringify(payload, null, 2) + '\n',
    { name: SR_FROM_NAME }
  );
}

