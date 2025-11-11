/***** CONFIG LOADER *****/

function getConfig_() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Config');
  if (!sh) throw new Error('Missing "Config" sheet.');

  const rows = sh.getDataRange().getValues().filter(r => r[0] && r[0] !== 'Key');
  const cfg = {};
  rows.forEach(r => cfg[String(r[0]).trim()] = String(r[1] ?? '').trim());

  if (!cfg.DOC_TEMPLATE_ID) throw new Error('Config: DOC_TEMPLATE_ID is required');
  if (!cfg.ROSTER_FILE_ID) throw new Error('Config: ROSTER_FILE_ID is required');
  if (!cfg.ROSTER_SHEET_NAME) cfg.ROSTER_SHEET_NAME = 'Roster';

  // Which roster checkbox columns define inclusion for this mailing.
  // Example: "QTC" or "QTC,DIRECTORY"
  const listColsRaw = cfg.LIST_COLUMNS || cfg.LIST_COLUMN || 'QTC';
  cfg.LIST_COLUMNS = listColsRaw.split(',')
    .map(s => s.trim())
    .filter(Boolean);
  if (!cfg.LIST_COLUMNS.length) {
    throw new Error('Config: LIST_COLUMNS (or LIST_COLUMN) must name at least one roster checkbox column.');
  }

  cfg.SUBJECT         = cfg.SUBJECT || 'Announcement';
  cfg.SENDER_NAME     = cfg.SENDER_NAME || 'Announcements';
  cfg.TEST_RECIPIENT  = cfg.TEST_RECIPIENT || Session.getActiveUser().getEmail();
  cfg.DELAY_MS        = Number(cfg.DELAY_MS || 1500);
  cfg.BATCH_LIMIT     = Number(cfg.BATCH_LIMIT || 500);
  cfg.ATTACH_PDF      = /^true$/i.test(cfg.ATTACH_PDF || 'FALSE');
  cfg.REPLY_TO        = cfg.REPLY_TO || '';

  return cfg;
}

function menuSendAll() {
  sendAll_(false);
}

function menuDryRun() {
  sendAll_(true);
}
/***** DATA ACCESS: MAILING LIST *****/

// Always read from sheet named "MailingList"
function getRows_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('MailingList');
  if (!sheet) throw new Error('Missing "MailingList" sheet. Run "Refresh mailing list from roster" first.');
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) throw new Error('MailingList has no data rows.');
  const headers = values[0];
  const headerMap = {};
  headers.forEach((h, i) => {
    if (h) headerMap[String(h).trim()] = i + 1; // 1-based for Range
  });
  const rows = values.slice(1).map(r => {
    const o = {};
    headers.forEach((h, i) => o[h] = r[i]);
    return o;
  });
  return { sheet, rows, headerMap, headers };
}

function markRow_(sheet, headerMap, rowIndex, statusText, clearSend) {
  const rowNum = rowIndex + 2;
  const now = new Date();
  const formatted =
     Utilities.formatDate(now, Session.getScriptTimeZone(), 'MM-dd-yy HH:mm');
  if (headerMap['Status'])  sheet.getRange(rowNum, headerMap['Status']).setValue(statusText);
  if (headerMap['Sent at']) sheet.getRange(rowNum, headerMap['Sent at']).setValue(formatted);
  if (clearSend && headerMap['Send?']) {
    sheet.getRange(rowNum, headerMap['Send?']).setValue('FALSE');
  }
}

/***** HELPERS *****/

function asBool_(v) {
  if (v === true) return true;
  const s = String(v || '').trim().toLowerCase();
  return ['true', 'yes', 'y', '1', 'x'].includes(s);
}

function escapeRegex_(s) {
  return String(s || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/***** DOC MERGE + EXPORT *****/

// Export merged Doc as HTML using Drive v3 (no Advanced Drive toggle)
function exportDocHtml_(docId) {
  const url = 'https://www.googleapis.com/drive/v3/files/' + docId +
              '/export?mimeType=text%2Fhtml&alt=media';
  const res = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  if (code !== 200) {
    throw new Error('HTML export failed (' + code + '): ' + res.getContentText());
  }
  return res.getContentText();
}

function renderDocToBlobs_(templateId, data) {
  // 1) Copy template and merge {{tokens}}
  const temp = DriveApp.getFileById(templateId).makeCopy('_merge_tmp_' + Date.now());
  const doc  = DocumentApp.openById(temp.getId());
  const body = doc.getBody();

  Object.keys(data).forEach(k => {
    const val = data[k] == null ? '' : String(data[k]);
    // {{ Token }} matching
    const pattern = '\\{\\{\\s*' + escapeRegex_(k) + '\\s*\\}\\}';
    body.replaceText(pattern, val);
  });

  doc.saveAndClose();

  // 2) PDF for optional attachment
  const pdfBlob = DriveApp.getFileById(doc.getId()).getAs('application/pdf');

  // 3) HTML body from merged Doc
  let htmlBody = exportDocHtml_(doc.getId());

  // Extract <body> content only
  const m = htmlBody.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
  htmlBody = m ? m[1] : htmlBody;
  // Normalize paragraph spacing to prevent cropped descenders
  htmlBody = htmlBody
  .replace(/<p([^>]*)>/g, '<p$1 style="margin-bottom:0.3em; line-height:1.3;">')
  .replace(/<li([^>]*)>/g, '<li$1 style="margin-bottom:0.3em; line-height:1.3;">');

  // 4) Plain text fallback
  const plainBody = htmlBody
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  // 5) Footer (unsubscribe info)
  const footerHtml =
    '<hr><p style="font-size:12px;color:#666">' +
    'Manage subscription: email news+unsubscribe@midlothianfriends.org ' +
    '(or reply “unsubscribe”).</p>';

  const footerTxt =
    '\n\nManage subscription: email news+unsubscribe@midlothianfriends.org ' +
    '(or reply "unsubscribe").\n';

  htmlBody = htmlBody + footerHtml;

  try { temp.setTrashed(true); } catch (e) {}

  return {
    pdfBlob,
    htmlBody,
    plainBody: plainBody + footerTxt
  };
}

/***** SENDING *****/

function sendOne_({to, subject, htmlBody, plainBody, pdfBlob, cfg}) {
  const opts = {
    name: cfg.SENDER_NAME,
    htmlBody: htmlBody
  };
  if (cfg.REPLY_TO) opts.replyTo = cfg.REPLY_TO;
  if (cfg.ATTACH_PDF && pdfBlob) opts.attachments = [pdfBlob];
  GmailApp.sendEmail(to, subject, plainBody, opts);
}

/***** MENUS *****/

// 1) Preview: send first Send?=TRUE row to TEST_RECIPIENT, using its data
function menuPreview() {
  const cfg = getConfig_();
  const {rows} = getRows_();

  const row = rows.find(r => asBool_(r['Send?']));
  if (!row) throw new Error('No rows in MailingList with Send? = TRUE');

  const merged = renderDocToBlobs_(cfg.DOC_TEMPLATE_ID, row);

  sendOne_({
    to: cfg.TEST_RECIPIENT,
    subject: cfg.SUBJECT,
    htmlBody: merged.htmlBody,
    plainBody: merged.plainBody,
    pdfBlob: cfg.ATTACH_PDF ? merged.pdfBlob : null,
    cfg
  });

  SpreadsheetApp.getUi().alert('Preview sent to ' + cfg.TEST_RECIPIENT);
}

// 2) Build MailingList from Roster based on LIST_COLUMNS
function menuRefreshMailingList() {
  const cfg = getConfig_();

  const rosterSS = SpreadsheetApp.openById(cfg.ROSTER_FILE_ID);
  const rosterSh = rosterSS.getSheetByName(cfg.ROSTER_SHEET_NAME);
  if (!rosterSh) throw new Error('Roster sheet "' + cfg.ROSTER_SHEET_NAME + '" not found.');

  const vals = rosterSh.getDataRange().getValues();
  if (vals.length < 2) throw new Error('Roster has no data rows.');

  const headers = vals[0].map(h => String(h).trim());
  const hmap = {};
  headers.forEach((h, i) => { if (h) hmap[h] = i; });

  const idxEmail = hmap['Email'];
  const idxFirst = hmap['First name'] ?? hmap['First Name'];
  const idxLast  = hmap['Last name']  ?? hmap['Last Name'];
  const idxPhone = hmap['Phone'] ?? hmap['phone'];
  const idxAddr  = hmap['Address'] ?? hmap['address'];

  if (idxEmail == null) throw new Error('Roster must have an "Email" column.');

  // Resolve which list columns exist
  const listIdxs = cfg.LIST_COLUMNS
    .map(name => ({ name, idx: hmap[name] }))
    .filter(o => o.idx != null);

  if (!listIdxs.length) {
    throw new Error(
      'None of the LIST_COLUMNS [' + cfg.LIST_COLUMNS.join(', ') + '] were found in the Roster.'
    );
  }

    const out = [];
  for (let r = 1; r < vals.length; r++) {
    const row = vals[r];
    const email = String(row[idxEmail] || '').trim();
    if (!email) continue;

    // Compute flags for each configured list column
    const flags = listIdxs.map(li => {
      const v = row[li.idx];
      const checked = (v === true) ||
        String(v).toUpperCase() === 'TRUE' ||
        String(v).trim() === '1';
      return checked;
    });

    // Include row if in ANY list
    const isInAnyList = flags.some(Boolean);
    if (!isInAnyList) continue;

    const first = idxFirst != null ? String(row[idxFirst] || '').trim() : '';
    const last  = idxLast  != null ? String(row[idxLast]  || '').trim() : '';
    const phone = idxPhone != null ? String(row[idxPhone] || '').trim() : '';
    const addr  = idxAddr  != null ? String(row[idxAddr]  || '').trim() : '';

    // Convert flags to visible TRUE/blank for each list column
    const flagTexts = flags.map(f => f ? 'TRUE' : '');

    out.push([
      email,
      first,
      last,
      phone,
      addr,
      ...flagTexts,
      'TRUE',   // Send?
      '',       // Status
      ''        // Sent at
    ]);
  }

  const ss = SpreadsheetApp.getActive();
  const name = 'MailingList';
  let ml = ss.getSheetByName(name);
  if (!ml) ml = ss.insertSheet(name);

  ml.clearContents();

  // Headers: base + one column per LIST_COLUMN + send/status fields
  let headersOut = [
    'Email',
    'First name',
    'Last name',
    'Phone',
    'Address',
    ...cfg.LIST_COLUMNS,
    'Send?',
    'Status',
    'Sent at'
  ];
  ml.getRange(1, 1, 1, headersOut.length).setValues([headersOut]);

  if (out.length) {
    ml.getRange(2, 1, out.length, headersOut.length).setValues(out);
  }

  ml = ss.getSheetByName(name);
  if (!ml) {
    ml = ss.insertSheet(name);
  }

  ml.clearContents();

  ml.getRange(1, 1, 1, headersOut.length).setValues([headersOut]);

  if (out.length) {
    ml.getRange(2, 1, out.length, headersOut.length).setValues(out);
  }

  SpreadsheetApp.getUi().alert(
    'MailingList refreshed with ' + out.length +
    ' rows from roster ' + cfg.ROSTER_SHEET_NAME +
    ' using [' + cfg.LIST_COLUMNS.join(', ') + '].'
  );
}

// 3) Send all from MailingList (throttled)
function sendAll_(dryRun) {
  const cfg = getConfig_();
  const {sheet, rows, headerMap} = getRows_();

  const ui = SpreadsheetApp.getUi();
  const label = dryRun ? 'DRY RUN (no emails will be sent)' : 'Send all';
  const confirm = ui.alert(
    'Confirm ' + (dryRun ? 'Dry Run' : 'Send'),
    (dryRun
      ? 'Proceed with dry run? This will NOT send any emails, but will mark matching rows as "Dry Run".'
      : 'Are you sure you want to send all messages marked TRUE?'),
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) {
    ui.alert(label + ' canceled.');
    return;
  }
  let sent = 0;

  for (let i = 0; i < rows.length; i++) {
    if (sent >= cfg.BATCH_LIMIT) break;

    const r = rows[i];
    if (!asBool_(r['Send?'])) continue;

    const to = String(r['Email'] || '').trim();
    if (!to) continue;

    try {
      const merged = renderDocToBlobs_(cfg.DOC_TEMPLATE_ID, r);
       if (dryRun) {
        markRow_(sheet, headerMap, i, 'Dry Run', true);
        sent++;
      } else {
      sendOne_({
        to, 
        subject: cfg.SUBJECT,
        htmlBody: merged.htmlBody,
        plainBody: merged.plainBody,
        pdfBlob: cfg.ATTACH_PDF ? merged.pdfBlob : null,
        cfg
      });
      markRow_(sheet, headerMap, i, 'Sent', true); // clear Send?
      sent++;
      if(sent % 10 === 0) {
       SpreadsheetApp.flush();
      }
      Utilities.sleep(cfg.DELAY_MS);
      }
    } catch (e) {
      markRow_(sheet, headerMap, i, 'ERROR: ' + e.message, false);
      Utilities.sleep(cfg.DELAY_MS);
    }
  }
  SpreadsheetApp.getUi().alert('Done. Sent: ' + sent);
}

/***** VALIDATOR *****/

// Checks {{tokens}} in Doc vs MailingList headers
function menuValidateTemplate() {
  const ui = SpreadsheetApp.getUi();
  const cfg = getConfig_();

  const ml = SpreadsheetApp.getActive().getSheetByName('MailingList');
  if (!ml) {
    ui.alert('No "MailingList" sheet. Run "Refresh mailing list from roster" first.');
    return;
  }

  const vals = ml.getDataRange().getValues();
  if (vals.length < 1) {
    ui.alert('"MailingList" has no headers.');
    return;
  }

  const headers = vals[0].map(h => String(h).trim()).filter(Boolean);

  const doc = DocumentApp.openById(cfg.DOC_TEMPLATE_ID);
  const text = doc.getBody().getText();

  const tokenSet = new Set();
  const re = /\{\{\s*([^}]+?)\s*\}\}/g;
  let m;
  while ((m = re.exec(text)) !== null) {
    tokenSet.add(m[1]);
  }

  const missingInSheet = [];
  tokenSet.forEach(t => { if (!headers.includes(t)) missingInSheet.push(t); });

  const unusedHeaders = headers
    .filter(h =>
      !['Email','First name','Last name','Phone','Address','Send?','Status','Sent at']
        .includes(h)
    )
    .filter(h => !tokenSet.has(h));

  const msg =
    'Found ' + tokenSet.size + ' placeholder(s) in Doc.\n\n' +
    (missingInSheet.length
      ? 'Missing in MailingList headers:\n- ' + missingInSheet.join('\n- ') + '\n\n'
      : 'All placeholders exist in MailingList.\n\n') +
    (unusedHeaders.length
      ? 'MailingList headers not used in Doc:\n- ' + unusedHeaders.join('\n- ') + '\n\n'
      : 'All MailingList headers used or optional.');

  ui.alert('Validation result', msg, ui.ButtonSet.OK);
}