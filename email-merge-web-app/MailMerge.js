/**** CONFIG LOADER *****/
let cfgGlobal   = null;
let auditBuffer = [];

function getConfig_() {
  
  if(cfgGlobal) return cfgGlobal;
  const startTime = Date.now();

  const props = PropertiesService.getScriptProperties();
  let snapShot = props.getProperties();
  const ss = SpreadsheetApp.openById(props.getProperty('SHEET_ID'));

  // Cache the Config sheet
  const configSheet = ss.getSheetByName('Config');
  if (!configSheet) throw new Error('Missing "Config" sheet.');

  const values = configSheet.getDataRange().getValues();
  const cfg = objectFromRows_(values);

  // Lock down critical infrastructure
  cfg.SHEET_ID          = snapShot.SHEET_ID;
  cfg.ss                = ss;
  cfg.config            = configSheet;
  cfg.ml                = ss.getSheetByName('MailingList'); 
  cfg.audit             = ss.getSheetByName('AuditLog');
  cfg.ROSTER_FILE_ID    = snapShot.ROSTER_FILE_ID;
  cfg.ROSTER_SHEET_NAME = snapShot.ROSTER_SHEET_NAME;
  cfg.state             = snapShot.state;

  // Defaults
  cfg.LIST_COLUMN       = cfg.LIST_COLUMN || 'QTC';
  cfg.SENDER_EMAIL      = Session.getActiveUser().getEmail();
  cfg.SENDER_NAME       = cfg.SENDER_NAME || 'Midlothian Friends Meeting';
  cfg.TEST_RECIPIENT    = cfg.TEST_RECIPIENT || cfg.SENDER_EMAIL;
  cfg.DELAY_MS          = Number(cfg.DELAY_MS || 1500);
  cfg.BATCH_LIMIT       = Number(cfg.BATCH_LIMIT || 500);

  // Load Gmail draft template (replaces exportDocHtml_ / DOC_TEMPLATE_ID)
  const tpl             = getGmailTemplate_();
  cfg.templateHtmlBody  = tpl.htmlBody;
  cfg.SUBJECT           = tpl.subject;   // dynamic from draft subject line

  let anyTriggers = hasTrigger();

  // Mirror buffer + headers
  const mirrorValues = cfg.ml.getDataRange().getDisplayValues();
  cfg.mirrorBuffer   = mirrorValues;
  cfg.headers        = mirrorValues[0];
  
  const sendCol = cfg.headers.indexOf('Send?');
  const timeCol = cfg.headers.indexOf('Timestamp');
  let remainingToSend = cfg.mirrorBuffer.filter((row, i) => i > 0 && row[sendCol] === 'TRUE').length;
  const listCount = cfg.mirrorBuffer.length - 1;
  
  let update = {
        currentUser: Session.getActiveUser().getEmail(),
        listCount: String(listCount),
        remainingToSend: String(remainingToSend)
      };
  props.setProperties(update);

  // New snapShot needed
  snapShot = props.getProperties();

  // Guard against unintended restart
  if(isRestartOkay(cfg.mirrorBuffer, sendCol, timeCol) === false && hasTrigger() === true) {
      clearResumeTriggers();
      if(listCount === remainingToSend) {
        if(snapShot.timestampStart !== undefined) props.deleteProperty('timestampStart');
        if(snapShot.timestampLast != undefined) props.deleteProperty('timestampLast');
      }
      msg = 'State ' + cfg.state + ' ==> Ready Restart denied in getConfig_()';
      props.setProperty('state', 'Ready');
      const sheet = cfg.ss.getSheetByName('eventStateLog');
      sheet.appendRow([ timeStamp(), state, msg || "", cfg.SENDER_EMAIL ]);

      // Clean up stray properties
      if(hasTrigger() === false) {
          if(snapShot.pauseStart !== undefined ) props.deleteProperty('pauseStart');
          if(snapShot.restartPending !== undefined ) props.deleteProperty('restartPending');
          if(snapShot.nextResume !== undefined) props.deleteProperty('nextResume');
          if(snapShot.resumeType !== undefined) props.deleteProperty('resumeType');
        };
  }

  cfgGlobal = cfg;
  return cfg;
}

/***** GMAIL DRAFT TEMPLATE *****/

/**
 * Finds a Gmail draft whose subject starts with one or more '*'
 * optionally followed by ':' and whitespace.
 * Examples that all match:
 *   "*: Quaker Town Crier - March 2026"
 *   "**: Subject"
 *   "* Subject"
 * Returns { subject, htmlBody, plainBody }
 */
function getGmailTemplate_() {
  const re = /^\*+:?\s*/;  // one or more *, optional colon, optional spaces
  const drafts = GmailApp.getDrafts();
  const matches = drafts.filter(d => re.test(d.getMessage().getSubject()));

  if (matches.length === 0) {
    throw new Error(
      'NO_TEMPLATE: No Gmail draft template found. ' +
      'Create a draft in Gmail whose subject line starts with "*:" ' +
      'followed by the actual email subject, ' +
      'e.g. "*: Quaker Town Crier \u2013 March 2026".'
    );
  }
  if (matches.length > 1) {
    const subjects = matches.map(d => '"' + d.getMessage().getSubject() + '"').join(', ');
    throw new Error(
      'MULTI_TEMPLATE: Multiple "*:" drafts found: ' + subjects + '. ' +
      'Please delete the unneeded drafts and keep only one.'
    );
  }

  const msg     = matches[0].getMessage();
  const subject = msg.getSubject().replace(re, '');
  return {
    subject,
    htmlBody:  msg.getBody(),
    plainBody: msg.getPlainBody()
  };
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

/***** MERGE + RENDER *****/

function renderDocToBlobs_(data) {
  const cfg = getConfig_();

  let htmlBody = cfg.templateHtmlBody;

  // Merge {{tokens}} as string replacements
  Object.keys(data).forEach(k => {
    const val = data[k] == null ? '' : String(data[k]);
    const re  = new RegExp('\\{\\{\\s*' + escapeRegex_(k) + '\\s*\\}\\}', 'g');
    htmlBody  = htmlBody.replace(re, val);
  });

  // Plain-text fallback from merged HTML
  let plainBody = htmlBody
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  // Footer
  const preparedNoteHtml = cfg.PREPARED_BY
    ? `<p>This mailing was prepared by ${escapeHtml_(cfg.PREPARED_BY)} on behalf of the Clerk of the Meeting.</p>`
    : '';

  const unsubscribeHtml =
    `<p>To unsubscribe: <a href="mailto:news+unsubscribe@midlothianfriends.org">news+unsubscribe@midlothianfriends.org</a></p>`;

  htmlBody = htmlBody + preparedNoteHtml + unsubscribeHtml;

  const preparedNoteTxt = cfg.PREPARED_BY
    ? `\n\nThis mailing was prepared by ${cfg.PREPARED_BY} on behalf of the Clerk of the Meeting.\n`
    : '';

  const unsubscribeTxt = `\nTo unsubscribe: news+unsubscribe@midlothianfriends.org`;

  // Normalize paragraph/list spacing
  htmlBody = htmlBody
    .replace(/<p(?![^>]*\sstyle=)([^>]*)>/gi, '<p$1 style="margin-bottom:0.3em; line-height:1.3;">')
    .replace(/<li(?![^>]*\sstyle=)([^>]*)>/gi, '<li$1 style="margin-bottom:0.3em; line-height:1.3;">');

  plainBody = plainBody + preparedNoteTxt + unsubscribeTxt;

  return { htmlBody, plainBody };
}

/***** SENDING *****/

function sendOne_({to, subject, htmlBody, plainBody, menu, cfg}) {
  checkAccess();
  try {
    const opts = { name: cfg.SENDER_NAME, htmlBody };
    gmailSender(to, subject, plainBody, opts);
    bufferAudit({ action: menu, recipient: to, result: 'Sent' });
    return true;
  } catch (err) {
    const msg = `Error ${err.message}  ${err.stack}`;
    bufferAudit({ action: menu, recipient: to, result: msg });
    return false;
  }
}

/***** MENUS *****/

// 1) Preview — sends a test email to the current user
function menuPreview() {
  checkAccess();
  const cfg = getConfig_();
  try {  
    let activeUserEmail = cfg.SENDER_EMAIL;
    const emailCol = cfg.headers.indexOf('Email');
    const testRow = cfg.mirrorBuffer.find((row, i) =>
      i > 0 && String(row[emailCol] || '').trim() === activeUserEmail
    );
    if (!testRow) throw new Error('Preview recipient ' + activeUserEmail + ' not found in the MailingList.');
    const data   = rowToObject_(cfg.headers, testRow);
    const merged = renderDocToBlobs_(data);
    sendOne_({
      to: activeUserEmail,
      subject: cfg.SUBJECT,
      htmlBody: merged.htmlBody,
      plainBody: merged.plainBody,
      menu: 'Preview',
      cfg
    });
    flushBuffers_(cfg);
    return { status: true };
  } catch (err) {
    flushBuffers_(cfg);
    return { status: false, msg: err.message };
  }
}

// 2) Build MailingList from Roster based on LIST_COLUMN
function menuRefreshMailingList() {
  checkAccess();
  cfg = getConfig_();
  const props = PropertiesService.getScriptProperties();
  try {
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
    const idxList  = hmap[cfg.LIST_COLUMN];

    if (idxEmail == null) throw new Error('Roster must have an "Email" column.');
    if (idxList  == null) throw new Error('LIST_COLUMN "' + cfg.LIST_COLUMN + '" not found in Roster.');
    
    const out = [];
    for (let r = 1; r < vals.length; r++) {
      const row   = vals[r];
      const email = String(row[idxEmail] || '').trim();
      if (!email) continue;
      if (row[idxList] === false) continue;
      const first = idxFirst != null ? String(row[idxFirst] || '').trim() : '';
      const last  = idxLast  != null ? String(row[idxLast]  || '').trim() : '';
      const phone = idxPhone != null ? String(row[idxPhone] || '').trim() : '';
      const addr  = idxAddr  != null ? String(row[idxAddr]  || '').trim() : '';
      out.push([ email, first, last, phone, addr, 'TRUE', '', '' ]);
    }  
    
    cfg.ml.clearContents();
    const headersOut = ['Email','First name','Last name','Phone','Address','Send?','Status','Timestamp'];
    clearResumeTriggers();
    cfg.ml.getRange(1, 1, 1, headersOut.length).setValues([headersOut]);
    if (out.length) {
      cfg.ml.getRange(2, 1, out.length, headersOut.length).setValues(out);
    }

    const listCount = out.length;
    if (hasTrigger()) clearResumeTriggers();
    snapShot = props.getProperties();
    ['pauseStart','restartPending','nextResume','resumeType','timestampStart','timestampLast'].forEach(key => {
      if (snapShot[key] !== undefined) props.deleteProperty(key);
    });
    props.setProperties({
      currentUser:     cfg?.SENDER_EMAIL || '',
      listCount:       String(listCount ?? 0),
      sent:            String(0),
      remainingToSend: String(listCount ?? 0),
      state:           'Refreshed'
    });
    return { status: true };
  } catch (err) {
    return { status: false, msg: err.message };
  }
}

// 3) Send all from MailingList (throttled, with auto-pause/resume)
function resumeMenuSendAll() {
  menuSendAll(true);
}

function stopSend() {
  PropertiesService.getScriptProperties().setProperty('stopSend', timeStamp());
}

globalValidSendStates = ['Blocked','Complete','Error','Failure','Running','Idle','Paused','Quota','Ready','Refreshed','restartFailure','Stopped'];

function menuSendAll(type) {
  if (PropertiesService.getScriptProperties().getProperty('stopSend')) {
    msg = "A request to stop dispatching emails has been received.";
    if (eventStateLog('Stopped', msg)) return { status: true, msg: msg };
    return { status: false, msg: "eventStateLog failure" };
  }

  if (!isAuthorizedUser()) {
    let msg = Session.getActiveUser().getEmail() + ' is not an authorized sender';
    if (eventStateLog('Blocked', msg)) return { status: true, msg: msg };
    return { status: false, msg: "eventStateLog failure" };
  }
  
  if (getRemainingDailyQuota() < 5) {
    result = getTestMode();
    setResumeType('Quota');
    let delay, delayMsg;
    if (result.mode === true) {
      delay    = 3 * 60 * 1000;
      scheduleResumeMenuSendAll(3 * 60 * 1000);
      testQuotaRefresh(150 * 1000);
      delayMsg = "three minutes";
    } else {
      delay    = 60 * 60 * 1000;
      scheduleResumeMenuSendAll(60 * 60 * 1000);
      delayMsg = "one hour";
    }
    msg = `Quota low (${Counters.estQuota}). Restart scheduled in ${delayMsg}.`;
    if (eventStateLog("Quota", msg)) return { status: true, msg: msg };
    return { status: false, msg: "eventStateLog failure" };
  }

  checkAccess();
  Counters.init(20);
  const cfg = getConfig_();
  props = PropertiesService.getScriptProperties();
  let state            = props.getProperty('state');
  let remainingToSend  = props.getProperty('remainingToSend');
  Counters.remainingToSend = remainingToSend;
  let msg;

  if (type === true && state !== 'Paused' && state !== 'Quota') {
    msg = 'state: ' + state + ' Re-starts are permitted only in validated Paused and Quota states.';
    eventStateLog(null, msg);
    return { status: 'false', msg: msg };
  }

  const emailCol  = cfg.headers.indexOf('Email');
  const sendCol   = cfg.headers.indexOf('Send?');
  const statusCol = cfg.headers.indexOf('Status');
  const timeCol   = cfg.headers.indexOf('Timestamp');

  if (type === 'false' && (state === 'Paused' || state === 'Quota')) {
    msg = 'state: ' + state + ' Dashboard activated startup not permitted in this state';
    if (eventStateLog('Ready', msg)) return { status: 'false', msg: msg };
    return { status: false, msg: "eventStateLog failure" };
  }

  if (remainingToSend == 0) {
    msg = 'state: ' + state + ' There aren\'t any emails queued for dispatch.';
    eventStateLog(null, msg);
    return { status: 'false', msg: msg };
  }

  if (remainingToSend === listCount || state === 'Ready' || state === 'Refreshed') {
    props.setProperty('timestampStart', timeStamp());
  }
  clearResumeTriggers();
  snapShot = props.getProperties();
  ['pauseStart','restartPending','nextResume','resumeType'].forEach(key => {
    if (snapShot[key] !== undefined) props.deleteProperty(key);
  });
  
  msg = (type === true ? "Re-entry" : "Startup") + " State =" + state + ' => Running';
  if (!eventStateLog('Running', msg)) return { status: false, msg: "eventStateLog failure" };

  try {
    let sent = 0;
  
    for (const [i, row] of cfg.mirrorBuffer.entries()) {
      if (i === 0) continue;
      if (row[sendCol] !== 'TRUE') continue;

      const props = PropertiesService.getScriptProperties().getProperties();
      if (props.stopSend) {
        Counters.flush();
        flushBuffers_(cfg);
        msg = "A request to stop dispatching emails has been received.";
        if (eventStateLog('Stopped', msg)) return { status: true, msg: msg };
        return { status: false, msg: "eventStateLog failure" };
      }

      let maxTime = (5 * 60 + 30) * 1000;
      if (Counters.testingMode === true) { maxTime = (60 + 30) * 1000; }

      if ((Date.now() - Counters.startTime) > maxTime) { 
        Counters.flush();
        flushBuffers_(cfg);
        setResumeType('Paused');
        if (scheduleResumeMenuSendAll(60 * 1000)) {
          msg = 'A short pause is required. An automated restart will begin in about one minute';
          if (eventStateLog("Paused", msg)) return { status: true, msg: msg };
          return { status: false, msg: "eventStateLog failure" };
        }
        msg = 'An automated restart could not be set';
        if (eventStateLog("restartFailure", msg)) return { status: false, msg: msg };
        return { status: false, msg: "eventStateLog failure" };
      }

      if (sent >= cfg.BATCH_LIMIT) { 
        Counters.flush();
        flushBuffers_(cfg);
        setResumeType('Paused');
        if (scheduleResumeMenuSendAll(60 * 1000)) {
          msg = 'A short pause is required. An automated restart will begin in about one minute';
          if (eventStateLog("Paused", msg)) return { status: true, msg: msg };
          return { status: false, msg: "eventStateLog failure" };
        }
        msg = 'An automated restart could not be set';
        if (eventStateLog("restartFailure", msg)) return { status: false, msg: msg };
        return { status: false, msg: "eventStateLog failure" };
      }

      if (Counters.estQuota < 5) {
        Counters.flush();
        flushBuffers_(cfg);
        setResumeType('Quota');
        let delay, delayMsg;
        if (Counters.testingMode === true) {
          delay    = 3 * 60 * 1000;
          delayMsg = "three minutes";
        } else {
          delay    = 60 * 60 * 1000;
          delayMsg = "one hour";
        }
        if (scheduleResumeMenuSendAll(delay)) {
          if (Counters.testingMode === true) { testQuotaRefresh(150 * 1000); }
          msg = `Quota low (${Counters.estQuota}). Restart scheduled in ${delayMsg}.`;
          if (eventStateLog("Quota", msg)) return { status: true, msg: msg };
          return { status: false, msg: "eventStateLog failure" };
        }
        msg = 'An automated restart could not be set';
        if (eventStateLog("restartFailure", msg)) return { status: false, msg: msg };
        return { status: false, msg: "eventStateLog failure" };
      }

      const to     = String(row[emailCol] || '').trim();
      const data   = rowToObject_(cfg.headers, row);
      const merged = renderDocToBlobs_(data);
      const result = sendOne_({
        to,
        subject:   cfg.SUBJECT,
        htmlBody:  merged.htmlBody,
        plainBody: merged.plainBody,
        menu:      'SendAll',
        cfg
      });

      if (result === true) {
        sent++;
        Counters.bump();
        Utilities.sleep(1000 + Math.floor(Math.random() * 1000));
        row[statusCol] = 'Sent';
        row[timeCol]   = timeStamp();
        row[sendCol]   = 'FALSE';
      } else {
        Counters.flush();
        flushBuffers_(cfg);
        msg = 'An unexpected error occurred while sending an email. Processing will restart after review.';
        if (eventStateLog("Error", msg)) return { status: false, msg: msg };
        return { status: false, msg: "eventStateLog failure" };
      }
    }

    Counters.flush();
    flushBuffers_(cfg);
    msg = 'All emails in the mailing list have been dispatched';
    if (eventStateLog("Complete", msg)) return { status: true, msg: msg };
    return { status: false, msg: "eventStateLog failure" };

  } catch (err) {
    flushBuffers_(cfg);
    if (eventStateLog("Failure", err.message)) return { status: false, msg: err.message };
    return { status: false, msg: 'eventStateLog failure ' + err.message };
  }
}

/***** VALIDATOR *****/

/**
 * Validates the Gmail draft template.
 * Returns { status, previewHtml, subject } on success.
 * Returns { status: false, msg } if no/multiple templates found or no tokens.
 * The previewHtml shows the email body with {{tokens}} highlighted in green
 * and a banner indicating validity.
 */
function menuValidateTemplate() {
  checkAccess();
  try {
    const tpl = getGmailTemplate_();

    // Extract {{tokens}} from the plain body
    const re     = /\{\{\s*([^}]+?)\s*\}\}/g;
    const tokens = [];
    let m;
    while ((m = re.exec(tpl.plainBody)) !== null) {
      if (!tokens.includes(m[1])) tokens.push(m[1]);
    }

    // Highlight tokens green in the HTML body
    const highlighted = tpl.htmlBody.replace(
      /\{\{\s*([^}]+?)\s*\}\}/g,
      '<span style="color:#2e7d32;font-weight:bold;">{{$1}}</span>'
    );

    const hasTokens   = tokens.length > 0;
    const tokenList   = tokens.map(t => '<code>{{' + t + '}}</code>').join(', ');
    const bannerColor = hasTokens ? '#2e7d32' : '#c62828';
    const bannerMsg   = hasTokens
      ? '&#10003; Template is valid. Tokens found: ' + tokenList +
        '<br><small>To modify: edit the draft in Gmail and re-validate.</small>'
      : '&#9888; No {{tokens}} found. Add merge fields like <code>{{First name}}</code> to your Gmail draft.';

    const previewHtml =
      '<div style="font-family:Arial,sans-serif;padding:12px;border:2px solid ' + bannerColor + ';' +
      'border-radius:8px;margin-bottom:12px;">' +
        '<div style="color:' + bannerColor + ';font-weight:bold;margin-bottom:8px;">' + bannerMsg + '</div>' +
      '</div>' +
      '<hr>' +
      highlighted;

    if (hasTokens) {
      PropertiesService.getScriptProperties().setProperty('validTemplate', timeStamp());
    }

    return { status: hasTokens, previewHtml, subject: tpl.subject };

  } catch (err) {
    return { status: false, msg: err.message.replace(/^(NO_TEMPLATE|MULTI_TEMPLATE):\s*/, '') };
  }
}
