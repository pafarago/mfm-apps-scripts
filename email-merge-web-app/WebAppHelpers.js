// Access restriction helper
function checkAccess() {
  try {
    const email = Session.getActiveUser().getEmail();
    const allowed = ['admin@midlothianfriends.org','clerk@midlothianfriends.org'];
    if (allowed.indexOf(email) === -1) {
      throw new Error('Access denied to ' + email);
    }
    return { status: true };
  } catch (err) {
    return { status: false , msg: err.message};
  }
}

function getMergeInfo() {
  const props = PropertiesService.getScriptProperties().getProperties();
  return {
    currentUser: props.currentUser || '',
    remainingDailyQuota: props.remainingDailyQuota || '0',
    listCount: props.listCount || '0',
    sent: props.sent || '0',
    remainingToSend: props.remainingToSend || '0',
    timestampStart: props.timestampStart || '',
    timestampLast: props.timestampLast || '',
    restartPending: props.restartPending || 'No',
    state: props.state || 'Idle',
    nextResume: props.nextResume || ''
  };
}

function timeStamp(ms) {
/**
      Format:                       Mon Nov 24 2025 15:24:00 GMT-0500
      timeStamp();                  current time
      timeStamp(Date.now()));       same as above
      timeStamp(1769324640000));    Mon Nov 24 2025 15:24:00 GMT-0500
*/
  const tz = Session.getScriptTimeZone();   // Apps Script timezone
  const d  = ms ? new Date(ms) : new Date();
  return Utilities.formatDate(d, tz, "EEE MMM dd yyyy HH:mm:ss 'GMT'Z");
}

function hasTrigger() {
  return ScriptApp.getProjectTriggers()
    .some(t => t.getHandlerFunction() === 'resumeMenuSendAll');
}

function clearResumeTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'resumeMenuSendAll') {
      ScriptApp.deleteTrigger(t);
    }
  });
}
function scheduleResumeMenuSendAll(milliseconds)
{
  // process in r=decreasing order of milliseconds (Quota before Pending)
  const props = PropertiesService.getScriptProperties()
  clearResumeTriggers();
  ScriptApp.newTrigger('resumeMenuSendAll')
  .timeBased()
  .after(milliseconds)
  .create();
  const futureMs   = Date.now() + milliseconds;
  nextResume = timeStamp(futureMs);
  props.setProperty("nextResume", nextResume);
}

function setNextResume(triggerTime) {
  const props = PropertiesService.getScriptProperties();
  let current = props.getProperty('nextResume');

  // Only overwrite if we have a real timestamp
  if (triggerTime) {
    props.setProperty('nextResume', String(triggerTime));
  } else if (!current) {
    // Initialize if missing
    props.setProperty('nextResume', '');
  }
}
function resumeMenuSendAll() {
  menuSendAll(true);
}

function incrementSent() {
  const props = PropertiesService.getScriptProperties();
  let current = Number(props.getProperty('sent')) || 0;
  props.setProperty('sent', String(Math.floor(current + 1)));
}

function setResumeType(type) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('resumeType',type );
}

// Lightweight status for dashboard startup
function remainingToSend_() {
  const cfg = getConfig_();
  const sendCol = cfg.headers.indexOf('Send?');
  let remainingToSend = cfg.mirrorBuffer.filter((row, i) =>
    i > 0 && row[sendCol] === 'TRUE' // skip header row
  ).length;
  return (remainingToSend );
}
function listCount() {
  const cfg = getConfig_();
  listCount = cfg.mirrorBuffer.length - 1;
  return (listCount);
}

function eventStateLog(state, msg) {
  const cfg = getConfig_();
  const sheet = cfg.ss.getSheetByName('eventStateLog');
  const stamp = timeStamp();
  PropertiesService.getScriptProperties().setProperty('state', state);
  sheet.appendRow([ stamp, state, msg || "", cfg.SENDER_EMAIL ]);
}

function bufferAudit({action, recipient, result}) {
  const stamp = timeStamp();
  auditBuffer.push([stamp, action, recipient, result]);
}

function flushBuffers_(cfg) {
  // --- Flush MailingList ---
  if (cfg.mirrorBuffer && cfg.mirrorBuffer.length > 0) {
    const mlRange = cfg.ml.getRange(1, 1, cfg.mirrorBuffer.length, cfg.mirrorBuffer[0].length);
    mlRange.setValues(cfg.mirrorBuffer);
    cfg.mirrorBuffer.length = 0 // clear the buffer in place
  }

  // --- Flush AuditLog ---
  if (auditBuffer && auditBuffer.length > 0) {
    const sheet = cfg.audit;
    const startRow = sheet.getLastRow() + 1; // append after last row
    sheet.getRange(startRow, 1, auditBuffer.length, auditBuffer[0].length)
         .setValues(auditBuffer);
    auditBuffer.length = 0; // clear global buffer in place
  }
}

function objectFromRows_(rows) {
  const obj = {};
  // Skip header row, assume first column is "Key" and second is "Value"
  rows.slice(1).forEach(r => {
    const key = String(r[0] || '').trim();
    if (key && key !== 'Key') {
      obj[key] = String(r[1] ?? '').trim();
    }
  });
  return obj;
}

function rowToObject_(headers, row) {
  const obj = {};
  headers.forEach((h, idx) => {
    obj[h] = row[idx];
  });
  return obj;
}
