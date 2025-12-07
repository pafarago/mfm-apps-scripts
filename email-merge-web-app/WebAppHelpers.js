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
  let elapsedTime;
  const props = PropertiesService.getScriptProperties().getProperties();
  if (props.pauseStart) {
    let pauseMs = new Date(props.pauseStart).getTime();
    
    const now = new Date();
    const elapsedMs = now.getTime() - pauseMs;

    const hours = Math.floor(elapsedMs / 3600000);
    const minutes = Math.floor((elapsedMs % 3600000) / 60000);
    const seconds = Math.floor((elapsedMs % 60000) / 1000);
    const millis = elapsedMs % 1000;

    elapsedTime = Utilities.formatString(
      "%02d:%02d:%02d.%03d",
      hours, minutes, seconds, millis
    );
  }
  else {
    elapsedTime = '00:00:00.000';
  }

  if (props.runningStart) {
    let runningMs = new Date(props.runningStart).getTime();
    
    const now = new Date();
    const elapsedMs = now.getTime() - runningMs;

    const hours = Math.floor(elapsedMs / 3600000);
    const minutes = Math.floor((elapsedMs % 3600000) / 60000);
    const seconds = Math.floor((elapsedMs % 60000) / 1000);
    const millis = elapsedMs % 1000;

    runningTime = Utilities.formatString(
      "%02d:%02d:%02d.%03d",
      hours, minutes, seconds, millis
    );
  }
  else {
    runningTime = '00:00:00.000';
  }

  let mode = getTestMode();

  return {
    currentUser: Session.getActiveUser().getEmail(),
    remainingDailyQuota: getRemainingDailyQuota() || '0',
    listCount: props.listCount || '0',
    sent: props.sent || '0',
    remainingToSend: props.remainingToSend || '0',
    testingMode: mode.msg || '',
    timestampStart: props.timestampStart || '',
    timestampLast: props.timestampLast || '',
    restartPending: props.restartPending || 'No',
    state: props.state || 'Idle',
    nextResume: props.nextResume || '',
    elapsedTime: elapsedTime || '',
    runningTime: runningTime || '',
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

/**
 * Determines if buffer represents a partially processed mailing list.
 *
 * @param {Array[]} buffer - 2D array of rows (first row is header).
 * @param {number} sendCOL - Index of send column.
 * @param {number} timeCOL - Index of time column.
 * @return {boolean} - True if partially processed, false otherwise.
 */
function isRestartOkay(buffer, sendCOL, timeCOL) {
  // Regex for Utilities.formatDate(d, tz, "EEE MMM dd yyyy HH:mm:ss 'GMT'Z")
  const tsRegex = /^[A-Za-z]{3} [A-Za-z]{3} \d{2} \d{4} \d{2}:\d{2}:\d{2} GMT[+-]\d{4}$/;

  let seenTimestampBlock = "FALSE";

  for (let i = 1; i < buffer.length; i++) { // skip header
    let row = buffer[i];
    let sendVal = row[sendCOL];
    let timeVal = row[timeCOL];
    if(i == 1 && (String(timeVal).trim() === "")) {
        return false;
    }
    if (!seenTimestampBlock) {
      // Expect FALSE + valid timestamp until we hit the first TRUE
      if (sendVal === "FALSE" && typeof timeVal === "string" && tsRegex.test(timeVal.trim())) {
        continue;
      } else if (sendVal === "TRUE" && (timeVal == String(timeVal).trim() === "")) {
        // Transition point: now we expect only TRUE + blank
        seenTimestampBlock = "TRUE";
      } else {
        return false; // invalid pattern
      }
    }
    return true;
  }

  // Must have seen at least one timestamp block and one remaining block
  return seenTimestampBlock;
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
    eventStateLog(null, 'menuSendAll triggers cleared')
  });
}
function scheduleResumeMenuSendAll(milliseconds)
{
  // process in decreasing order of milliseconds (Quota before Pending)
  const props = PropertiesService.getScriptProperties();
  clearResumeTriggers();
  /**
  ScriptApp.newTrigger('resumeMenuSendAll')
  .timeBased()
  .after(milliseconds)
  .create();
  **/
  if(setResumeMenuSendAllTimeTrigger(milliseconds)) {
    const futureMs   = Date.now() + milliseconds;
    const nextResume = timeStamp(futureMs);
    props.setProperty("nextResume", nextResume);
    eventStateLog(null, 'menuSendAll trigger set');
    return true;
  } 
  return false;
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

function logOnly(user, sheet, stamp, state, msg) {
  if(sheet) {
    try {
        sheet.appendRow([ stamp, state, msg || "", user ]);
      } catch(error) {
        Logger.log( stamp + ' state = ' + state + ' ' + (msg || "No message") + ' ' + error.message);
      }
    } else {
      Logger.log(stamp + ' state = ' + state + ' ' + (msg || "No message"));
    }
}

function eventStateLog(state, msg) {
  // New lightweight version without getConfig_() overhead
  let sheet;
  let user;
  const props = PropertiesService.getScriptProperties();
  const stamp = timeStamp();

  if(cfgGlobal && cfgGlobal.ss) {
    sheet = cfgGlobal.ss.getSheetByName('eventStateLog');
    user = cfgGlobal.SENDER_EMAIL;
  } else {
      user = 'news@midlothianfriends.org';
      const ss = SpreadsheetApp.openById(props.getProperty('SHEET_ID'));
      if(ss) {
        sheet = ss.getSheetByName('eventStateLog');
      }
  }

  if (state !== null) {
    if(!globalValidSendStates.includes(state)) {
      logOnly(user, sheet, stamp, state, 'Invalid state');
      return false;
    }
    if (state === 'Paused' || state === 'Quota') {
      props.setProperty('pauseStart', stamp);
      props.setProperty('state', state);
      logOnly(user, sheet, stamp, state, msg);
      return true;
    } else {
      props.deleteProperty('pauseStart');
    }
    if (state === 'Running') {
      props.setProperty('runningStart', stamp);
      props.setProperty('state', state);
      logOnly(user, sheet, stamp, state, msg);
      return true;
    } else {
      props.deleteProperty('runningStart');
    }
    let stateArray = ['Blocked', 'Complete', 'Error', 'Failure', 'Idle', 'Ready', 'Refreshed', 'restartFailure', 'Stopped'];
    if (stateArray.includes(state)) {
      if(hasTrigger()) {clearResumeTriggers();}
      ['pauseStart','restartPending','nextResume','resumeType', 'stopSend'].forEach(key => {
        props.deleteProperty(key);
      });
      props.setProperty('state', state);
      logOnly(user, sheet, stamp, state, msg);
      if (state === 'restartFailure') {
        logOnly(user,sheet, stamp, state, "recording state transition");
        props.setProperty('state', 'Ready'); // Returns to requiring manual re-start in Dashboard
        logOnly(user, sheet, stamp, 'Ready', "return control to the Dashboard");
      }
      return true;
    }
  } else { // No state or properety changes
      logOnly(user, sheet, stamp, 'Info', msg);
      return true;
  }
  return false; // should not occur
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
    eventStateLog(null, 'Mailing list buffer flushed.')
  }

  // --- Flush AuditLog ---
  if (auditBuffer && auditBuffer.length > 0) {
    const sheet = cfg.audit;
    const startRow = sheet.getLastRow() + 1; // append after last row
    sheet.getRange(startRow, 1, auditBuffer.length, auditBuffer[0].length)
         .setValues(auditBuffer);
    auditBuffer.length = 0; // clear global buffer in place
  }
  eventStateLog(null, 'Audit buffer flushed.')
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
