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

    const hours   = Math.floor(elapsedMs / 3600000);
    const minutes = Math.floor((elapsedMs % 3600000) / 60000);
    const seconds = Math.floor((elapsedMs % 60000) / 1000);
    const millis  = elapsedMs % 1000;

    elapsedTime = Utilities.formatString(
      "%02d:%02d:%02d.%03d",
      hours, minutes, seconds, millis
    );
  } else {
    elapsedTime = '00:00:00.000';
  }

  let runningTime;
  if (props.runningStart) {
    let runningMs = new Date(props.runningStart).getTime();
    
    const now = new Date();
    const elapsedMs = now.getTime() - runningMs;

    const hours   = Math.floor(elapsedMs / 3600000);
    const minutes = Math.floor((elapsedMs % 3600000) / 60000);
    const seconds = Math.floor((elapsedMs % 60000) / 1000);
    const millis  = elapsedMs % 1000;

    runningTime = Utilities.formatString(
      "%02d:%02d:%02d.%03d",
      hours, minutes, seconds, millis
    );
  } else {
    runningTime = '00:00:00.000';
  }

  const mode = getTestMode();

  return {
    currentUser:          Session.getActiveUser().getEmail(),
    remainingDailyQuota:  getRemainingDailyQuota() || '0',
    listCount:            props.listCount      || '0',
    sent:                 props.sent           || '0',
    remainingToSend:      props.remainingToSend || '0',
    testingMode:          mode.msg             || '',
    timestampStart:       props.timestampStart || '',
    timestampLast:        props.timestampLast  || '',
    restartPending:       props.restartPending || 'No',
    state:                props.state          || 'Idle',
    nextResume:           props.nextResume     || '',
    elapsedTime:          elapsedTime          || '',
    runningTime:          runningTime          || '',
  };
}

function timeStamp(ms) {
/**
      Format:                       Mon Nov 24 2025 15:24:00 GMT-0500
      timeStamp();                  current time
      timeStamp(Date.now()));       same as above
      timeStamp(1769324640000));    Mon Nov 24 2025 15:24:00 GMT-0500
*/
  const tz = Session.getScriptTimeZone();
  const d  = ms ? new Date(ms) : new Date();
  return Utilities.formatDate(d, tz, "EEE MMM dd yyyy HH:mm:ss 'GMT'Z");
}

/**
 * Determines if the mailing list buffer represents a partially processed
 * send that is safe to restart.
 *
 * A valid restart state is:
 *   - One or more consecutive rows with Send? === "FALSE" and a valid timestamp
 *   - Followed by one or more rows with Send? === "TRUE" and a blank timestamp
 *
 * @param {Array[]} buffer - 2D array of rows (first row is header).
 * @param {number}  sendCOL - Index of the Send? column.
 * @param {number}  timeCOL - Index of the Timestamp column.
 * @return {boolean} - True if a restart is safe, false otherwise.
 */
function isRestartOkay(buffer, sendCOL, timeCOL) {
  const tsRegex = /^[A-Za-z]{3} [A-Za-z]{3} \d{2} \d{4} \d{2}:\d{2}:\d{2} GMT[+-]\d{4}$/;

  // Must have at least one data row
  if (buffer.length < 2) return false;

  // First data row must be FALSE with a valid timestamp (partial send has started)
  const firstRow = buffer[1];
  if (firstRow[sendCOL] !== 'FALSE' || !tsRegex.test(String(firstRow[timeCOL]).trim())) {
    return false;
  }

  let seenTrue = false;

  for (let i = 1; i < buffer.length; i++) {
    const sendVal = String(buffer[i][sendCOL]).trim();
    const timeVal = String(buffer[i][timeCOL]).trim();

    if (!seenTrue) {
      if (sendVal === 'FALSE' && tsRegex.test(timeVal)) {
        continue;               // still in the "already sent" block — good
      } else if (sendVal === 'TRUE' && timeVal === '') {
        seenTrue = true;        // transition to "remaining to send" block
      } else {
        return false;           // unexpected pattern
      }
    } else {
      // Once we've seen a TRUE, all remaining rows must also be TRUE with blank timestamp
      if (sendVal !== 'TRUE' || timeVal !== '') {
        return false;
      }
    }
  }

  // Must have seen both blocks — sent rows AND rows still to send
  return seenTrue;
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
  eventStateLog(null, 'menuSendAll triggers cleared');
}

function scheduleResumeMenuSendAll(milliseconds) {
  const props = PropertiesService.getScriptProperties();
  clearResumeTriggers();
  if (setResumeMenuSendAllTimeTrigger(milliseconds)) {
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
  const current = Number(props.getProperty('sent')) || 0;
  props.setProperty('sent', String(Math.floor(current + 1)));
}

function setResumeType(type) {
  PropertiesService.getScriptProperties().setProperty('resumeType', type);
}

// Lightweight status for dashboard startup
function remainingToSend_() {
  const cfg     = getConfig_();
  const sendCol = cfg.headers.indexOf('Send?');
  return cfg.mirrorBuffer.filter((row, i) => i > 0 && row[sendCol] === 'TRUE').length;
}

function listCount() {
  const cfg   = getConfig_();
  const count = cfg.mirrorBuffer.length - 1;
  return count;
}

function logOnly(user, sheet, stamp, state, msg) {
  if (sheet) {
    try {
      sheet.appendRow([ stamp, state, msg || "", user ]);
    } catch(error) {
      Logger.log(stamp + ' state = ' + state + ' ' + (msg || "No message") + ' ' + error.message);
    }
  } else {
    Logger.log(stamp + ' state = ' + state + ' ' + (msg || "No message"));
  }
}

function eventStateLog(state, msg) {
  let sheet;
  let user;
  const props = PropertiesService.getScriptProperties();
  const stamp = timeStamp();

  if (cfgGlobal && cfgGlobal.ss) {
    sheet = cfgGlobal.ss.getSheetByName('eventStateLog');
    user  = cfgGlobal.SENDER_EMAIL;
  } else {
    user     = 'news@midlothianfriends.org';
    const ss = SpreadsheetApp.openById(props.getProperty('SHEET_ID'));
    if (ss) {
      sheet = ss.getSheetByName('eventStateLog');
    }
  }

  if (state !== null) {
    if (!globalValidSendStates.includes(state)) {
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
    const stateArray = ['Blocked','Complete','Error','Failure','Idle','Ready','Refreshed','restartFailure','Stopped'];
    if (stateArray.includes(state)) {
      if (hasTrigger()) { clearResumeTriggers(); }
      ['pauseStart','restartPending','nextResume','resumeType','stopSend'].forEach(key => {
        props.deleteProperty(key);
      });
      props.setProperty('state', state);
      logOnly(user, sheet, stamp, state, msg);
      if (state === 'restartFailure') {
        logOnly(user, sheet, stamp, state, "recording state transition");
        props.setProperty('state', 'Ready');
        logOnly(user, sheet, stamp, 'Ready', "return control to the Dashboard");
      }
      return true;
    }
  } else {
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
    cfg.mirrorBuffer.length = 0;
    eventStateLog(null, 'Mailing list buffer flushed.');
  }

  // --- Flush AuditLog ---
  if (auditBuffer && auditBuffer.length > 0) {
    const sheet    = cfg.audit;
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, auditBuffer.length, auditBuffer[0].length)
         .setValues(auditBuffer);
    auditBuffer.length = 0;
  }
  eventStateLog(null, 'Audit buffer flushed.');
}

function objectFromRows_(rows) {
  const obj = {};
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
