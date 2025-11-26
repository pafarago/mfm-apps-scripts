function doGet(e) {
  assertAccess_();  // guardrail
  return HtmlService.createTemplateFromFile('Dashboard').evaluate()
    .setTitle('MailMerge Dashboard');
}

// Preview handler
function previewEmails() {
  const cfg = getConfig_();
  // generate sample output but don’t send
  bufferMailingList({rowIndex: 2, status: 'Dry Run'});
  flushBuffers_(cfg);
  return {
    sent: 0,
    quota: MailApp.getRemainingDailyQuota(),
    message: 'Preview complete. Check AuditLog for details.'
  };
}

// Send handler
function sendAll_(dryRun) {
  assertAccess_();
  const cfg = getConfig_();
  let sent = 0;

  // simplified loop
  for (let i = 2; i <= 5; i++) {
    try {
      if (dryRun) {
        bufferMailingList({rowIndex: i, status: 'Dry Run'});
      } else {
        // sendOne_ logic here
        bufferMailingList({rowIndex: i, status: 'Sent'});
        sent++;
      }
    } catch (e) {
      bufferMailingList({rowIndex: i, status: 'ERROR: ' + e.message});
    }
  }

  flushBuffers_(cfg);

  return {
    sent: sent,
    quota: MailApp.getRemainingDailyQuota(),
    message: dryRun ? 'Dry run complete.' : 'Sending finished.'
  };
}

