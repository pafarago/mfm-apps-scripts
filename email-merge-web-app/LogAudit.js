function logAudit_({action, recipient, result, cfg}) {
  const sheet = SpreadsheetApp.openById(cfg.SHEET_ID).getSheetByName('AuditLog');
  sheet.appendRow([
    new Date().toLocaleString(),
    Session.getActiveUser().getEmail(),
    action,
    recipient,
    result,
    GmailApp.getRemainingDailyQuota(),
    cfg.SUBJECT,
    cfg.PREPARED_BY
  ]);
}