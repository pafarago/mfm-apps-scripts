function testAuditLogWrite() {
  const sheet = SpreadsheetApp.openById(EMAIL_MERGE_SHEET_ID).getSheetByName('AuditLog');
  sheet.appendRow([
    new Date(),                         // Timestamp
    Session.getActiveUser().getEmail(), // Sender identity
    'testAuditLogWrite',               // Action label
    'Success'                          // Result
  ]);
}
