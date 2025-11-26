function testSheetAccess() {
  const sheet = SpreadsheetApp.openById(EMAIL_MERGE_SHEET_ID).getSheetByName('AuditLog'); // replace 'Sheet1' with your actual tab name
  const value = sheet.getRange(1, 1).getValue(); // reads cell A1
  Logger.log(`A1 value: ${value}`);
}
