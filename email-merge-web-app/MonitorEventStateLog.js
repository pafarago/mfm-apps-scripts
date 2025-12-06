function monitorEventStateLog() {
  const ss = SpreadsheetApp.openById('YOUR_SHEET_ID');
  const sheet = ss.getSheetByName('eventStateLog');
  const lastRows = sheet.getRange(sheet.getLastRow()-10, 1, 10, 5).getValues();

  const troubleStates = ['Failure','Error','Blocked','restartFailure'];
  let alerts = [];

  lastRows.forEach(row => {
    const [timestamp, user, state, msg] = [row[0], row[1], row[2], row[3]];
    if (troubleStates.includes(state)) {
      alerts.push(`${timestamp}: ${state} — ${msg}`);
    }
  });

  if (alerts.length > 0) {
    const body = alerts.join('\n');
    MailApp.sendEmail("8045551234@vtext.com", "EventStateLog Alert", body);
    MailApp.sendEmail("peterfarago@outlook.com", "EventStateLog Alert", body);
  }
}
