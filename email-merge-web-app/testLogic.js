
function testQuotaRefresh() {
  PropertiesService.getScriptProperties().setProperty('testQuota', '100');
}

function testQuotaStop() {
  PropertiesService.getScriptProperties().deleteProperty('testQuota');
}

function gmailSender(to, subject, plainBody, opts) {
  const props = PropertiesService.getScriptProperties();
  let snapShot = props.getProperties();
  if(snapShot.testQuota) { // test mode
    Utilities.sleep(500);
    let quota = Number(snapShot.testQuota);
    if(quota > 0) quota--;
    props.setProperty('testQuota', String(quota));
  } else if (isAuthorizedUser()){
    GmailApp.sendEmail(to, subject, plainBody, opts);
  } else {
    let msg = Session.getActiveUser().getEmail() + ' is not a authorized sender';
    eventStateLog('Blocked', msg); // changes state, clear time triggers, manages properties, and logs the event
  }
}

function getRemainingDailyQuota() {
  const props = PropertiesService.getScriptProperties();
  let snapShot = props.getProperties();
  if(snapShot.testQuota) {
    return(snapShot.testQuota);
  } else {
   return(MailApp.getRemainingDailyQuota()); 
  }
}
  
function getTestMode() {
  const props = PropertiesService.getScriptProperties();
  let snapShot = props.getProperties();
  if(snapShot.testQuota) {
    return({ mode: true , msg: 'Test Mode, no emails will be sent'});
  } else {
   return({ mode: false, msg: 'All remaining emails will be sent'}); 
  }
}

function setQuotaRefreshTrigger(milliseconds)
{
  ScriptApp.newTrigger('testQuotaRefresh')
  .timeBased()
  .after(milliseconds)
  .create();
  const futureMs   = Date.now() + milliseconds;
  eventStateLog(null, 'testQuota referesh trigger set ' + timeStamp(furuteMS))
}