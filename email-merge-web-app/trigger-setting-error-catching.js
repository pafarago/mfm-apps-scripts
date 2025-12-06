function setResumeMenuSendAllTimeTrigger(ms) {
  const start = Date.now();
  while (Date.now() - start < 5 * 60 * 1000) { // 5 min safety window
    try {
      ScriptApp.newTrigger("resumeMenuSendAll")
               .timeBased()
               .after(ms)
               .create();
      console.log("Trigger scheduled successfully.");
      return true;
    } catch (err) {
      if (err.message.includes("Unable to talk to trigger service")) {
        Utilities.sleep(5000); // wait 5s before retry
        continue;
      } else {
        throw err; // other errors, bail out
      }
    }
  }
  console.log("Failed to schedule trigger within 5 minutes.");
  return false;
}
