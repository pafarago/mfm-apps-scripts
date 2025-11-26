const Counters = {
  sent: 0,
  remainingToSend: 0,
  estQuota: 0,
  batchSize: 20,
  startTime: 0,
  breakReason: 'NOTSET',

  init(n) {
    const props = PropertiesService.getScriptProperties().getProperties();
    this.sent = Number(props.sent) || 0;
    this.remainingToSend = Number(props.remainingToSend) || 0;
    this.estQuota = MailApp.getRemainingDailyQuota();
    this.batchSize = n || 20;
    this.startTime = new Date().getTime();
    this.timestampStart = timeStamp(this.startTime);
  },

  bump() {
    this.sent += 1;
    this.remainingToSend = this.remainingToSend > 0 ? this.remainingToSend - 1 : 0;
    this.estQuota = this.estQuota > 0 ? this.estQuota - 1 : 0;

    if (this.sent % this.batchSize === 0) {
      this.flush();
    }
    return this.estQuota;
  },

  flush() {
    this.estQuota = MailApp.getRemainingDailyQuota();
    PropertiesService.getScriptProperties().setProperties({
      sent: String(this.sent),
      remainingToSend: String(this.remainingToSend),
      remainingDailyQuota: String(this.estQuota),
      timestampLast: timeStamp()
    });
  }
};
/**
function sendAll() {
  Counters.init(20);

  let breakReason = null;

  for (let i = 0; i < recipients.length; i++) {
    try {
      // --- send email logic here ---
      const quota = Counters.bump();

      // Break conditions visible inline
      if (quota < 5) {
        breakReason = "QUOTA";
        Counters.flush();
        break;
      }

      const elapsed = (new Date().getTime() - Counters.startTime) / 1000;
      if (elapsed > 5.5) {
        breakReason = "TIME";
        Counters.flush();
        break;
      }

      const props = PropertiesService.getScriptProperties().getProperties();
      if (props.stopSend === 'true') {
        breakReason = "STOP";
        Counters.flush();
        break;
      }

    } catch (err) {
      breakReason = "ERROR";
      Counters.flush();
      PropertiesService.getScriptProperties().setProperty("lastError", err.message);
      break;
    }
  }

  Counters.flush();

  // React appropriately
  switch (breakReason) {
    case "QUOTA":
      scheduleTrigger(60); // one hour
      break;
    case "TIME":
      scheduleTrigger(1);  // one minute
      break;
    case "ERROR":
      // handle error differently
      break;
    case "STOP":
      // manual stop, no auto restart
      break;
    default:
      // finished normally
      break;
  }
*/