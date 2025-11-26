/**
 * MFM Subscription Processor (QTC + DIRECTORY) v2
 *
 * Handles subscribe/unsubscribe via plus-addressed aliases:
 *   QTC (general/news):       news+subscribe@midlothianfriends.org / news+unsubscribe@midlothianfriends.org
 *   DIRECTORY (directory):    directory+subscribe@midlothianfriends.org / directory+unsubscribe@midlothianfriends.org
 *
 * Behavior:
 *   - SUB via list alias toggles only that list column TRUE (creating row if needed).
 *   - UNSUB via DIRECTORY alias: clears DIRECTORY only.
 *   - UNSUB via NEWS/QTC alias: clears QTC AND DIRECTORY.
 *
 * Depends on:
 *   - getConfig_() returning ROSTER_FILE_ID and ROSTER_SHEET_NAME
 *   - Roster sheet having columns: "Email", "QTC", "DIRECTORY" (booleans)
 */

const SUBS_PROCESSED_LABEL = 'Newsletter/Processed';
const CLERK_EMAIL = 'news@midlothianfriends.org'; // or clerk@..., adjust if preferred

// Map each alias pair to its roster column
const LIST_ALIASES = [
  {
    column: 'QTC',
    subscribe: 'news+subscribe@midlothianfriends.org',
    unsubscribe: 'news+unsubscribe@midlothianfriends.org'
  },
  {
    column: 'DIRECTORY',
    subscribe: 'directory+subscribe@midlothianfriends.org',
    unsubscribe: 'directory+unsubscribe@midlothianfriends.org'
  }
];

/** Menu entry: manual run */
function menuProcessSubscriptions() {
  const count = processSubscriptionRequests_();
  SpreadsheetApp.getUi().alert('Processed ' + count + ' subscription/unsubscription request(s).');
}

/**
 * Core processor (target for time-based trigger).
 * Reads relevant emails, updates roster booleans, sends confirmations,
 * and labels processed threads.
 */
function processSubscriptionRequests_() {
  const cfg = getConfig_();
  if (!cfg.ROSTER_FILE_ID || !cfg.ROSTER_SHEET_NAME) {
    throw new Error('Config: ROSTER_FILE_ID and ROSTER_SHEET_NAME are required.');
  }

  const rosterSS = SpreadsheetApp.openById(cfg.ROSTER_FILE_ID);
  const rosterSh = rosterSS.getSheetByName(cfg.ROSTER_SHEET_NAME);
  if (!rosterSh) {
    throw new Error('Roster sheet "' + cfg.ROSTER_SHEET_NAME + '" not found.');
  }

  const vals = rosterSh.getDataRange().getValues();
  if (vals.length < 2) return 0;

  const headers = vals[0].map(h => String(h).trim());
  const hmap = {};
  headers.forEach((h, i) => { if (h) hmap[h] = i; });

  const idxEmail = hmap['Email'];
  if (idxEmail == null) {
    throw new Error('Roster must have an "Email" column.');
  }

  const idxQTC = hmap['QTC'];
  const idxDIR = hmap['DIRECTORY'];

  // Build email -> row index map
  const emailToRow = {};
  for (let r = 1; r < vals.length; r++) {
    const em = String(vals[r][idxEmail] || '').trim().toLowerCase();
    if (em && !emailToRow[em]) {
      emailToRow[em] = r;
    }
  }

  const processedLabel = getOrCreateGmailLabel_(SUBS_PROCESSED_LABEL);
  let totalChanges = 0;

  LIST_ALIASES.forEach(list => {
    const colName = list.column;
    const colIdx = hmap[colName];
    if (colIdx == null) {
      // Column not present; skip this list
      return;
    }

    if (list.unsubscribe) {
      const qUnsub = 'to:' + list.unsubscribe + ' -label:"' + SUBS_PROCESSED_LABEL + '"';
      totalChanges += handleListQuery_(
        qUnsub,
        'UNSUB',
        colIdx,
        colName,
        vals,
        emailToRow,
        idxEmail,
        processedLabel,
        idxQTC,
        idxDIR
      );
    }
    if (list.subscribe) {
      const qSub = 'to:' + list.subscribe + ' -label:"' + SUBS_PROCESSED_LABEL + '"';
      totalChanges += handleListQuery_(
        qSub,
        'SUB',
        colIdx,
        colName,
        vals,
        emailToRow,
        idxEmail,
        processedLabel,
        idxQTC,
        idxDIR
      );
    }
  });

  if (totalChanges > 0) {
    rosterSh.getRange(1, 1, vals.length, vals[0].length).setValues(vals);
  }

  return totalChanges;
}

/**
 * Handle one alias query for a specific list column.
 *
 * mode: 'SUB' or 'UNSUB'
 * colIdx: index of the roster column to toggle
 * listName: e.g. 'QTC' or 'DIRECTORY'
 *
 * Special rule:
 *   - If UNSUB from QTC, also clear DIRECTORY (if present).
 */
function handleListQuery_(query, mode, colIdx, listName,
                          vals, emailToRow, idxEmail, processedLabel,
                          idxQTC, idxDIR) {
  const threads = GmailApp.search(query);
  let changes = 0;

  threads.forEach(thread => {
    const msgs = thread.getMessages();
    if (!msgs || !msgs.length) {
      processedLabel.addToThread(thread);
      return;
    }

    const msg = msgs[msgs.length - 1]; // most recent
    const fromEmail = extractEmail_(msg.getFrom());
    if (!fromEmail) {
      processedLabel.addToThread(thread);
      return;
    }

    const lower = fromEmail.toLowerCase();
    let rowIdx = emailToRow[lower];
    const affectedLists = [];

    if (mode === 'UNSUB') {
      if (rowIdx != null) {
        // Always clear the targeted list
        vals[rowIdx][colIdx] = false;
        affectedLists.push(listName);

        // If unsubscribing from QTC, also clear DIRECTORY
        if (listName === 'QTC' && idxDIR != null) {
          vals[rowIdx][idxDIR] = false;
          affectedLists.push('DIRECTORY');
        }

        changes++;
        if (affectedLists.length) {
          sendConfirm_(fromEmail, 'unsubscribe', affectedLists);
        }
      }
      // Unknown sender: no roster change, still mark processed.
    } else if (mode === 'SUB') {
      if (rowIdx == null) {
        // New subscriber: append row with Email + this list column = TRUE
        rowIdx = vals.length;
        const newRow = new Array(vals[0].length).fill('');
        newRow[idxEmail] = lower;
        newRow[colIdx] = true;
        vals.push(newRow);
        emailToRow[lower] = rowIdx;
      } else {
        // Existing record: set this list column TRUE
        vals[rowIdx][colIdx] = true;
      }
      changes++;
      affectedLists.push(listName);
      sendConfirm_(fromEmail, 'subscribe', affectedLists);
    }

    processedLabel.addToThread(thread);
  });

  return changes;
}

/** Get or create a Gmail label by name */
function getOrCreateGmailLabel_(name) {
  let lbl = GmailApp.getUserLabelByName(name);
  if (!lbl) {
    lbl = GmailApp.createLabel(name);
  }
  return lbl;
}

/** Extract plain email from "Name <addr@example.com>" or raw string */
function extractEmail_(from) {
  if (!from) return '';
  const m = from.match(/<([^>]+)>/);
  const raw = m ? m[1] : from;
  return String(raw).trim().toLowerCase();
}

/**
 * Send confirmation to the requester, CC clerk/news for audit.
 * action: 'subscribe' | 'unsubscribe'
 * lists: array of list names affected
 */
function sendConfirm_(to, action, lists) {
  if (!to) return;

  const niceLists = (lists && lists.length)
    ? lists.join(', ')
    : 'our mailing list';

  const verb = (action === 'unsubscribe') ? 'unsubscribe' : 'subscribe';
  const subj = 'Midlothian Friends Meeting: your ' + verb + ' request';

  const line =
    (action === 'unsubscribe')
      ? 'stop receiving messages'
      : 'receive messages';

  const body =
    'Dear Friend,\n\n' +
    'This is to confirm we have processed your request to ' + line +
    ' for: ' + niceLists + '.\n\n' +
    'If this does not look right, please reply to this message.\n\n' +
    'In Friendship,\n' +
    'Midlothian Friends Meeting\n';

  const opts = {
    cc: CLERK_EMAIL,
    name: 'Midlothian Friends Meeting'
  };

  GmailApp.sendEmail(to, subj, body, opts);
}
