/**
 * MFM Roster Editor Web App (v2)
 *
 * Bound to: "Midlothian Friends Meeting Roster" spreadsheet
 * Tab name: Midlothian Friends Meeting Roster
 *
 * Access control:
 * - Only users who are owner or editors of this spreadsheet
 *   can use the web app; others see "Access denied".
 */

/** Serve the HTML UI */
function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var user = Session.getActiveUser().getEmail();
  var editors = ss.getEditors().map(function(ed){ return ed.getEmail(); });
  var owner = ss.getOwner() ? ss.getOwner().getEmail() : null;

  // Allow if user is owner or editor; otherwise block.
  if (!user || (user !== owner && editors.indexOf(user) === -1)) {
    return HtmlService
      .createHtmlOutput('<h2>Access denied</h2><p>Your account is not authorized to edit this roster.</p>')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('MFM Roster Editor')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Include helper (if you later add partials) */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Get full roster data.
 * Returns: { headers: [...], rows: [ { rowNum: n, values: [...] }, ... ] }
 */
function getRosterData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Midlothian Friends Meeting Roster');
  if (!sh) throw new Error('Sheet "Midlothian Friends Meeting Roster" not found.');

  var data = sh.getDataRange().getValues();
  if (data.length < 2) {
    return { headers: [], rows: [] };
  }

  var headers = data[0];
  var outRows = [];

  for (var r = 1; r < data.length; r++) {
    var rowValues = data[r];
    // Skip completely blank rows at the end
    var hasContent = rowValues.some(function(v){ return v !== '' && v !== null; });
    if (!hasContent) continue;
    outRows.push({
      rowNum: r + 1,       // 1-based row index in sheet
      values: rowValues
    });
  }

  return {
    headers: headers,
    rows: outRows
  };
}

/**
 * Save updated rows.
 * Input: [{ rowNum: Number|null, values: [..] }, ...]
 * - If rowNum is a number: update that existing row.
 * - If rowNum is null/0: append as a new row.
 */
function saveRosterRows(changes) {
  if (!Array.isArray(changes) || changes.length === 0) {
    return { updated: 0 };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Midlothian Friends Meeting Roster');
  if (!sh) throw new Error('Sheet "Midlothian Friends Meeting Roster" not found.');

  var maxCols = sh.getLastColumn();
  var updated = 0;

  changes.forEach(function(change) {
    var rowNum = change.rowNum;
    var vals = change.values;
    if (!Array.isArray(vals)) return;

    // Normalize booleans: treat true-like/false-like strings as booleans
    var norm = vals.map(function(v) {
      if (v === true || v === false) return v;
      if (typeof v === 'string') {
        var s = v.trim().toLowerCase();
        if (['true','yes','y','1','on','checked'].indexOf(s) !== -1) return true;
        if (['false','no','n','0','off',''].indexOf(s) !== -1) return false;
      }
      return v;
    });

    // Pad or trim to sheet width
    if (norm.length < maxCols) {
      while (norm.length < maxCols) norm.push('');
    } else if (norm.length > maxCols) {
      norm = norm.slice(0, maxCols);
    }

    if (!rowNum || isNaN(rowNum)) {
      // Append as new row
      sh.appendRow(norm);
      updated++;
    } else {
      // Update existing row
      sh.getRange(rowNum, 1, 1, maxCols).setValues([norm]);
      updated++;
    }
  });

  return { updated: updated };
}
