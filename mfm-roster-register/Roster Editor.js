/**
 * MFM Roster Editor Web App (v2.2)
 */

const SHEET_NAME = 'Midlothian Friends Meeting Roster';

// Only these headers are treated as boolean columns
const BOOL_HEADERS = ['Directory', 'QTC', 'Survey'];

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var user = Session.getActiveUser().getEmail();
  var editors = ss.getEditors().map(function (ed) { return ed.getEmail(); });
  var owner = ss.getOwner() ? ss.getOwner().getEmail() : null;

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

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getRosterData() {
  var sh = getSheet_();

  // Ensure Survey exists and is positioned right after QTC (or after Directory)
  ensureSurveyColumn_(sh);

  var data = sh.getDataRange().getValues();
  if (data.length < 2) return { headers: [], rows: [] };

  var headers = data[0];
  var outRows = [];

  for (var r = 1; r < data.length; r++) {
    var rowValues = data[r];
    var hasContent = rowValues.some(function (v) { return v !== '' && v !== null; });
    if (!hasContent) continue;
    outRows.push({ rowNum: r + 1, values: rowValues });
  }

  return { headers: headers, rows: outRows };
}

function saveRosterRows(changes) {
  if (!Array.isArray(changes) || changes.length === 0) return { updated: 0 };

  var sh = getSheet_();
  ensureSurveyColumn_(sh);

  var maxCols = sh.getLastColumn();
  var headers = sh.getRange(1, 1, 1, maxCols).getValues()[0];
  var idx = buildIndex_(headers);
  var boolCols = getBooleanColSet_(headers);

  // Validate incoming non-blank emails
  var invalids = [];
  changes.forEach(function (chg) {
    if (!chg || !Array.isArray(chg.values)) return;
    var email = getCell_(chg.values, idx.email);
    if (email && String(email).trim()) {
      if (!isValidEmail_(String(email).trim())) {
        invalids.push({ rowNum: chg.rowNum || null, email: String(email) });
      }
    }
  });
  if (invalids.length) {
    throw new Error('Invalid email(s): ' + invalids.map(function (x) {
      return (x.rowNum ? ('row ' + x.rowNum + ': ') : '') + x.email;
    }).join(' | '));
  }

  var updated = 0;

  changes.forEach(function (change) {
    if (!change || !Array.isArray(change.values)) return;

    var rowNum = change.rowNum;
    var vals = change.values.slice();

    while (vals.length < maxCols) vals.push('');
    if (vals.length > maxCols) vals = vals.slice(0, maxCols);

    vals = normalizeRow_(vals, idx, boolCols);

    if (!rowNum || isNaN(rowNum)) {
      sh.appendRow(vals);
      updated++;
    } else {
      sh.getRange(rowNum, 1, 1, maxCols).setValues([vals]);
      updated++;
    }
  });

  // Normalize entire sheet + sort after save
  normalizeWholeSheet_(sh);
  sortSheet_(sh);

  return { updated: updated };
}

/* ---------------- helpers ---------------- */

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('Sheet "' + SHEET_NAME + '" not found.');
  return sh;
}

function normHeader_(h) {
  return String(h || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function ensureSurveyColumn_(sh) {
  var lastCol = sh.getLastColumn();
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];

  function findCol(name) {
    var target = normHeader_(name);
    for (var i = 0; i < headers.length; i++) {
      if (normHeader_(headers[i]) === target) return i + 1; // 1-based
    }
    return -1;
  }

  var surveyCol = findCol('Survey');
  if (surveyCol !== -1) return;

  var qtcCol = findCol('QTC');
  var dirCol = findCol('Directory');

  var insertAfter = qtcCol !== -1 ? qtcCol : (dirCol !== -1 ? dirCol : lastCol);

  sh.insertColumnAfter(insertAfter);
  sh.getRange(1, insertAfter + 1).setValue('Survey');
}

function buildIndex_(headers) {
  var map = {};
  headers.forEach(function (h, i) { map[normHeader_(h)] = i; });

  function pick() {
    for (var i = 0; i < arguments.length; i++) {
      var key = arguments[i];
      if (key in map) return map[key];
    }
    return -1;
  }

  return {
    first: pick('first name', 'firstname', 'given name', 'given'),
    last: pick('last name', 'lastname', 'surname', 'family name', 'family'),
    email: pick('email', 'email address', 'e-mail', 'e-mail address'),
    phone: pick('phone', 'phone number', 'mobile', 'cell', 'cell phone'),
    address: pick('address', 'mailing address', 'street address')
  };
}

function getCell_(rowValues, colIndex) {
  if (colIndex == null || colIndex < 0) return '';
  return rowValues[colIndex];
}

function getBooleanColSet_(headers) {
  var boolSet = {};
  var want = BOOL_HEADERS.map(normHeader_);
  headers.forEach(function (h, i) {
    if (want.indexOf(normHeader_(h)) !== -1) boolSet[i] = true;
  });
  return boolSet; // keys are 0-based indexes
}

function coerceBoolean_(v) {
  if (v === true || v === false) return v;
  if (v == null) return false;
  if (typeof v === 'number') return v !== 0;
  var s = String(v).trim().toLowerCase();
  if (['true', 'yes', 'y', '1', 'on', 'checked'].indexOf(s) !== -1) return true;
  if (['false', 'no', 'n', '0', 'off', ''].indexOf(s) !== -1) return false;
  return false;
}

function isTrueFalseString_(v) {
  if (typeof v !== 'string') return false;
  var s = v.trim().toUpperCase();
  return s === 'TRUE' || s === 'FALSE';
}

function normalizePhone_(v) {
  var s = String(v == null ? '' : v).trim();
  if (!s) return '';
  var digits = s.replace(/\D/g, '');
  if (digits.length === 11 && digits.charAt(0) === '1') digits = digits.slice(1);
  if (digits.length !== 10) return s;
  return '(' + digits.slice(0, 3) + ') ' + digits.slice(3, 6) + '-' + digits.slice(6);
}

function isValidEmail_(email) {
  var s = String(email || '').trim();
  if (!s) return true;
  if (s.length > 254) return false;
  if (/\s/.test(s)) return false;
  return /^[^@]+@[^@]+\.[^@]{2,}$/.test(s);
}

// Email de-dupe: preserve user tags; ignore only trailing +digits
function parseEmail_(email) {
  var s = String(email || '').trim();
  var m = s.match(/^([^@]+)@([^@]+)$/);
  if (!m) return null;
  return { local: m[1], domain: m[2] };
}
function stripTrailingPlusNLocal_(local) {
  return String(local || '').replace(/\+(\d+)$/, '');
}
function dupKey_(email) {
  var p = parseEmail_(email);
  if (!p) return null;
  return stripTrailingPlusNLocal_(p.local).toLowerCase() + '@' + p.domain.toLowerCase();
}
function dedupeEmailValue_(email, n) {
  var p = parseEmail_(email);
  if (!p) return String(email || '');
  var baseLocal = stripTrailingPlusNLocal_(p.local);
  return n > 0 ? (baseLocal + '+' + n + '@' + p.domain) : (baseLocal + '@' + p.domain);
}

function normalizeRow_(vals, idx, boolCols) {
  // booleans only in boolean columns; TRUE/FALSE removed from others
  for (var c = 0; c < vals.length; c++) {
    var v = vals[c];
    if (boolCols[c]) {
      vals[c] = coerceBoolean_(v); // interprets TRUE/FALSE, etc.
    } else {
      if (v === true || v === false || isTrueFalseString_(v)) vals[c] = '';
    }
  }

  // phone formatting
  if (idx.phone >= 0) vals[idx.phone] = normalizePhone_(vals[idx.phone]);

  return vals;
}

function normalizeWholeSheet_(sh) {
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2) return;

  var range = sh.getRange(1, 1, lastRow, lastCol);
  var data = range.getValues();
  var headers = data[0];
  var idx = buildIndex_(headers);
  var boolCols = getBooleanColSet_(headers);

  // row-wise cleanup
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var hasContent = row.some(function (v) { return v !== '' && v !== null; });
    if (!hasContent) continue;
    data[r] = normalizeRow_(row, idx, boolCols);
  }

  // email de-dupe across sheet
  if (idx.email >= 0) {
    var seen = new Map();
    for (var rr = 1; rr < data.length; rr++) {
      var email = data[rr][idx.email];
      if (!email || !String(email).trim()) continue;

      var s = String(email).trim();
      if (!isValidEmail_(s)) continue;

      var key = dupKey_(s);
      if (!key) continue;

      var count = seen.get(key) || 0;
      data[rr][idx.email] = dedupeEmailValue_(s, count);
      seen.set(key, count + 1);
    }
  }

  range.setValues(data);
}

function sortSheet_(sh) {
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 3) return;

  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idx = buildIndex_(headers);

  if (idx.first < 0 || idx.last < 0) return;

  sh.getRange(2, 1, lastRow - 1, lastCol).sort([
    { column: idx.first + 1, ascending: true },
    { column: idx.last + 1, ascending: true }
  ]);
}