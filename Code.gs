/**
 * KPI Data Entry - Google Sheet backend
 *
 * HOW TO DEPLOY
 * 1. Open your Google Sheet -> Extensions -> Apps Script
 * 2. Delete any existing code, paste this file, Ctrl+S
 * 3. Deploy -> New deployment -> Type: Web app
 *      Execute as: Me
 *      Who has access: Anyone
 *    -> Deploy -> copy the /exec URL
 * 4. Paste that URL into kpi-entry.html -> Settings -> Save URL -> Test connection
 *
 * The script creates two tabs automatically:
 *   KPI Data   -> month-year data entries (upsert by Month-Year + SBU + KPI Name)
 *   KPI Master -> KPI definitions (upsert by SBU + KPI Name)
 */

var DATA_SHEET = 'KPI Data';
var MASTER_SHEET = 'KPI Master';
var DATA_HEADERS = ['Month-Year','SBU','Department','Section','KPI Name','Measure KPI','Baseline','SRF','KPI Target','Direction','Basis','Weight %','Actual','KPI Achievement','KPI Status','Supporting File','Root Cause','Action Plan','Remarks','Submitted At','Updated At'];
var MASTER_HEADERS = ['SBU','Department','Section','KPI Name','Measure KPI','Baseline','SRF','KPI Target','Direction','Basis','Weight %','Calc','Updated At'];

function doPost(e) {
  try {
    var p = JSON.parse(e.postData.contents);
    if (p.action === 'save')       return json_({ ok: true, saved: saveRows_(p.rows || []) });
    if (p.action === 'saveMaster') return json_({ ok: true, saved: saveMaster_(p.kpis || []) });
    if (p.action === 'clearMonth') return json_({ ok: true, cleared: clearMonth_(p.month, p.sbu) });
    return json_({ ok: false, error: 'Unknown action: ' + p.action });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  try {
    var a = (e && e.parameter && e.parameter.action) || 'list';
    if (a === 'ping') return json_({ ok: true, msg: 'connected', spreadsheet: SpreadsheetApp.getActiveSpreadsheet().getName() });
    if (a === 'list') return json_({ ok: true, rows: readAll_(DATA_SHEET, DATA_HEADERS), master: readAll_(MASTER_SHEET, MASTER_HEADERS) });
    return json_({ ok: false, error: 'Unknown action: ' + a });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function sheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setBackground('#1d4ed8')
      .setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

function key_(row) { return row[0] + '||' + row[1] + '||' + row[4]; }

function saveRows_(rows) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = sheet_(ss, DATA_SHEET, DATA_HEADERS);
  var index = {};
  var last = sh.getLastRow();
  if (last > 1) {
    var old = sh.getRange(2, 1, last - 1, DATA_HEADERS.length).getValues();
    for (var i = 0; i < old.length; i++) {
      if (old[i][0]) index[key_(old[i])] = i + 2;
    }
  }
  var now = new Date();
  for (var r = 0; r < rows.length; r++) {
    var row = normalize_(rows[r], DATA_HEADERS);
    row[row.length - 1] = now;
    var k = key_(row);
    if (index[k]) {
      sh.getRange(index[k], 1, 1, DATA_HEADERS.length).setValues([row]);
    } else {
      sh.appendRow(row);
    }
  }
  return rows.length;
}

function saveMaster_(kpis) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = sheet_(ss, MASTER_SHEET, MASTER_HEADERS);
  var index = {};
  var last = sh.getLastRow();
  if (last > 1) {
    var old = sh.getRange(2, 1, last - 1, MASTER_HEADERS.length).getValues();
    for (var i = 0; i < old.length; i++) {
      if (old[i][0]) index[old[i][0] + '||' + old[i][3]] = i + 2;
    }
  }
  var now = new Date();
  for (var r = 0; r < kpis.length; r++) {
    var k = kpis[r];
    var row = [k.sbu, k.dept, k.sec, k.name, k.measure, k.baseline, k.srf, k.target,
               k.dir, k.basis, k.weight, k.calc || '', now];
    var key = row[0] + '||' + row[3];
    if (index[key]) {
      sh.getRange(index[key], 1, 1, MASTER_HEADERS.length).setValues([row]);
    } else {
      sh.appendRow(row);
    }
  }
  return kpis.length;
}

function clearMonth_(month, sbu) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(DATA_SHEET);
  if (!sh || sh.getLastRow() < 2) return 0;
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  var n = 0;
  for (var i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][0]) === String(month) && String(vals[i][1]) === String(sbu)) {
      sh.deleteRow(i + 2);
      n++;
    }
  }
  return n;
}

function normalize_(obj, headers) {
  var out = [];
  for (var i = 0; i < headers.length; i++) {
    var v = obj[headers[i]];
    out.push(v === undefined || v === null ? '' : v);
  }
  return out;
}

function readAll_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) return [];
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, headers.length).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var o = {};
    for (var j = 0; j < headers.length; j++) o[headers[j]] = vals[i][j];
    out.push(o);
  }
  return out;
}
