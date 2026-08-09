/**
 * 座位圖 — Google 試算表後端
 * 用法：在試算表中「擴充功能 → Apps Script」，把整段貼上，
 *      先執行一次 setup，再「部署 → 新增部署作業 → 網頁應用程式」。
 * 詳細步驟見「設定說明.md」。
 */

// ── 可自行修改的設定 ─────────────────────────────
var SEAT_COUNT = 39;          // 座位數
var TOKEN      = '';          // 通行碼：留空＝不檢查。想更安全就填一組字（例如 'office2026'），
                              // 並在網頁的「連線設定」裡輸入同一組。
// ────────────────────────────────────────────────

var SEAT_SHEET   = '座位資料';
var LEGEND_SHEET = '顏色說明';
var HEADERS      = ['座位', '姓名/標籤', '色碼', '備註', '更新時間', '更新者'];


/** 第一次使用請手動執行這個函式：建立工作表並授權 */
function setup() {
  seatSheet_();
  legendSheet_();
  SpreadsheetApp.getActiveSpreadsheet().toast('座位圖工作表已就緒', '完成', 5);
}


/** 瀏覽器直接打開網址時，回傳目前資料（方便你確認部署成功） */
function doGet(e) {
  try {
    var p = (e && e.parameter) || {};
    checkToken_(p.token);
    return json_({ ok: true, data: readAll_() });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}


/** 網頁的讀寫都走這裡 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    checkToken_(req.token);

    switch (req.action) {
      case 'read':                                   break;
      case 'saveSeat':  saveSeat_(req);              break;
      case 'saveMany':  (req.seats || []).forEach(saveSeat_); break;
      case 'legend':    saveLegend_(req.hex, req.text); break;
      case 'clearAll':  clearAll_();                 break;
      default: throw new Error('未知的動作：' + req.action);
    }
    return json_({ ok: true, data: readAll_() });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}


// ── 讀取 ────────────────────────────────────────

function readAll_() {
  var sh = seatSheet_();
  var vals = sh.getRange(2, 1, SEAT_COUNT, HEADERS.length).getValues();
  var seats = {};
  for (var i = 0; i < vals.length; i++) {
    var r = vals[i];
    var no = Number(r[0]);
    if (!no) continue;
    var label = String(r[1] || '').trim();
    var color = String(r[2] || '').trim();
    var note  = String(r[3] || '');
    if (!label && !color && !note) continue;
    seats[no] = {
      label: label,
      color: color,
      note:  note,
      at:    r[4] instanceof Date ? r[4].toISOString() : String(r[4] || ''),
      by:    String(r[5] || '')
    };
  }

  var lg = legendSheet_();
  var legend = {};
  var last = lg.getLastRow();
  if (last > 1) {
    lg.getRange(2, 1, last - 1, 2).getValues().forEach(function (row) {
      var hex = String(row[0] || '').trim();
      if (hex) legend[hex] = String(row[1] || '');
    });
  }

  return { seats: seats, legend: legend, rev: new Date().toISOString() };
}


// ── 寫入 ────────────────────────────────────────

function saveSeat_(o) {
  var no = Number(o.seat);
  if (!(no >= 1 && no <= SEAT_COUNT)) throw new Error('座位號碼不正確：' + o.seat);
  var sh = seatSheet_();
  var row = no + 1;
  var color = String(o.color || '').trim();
  sh.getRange(row, 2, 1, 5).setValues([[
    String(o.label || ''),
    color,
    String(o.note || ''),
    new Date(),
    String(o.by || '')
  ]]);
  // 讓試算表本身也看得出顏色
  sh.getRange(row, 3).setBackground(color || null);
}

function saveLegend_(hex, text) {
  hex = String(hex || '').trim();
  if (!hex) return;
  var sh = legendSheet_();
  var last = sh.getLastRow();
  if (last > 1) {
    var col = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < col.length; i++) {
      if (String(col[i][0] || '').trim() === hex) {
        sh.getRange(i + 2, 2).setValue(String(text || ''));
        return;
      }
    }
  }
  sh.appendRow([hex, String(text || '')]);
  sh.getRange(sh.getLastRow(), 1).setBackground(hex);
}

function clearAll_() {
  var sh = seatSheet_();
  var blank = [];
  for (var i = 0; i < SEAT_COUNT; i++) blank.push(['', '', '', '', '']);
  sh.getRange(2, 2, SEAT_COUNT, 5).setValues(blank);
  sh.getRange(2, 3, SEAT_COUNT, 1).setBackground(null);
}


// ── 工作表 ──────────────────────────────────────

function seatSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SEAT_SHEET) || ss.insertSheet(SEAT_SHEET);

  if (String(sh.getRange(1, 1).getValue()) !== HEADERS[0]) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidth(2, 140);
    sh.setColumnWidth(4, 320);
    sh.setColumnWidth(5, 160);
  }
  // 補滿 1..SEAT_COUNT 的列
  var need = SEAT_COUNT + 1 - sh.getMaxRows();
  if (need > 0) sh.insertRowsAfter(sh.getMaxRows(), need);
  var nums = sh.getRange(2, 1, SEAT_COUNT, 1).getValues();
  var fix = [], changed = false;
  for (var i = 0; i < SEAT_COUNT; i++) {
    fix.push([i + 1]);
    if (Number(nums[i][0]) !== i + 1) changed = true;
  }
  if (changed) sh.getRange(2, 1, SEAT_COUNT, 1).setValues(fix);
  return sh;
}

function legendSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(LEGEND_SHEET) || ss.insertSheet(LEGEND_SHEET);
  if (String(sh.getRange(1, 1).getValue()) !== '色碼') {
    sh.getRange(1, 1, 1, 2).setValues([['色碼', '說明']]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidth(2, 220);
  }
  return sh;
}


// ── 雜項 ────────────────────────────────────────

function checkToken_(t) {
  if (TOKEN && String(t || '') !== TOKEN) throw new Error('通行碼不正確');
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
