/**
 * 道館教練備課助手 — 全館共用後端（Google Apps Script + Google Sheets）
 *
 * 功能：讓所有教練共用同一套「教案素材庫 / 影片庫 / 常用模板」。
 *  - GET  ?action=pull            → 回傳 {materials, videos, templates}
 *  - POST {action:'upsert',coll,item}  → 新增或更新一筆
 *  - POST {action:'delete',coll,id}    → 刪除一筆
 *  - POST {action:'bulk',coll,item:[]} → 整批覆蓋（初始化用）
 *
 * 部署步驟：
 *  1. 到 sheets.new 建立一個 Google 試算表。
 *  2. 上方選單「擴充功能 → Apps Script」，把本檔內容整個貼進去、儲存。
 *  3. 右上「部署 → 新增部署作業 → 類型選『網頁應用程式』」。
 *     - 執行身分：我（你自己）
 *     - 誰可以存取：『任何人』
 *  4. 部署後複製「網頁應用程式」的 /exec 網址。
 *  5. 打開備課助手 App → ⚙️ → 貼到「全館雲端同步網址」→ 儲存。
 *  6. 在「資料最齊全」的那台按一次「⬆️ 上傳本機到雲端（初始化）」，
 *     之後其他教練貼同一個網址、按「⬇️ 立即下載雲端」即可共用。
 */

var COLLS = ['materials', 'videos', 'templates'];

function getSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(['id', 'json']); }
  if (sh.getLastRow() === 0) sh.appendRow(['id', 'json']);
  return sh;
}

function readColl_(name) {
  var sh = getSheet_(name);
  var v = sh.getDataRange().getValues();
  var arr = [];
  for (var i = 1; i < v.length; i++) {
    if (v[i][1]) { try { arr.push(JSON.parse(v[i][1])); } catch (e) {} }
  }
  return arr;
}

function upsert_(sh, item) {
  var v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][0]) === String(item.id)) {
      sh.getRange(i + 1, 2).setValue(JSON.stringify(item));
      return;
    }
  }
  sh.appendRow([item.id, JSON.stringify(item)]);
}

function del_(sh, id) {
  var v = sh.getDataRange().getValues();
  for (var i = v.length - 1; i >= 1; i--) {
    if (String(v[i][0]) === String(id)) sh.deleteRow(i + 1);
  }
}

function clearColl_(sh) {
  var last = sh.getLastRow();
  if (last > 1) sh.deleteRows(2, last - 1);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var out = {};
  COLLS.forEach(function (c) { out[c] = readColl_(c); });
  out.ok = true;
  return json_(out);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var body = JSON.parse(e.postData.contents);
    var action = body.action, coll = body.coll;
    if (COLLS.indexOf(coll) < 0) return json_({ error: 'bad collection' });
    var sh = getSheet_(coll);
    if (action === 'upsert') { upsert_(sh, body.item); return json_({ ok: true }); }
    if (action === 'delete') { del_(sh, body.id); return json_({ ok: true }); }
    if (action === 'bulk') {
      clearColl_(sh);
      (body.item || []).forEach(function (it) { sh.appendRow([it.id, JSON.stringify(it)]); });
      return json_({ ok: true, count: (body.item || []).length });
    }
    return json_({ error: 'bad action' });
  } catch (err) {
    return json_({ error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}
