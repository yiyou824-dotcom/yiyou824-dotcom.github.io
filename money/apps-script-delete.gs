/**
 * 記帳本：刪除功能（貼進 Apps Script）
 *
 * 1. 把下面整段 handleDelete_ 函式貼到 Apps Script 的最下面。
 * 2. 在原本的 doGet(e) 函式「第一行」加上這行：
 *
 *      if (e && e.parameter && e.parameter.action === 'delete') return handleDelete_(e);
 *
 * 3. 部署 → 管理部署作業 → 編輯（鉛筆）→ 版本選「新版本」→ 部署。
 *    網址不會變，網頁不用改。
 */
function handleDelete_(e) {
  var p = e.parameter;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];   // 記帳資料在第一個工作表
  var data = sheet.getDataRange().getValues();
  var head = data[0].map(function (h) { return String(h).trim(); });
  var col = function (names) {
    for (var i = 0; i < names.length; i++) { var k = head.indexOf(names[i]); if (k > -1) return k; }
    return -1;
  };
  var cDate = col(['日期']), cCat = col(['類型', '消費類型']), cAmt = col(['金額']),
      cUser = col(['記帳人']), cNote = col(['備註']);

  var fmtDate = function (v) {
    if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Taipei', 'yyyy-MM-dd');
    var m = String(v).match(/(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
    return m ? m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2) : String(v);
  };
  var same = function (row) {
    return fmtDate(row[cDate]) === p.date &&
           Number(String(row[cAmt]).replace(/[^0-9.\-]/g, '')) === Number(p.amount) &&
           (cCat < 0 || String(row[cCat]).trim() === p.category) &&
           (cUser < 0 || String(row[cUser]).trim() === p.user) &&
           (cNote < 0 || String(row[cNote]).trim() === String(p.note || '').trim());
  };

  var result = { ok: false, error: 'not_found' };
  for (var r = data.length - 1; r >= 1; r--) {          // 從最新的一筆往回找，重複時刪最近那筆
    if (same(data[r])) { sheet.deleteRow(r + 1); result = { ok: true }; break; }
  }

  var json = JSON.stringify(result);
  var cb = p.callback;
  return cb
    ? ContentService.createTextOutput(cb + '(' + json + ')').setMimeType(ContentService.MimeType.JAVASCRIPT)
    : ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
