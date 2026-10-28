/**
 * 卡夫卡好物連結頁｜點擊統計
 * 連結頁每次有人點商品，就送 ?no=編號 過來：
 *  - 找到這個編號的那一列（商品清單或表單回應），「點擊數」加 1（沒有這欄會自動加在最右邊）
 *  - 「點擊紀錄」分頁記一筆時間＋編號，之後可以看哪幾天點得多
 */
function doPost(e) { return handle(e); }
function doGet(e) { return handle(e); }

function handle(e) {
  var no = String((e && e.parameter && e.parameter.no) || '').replace(/\D/g, '');
  if (!no) return out('missing');
  var lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    var ss = SpreadsheetApp.getActive();
    // 商品清單、表單回應都可能有這個編號：每個分頁都找，表單回應（有「時間戳記」欄）優先，跟連結頁的顯示一致
    var sheets = ss.getSheets().filter(function (s) { return s.getName() !== '點擊紀錄'; });
    var hit = null;
    sheets.forEach(function (sh) {
      if (hit && hit.form) return;
      if (sh.getLastRow() < 2 || sh.getLastColumn() < 1) return;
      var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
      var isForm = head[0] === '時間戳記';
      var noCol = -1;
      for (var h = 0; h < head.length; h++) { if (head[h].indexOf('編號') === 0) { noCol = h + 1; break; } }
      if (noCol < 0) return;
      var nos = sh.getRange(2, noCol, sh.getLastRow() - 1, 1).getValues();
      for (var i = nos.length - 1; i >= 0; i--) {
        if (String(nos[i][0]).replace(/\D/g, '') === no) { hit = { sh: sh, head: head, row: i + 2, form: isForm }; break; }
      }
    });
    if (hit) {
      var col = hit.head.indexOf('點擊數') + 1;
      if (!col) { col = hit.head.length + 1; hit.sh.getRange(1, col).setValue('點擊數'); }
      var cell = hit.sh.getRange(hit.row, col);
      cell.setValue((Number(cell.getValue()) || 0) + 1);
    }
    var log = ss.getSheetByName('點擊紀錄');
    if (!log) { log = ss.insertSheet('點擊紀錄', ss.getNumSheets()); log.appendRow(['時間', '編號']); }
    log.appendRow([new Date(), no]);
  } finally {
    lock.releaseLock();
  }
  return out('ok');
}

function out(s) { return ContentService.createTextOutput(s); }
