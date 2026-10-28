/**
 * 卡夫卡好物連結頁｜點擊統計
 * 連結頁每次有人點商品，就送 ?no=編號 過來：
 *  - 第一個工作表（商品清單）那一列的「點擊數」加 1（沒有這欄會自動加在最右邊）
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
    var sh = ss.getSheets()[0];
    var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    var col = head.indexOf('點擊數') + 1;
    if (!col) { col = head.length + 1; sh.getRange(1, col).setValue('點擊數'); }
    var noCol = 1;
    for (var h = 0; h < head.length; h++) { if (head[h].indexOf('編號') === 0) { noCol = h + 1; break; } }
    var last = sh.getLastRow();
    if (last >= 2) {
      var nos = sh.getRange(2, noCol, last - 1, 1).getValues();
      for (var i = 0; i < nos.length; i++) {
        if (String(nos[i][0]).replace(/\D/g, '') === no) {
          var cell = sh.getRange(i + 2, col);
          cell.setValue((Number(cell.getValue()) || 0) + 1);
          break;
        }
      }
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
