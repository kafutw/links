/**
 * 卡夫卡好物連結頁｜點擊統計
 * 連結頁每次有人點商品，就送 ?no=編號 過來：
 *  - 找到這個編號的那一列（商品清單或表單回應），「點擊數」加 1（沒有這欄會自動加在最右邊）
 *  - 「點擊紀錄」分頁記一筆時間＋編號，之後可以看哪幾天點得多
 * 上架頁（admin.html）送 JSON {action:'add', key, ...} 過來：新增一筆到「上架清單」
 */
var ADMIN_KEY = '__ADMIN_KEY__';  // 上架頁的通行碼（真正的值只放在 Apps Script，不進 repo）
var LIST_SHEET = '上架清單';
var COVER_FOLDER = '好物連結頁封面';
var LIST_HEAD = ['時間戳記', '編號', '商品名', '購買連結', '折扣碼', 'YT 影片網址', '狀態', '團購截止日', '封面圖', '類型', '展開內容'];

function doPost(e) {
  var body = null;
  try { body = JSON.parse((e && e.postData && e.postData.contents) || ''); } catch (err) {}
  if (body && body.action === 'add') return addItem(body);
  return handle(e);
}
function doGet(e) { if (e && e.parameter && e.parameter.meta) return fetchMeta(e.parameter); return handle(e); }

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

/* 上架頁送來的一筆商品：檢查通行碼 → 封面存到雲端硬碟（公開檢視）→ 寫進「上架清單」分頁 */
function addItem(b) {
  if (String(b.key || '') !== ADMIN_KEY) return json({ ok: false, error: '通行碼不對' });
  var no = String(b.no || '').replace(/\D/g, '');
  var name = String(b.name || '').trim();
  if (!no || !name) return json({ ok: false, error: '編號跟商品名一定要填' });
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var cover = '';
    if (b.image && b.image.data) {
      var it = DriveApp.getFoldersByName(COVER_FOLDER);
      var folder = it.hasNext() ? it.next() : DriveApp.createFolder(COVER_FOLDER);
      var blob = Utilities.newBlob(Utilities.base64Decode(b.image.data), b.image.type || 'image/jpeg', 'No' + no + '_' + Date.now() + '.jpg');
      var file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      cover = 'https://drive.google.com/file/d/' + file.getId() + '/view';
    } else if (b.keepCover) {
      cover = String(b.keepCover);
    }
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(LIST_SHEET);
    if (!sh) { sh = ss.insertSheet(LIST_SHEET); sh.appendRow(LIST_HEAD); sh.setFrozenRows(1); }
    // 依標題名稱寫欄（中間有「點擊數」之類別的欄也不會錯位）；缺的標題自動補在最右邊
    var head = ensureHead_(sh);
    var rec = { '時間戳記': new Date(), '編號': no, '商品名': name, '購買連結': String(b.url || ''), '折扣碼': String(b.code || ''),
      'YT 影片網址': String(b.video || ''), '狀態': String(b.status || '隨時可買'), '團購截止日': String(b.until || ''), '封面圖': cover,
      '類型': String(b.type || '商品'), '展開內容': String(b.more || '') };
    sh.appendRow(head.map(function (h) { return rec.hasOwnProperty(h) ? rec[h] : ''; }));
    return json({ ok: true, no: no, cover: cover });
  } finally {
    lock.releaseLock();
  }
}

function ensureHead_(sh) {
  var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  LIST_HEAD.forEach(function (h) {
    if (head.indexOf(h) < 0) { sh.getRange(1, head.length + 1).setValue(h); head.push(h); }
  });
  return head;
}

function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* 只用來叫出雲端硬碟授權視窗：在編輯器選這個函式按「執行」一次就好 */
function authorizeDrive() { Logger.log(DriveApp.getRootFolder().getName()); }

/* 上架頁「貼連結自動填」：代讀商品頁的標題、簡介、主圖（網頁本身不能跨站讀）；嘖嘖另外抓募資截止日 */
function fetchMeta(p) {
  if (String(p.key || '') !== ADMIN_KEY) return json({ ok: false, error: '通行碼不對' });
  var u = String(p.u || '');
  if (!/^https?:\/\//i.test(u)) return json({ ok: false, error: '網址不對' });
  try {
    var res = UrlFetchApp.fetch(u, { muteHttpExceptions: true, followRedirects: true,
      headers: { 'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36', 'Accept-Language': 'zh-TW,zh;q=0.9' } });
    if (res.getResponseCode() >= 400) return json({ ok: false, error: '讀不到（' + res.getResponseCode() + '）' });
    var html = res.getContentText();
    var meta = function (k) {
      var m = new RegExp('<meta[^>]+(?:property|name)=["\']' + k + '["\'][^>]*content=["\']([^"\']*)', 'i').exec(html)
        || new RegExp('<meta[^>]+content=["\']([^"\']*)["\'][^>]*(?:property|name)=["\']' + k + '["\']', 'i').exec(html);
      return m ? decode_(m[1]) : '';
    };
    var t = /<title[^>]*>([^<]*)/i.exec(html);
    var until = /募資期間[\s\S]{0,120}?[–~～-]\s*(\d{4}\/\d{1,2}\/\d{1,2})/.exec(html);
    return json({ ok: true, title: meta('og:title') || (t ? decode_(t[1]) : ''), desc: meta('og:description') || meta('description'),
      image: meta('og:image'), until: until ? until[1] : '' });
  } catch (err) {
    return json({ ok: false, error: String(err.message || err) });
  }
}
function decode_(s) {
  return String(s).replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&#39;|&#x27;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').trim();
}
