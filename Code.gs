/**
 * 攝影器材管理與預約系統 — 後端 API
 * 部署方式：Google Apps Script Web App（見 README「部署步驟」）
 *
 * 工作表「器材」欄位（第 1 列為標題）：
 *   A:id B:名稱 C:分類 D:型號 E:序號 F:狀態 G:照片網址 H:備註 I:群組 J:數量
 *   狀態值：可借用 / 維修中 / 報廢
 *   群組：同群組的器材在前端會收合成一張可展開的卡（例如記憶卡分容量）；留空＝獨立一項
 *   數量：庫存數量（同型可借幾份）；獨立器材填 1
 *
 * 工作表「預約」欄位（第 1 列為標題）：
 *   A:id B:器材id C:器材名稱 D:借用人 E:部門 F:聯絡方式(Email)
 *   G:借出日 H:歸還日 I:用途 J:狀態 K:申請時間 L:審核備註 M:數量
 *   狀態值：待審核 / 已核准 / 已拒絕 / 已歸還 / 已取消
 */

// ====== 設定 ======
var ADMIN_PIN = '1234';          // 管理端操作密碼，請改成你自己的
var NOTIFY = true;               // 是否開啟 email 通知
var ADMIN_EMAIL = '';            // 收「新申請通知」的信箱；留空＝寄給指令碼擁有者（你）
var APP_URL = 'https://kevinyee-yee.github.io/gear-booking/';

var SHEET_EQUIP = '器材';
var SHEET_RESV  = '預約';

// ====== 入口 ======
function doGet(e)  { return handle(e); }
function doPost(e) { return handle(e); }

function handle(e) {
  var params = {};
  if (e && e.parameter) params = e.parameter;
  if (e && e.postData && e.postData.contents) {
    try {
      var body = JSON.parse(e.postData.contents);
      for (var k in body) params[k] = body[k];
    } catch (err) {}
  }

  var action = params.action || '';
  var out;
  try {
    switch (action) {
      case 'list':                out = apiList(); break;
      case 'create_reservation':  out = apiCreateReservation(params); break;
      case 'create_reservations': out = apiCreateReservations(params); break; // 購物車批次
      case 'cancel_reservation':  out = apiCancelReservation(params); break;
      case 'review':              out = apiReview(params); break;
      case 'review_batch':        out = apiReviewBatch(params); break;
      case 'return':              out = apiReturn(params); break;
      case 'return_batch':        out = apiReturnBatch(params); break;
      case 'add_equipment':       out = apiAddEquipment(params); break;
      case 'update_equipment':    out = apiUpdateEquipment(params); break;
      case 'delete_equipment':    out = apiDeleteEquipment(params); break;
      default: out = { ok: false, error: '未知的 action: ' + action };
    }
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  return json(out);
}

// ====== 讀取（公開）======
function apiList() {
  return { ok: true, equipment: readSheet(SHEET_EQUIP), reservations: readSheet(SHEET_RESV) };
}

// ====== 建立單筆預約（公開，保留相容）======
function apiCreateReservation(p) {
  requireFields(p, ['器材id', '借用人', '部門', '借出日', '歸還日']);
  var qty = Number(p['數量']) || 1;
  var res = tryReserveOne(p['器材id'], p['借出日'], p['歸還日'], qty, null);
  if (!res.ok) return res;

  var id = writeReservation(res.equip, p, qty, 'B' + new Date().getTime());
  notifyAdminNew([{ 名稱: res.equip['名稱'], 數量: qty }], p);
  return { ok: true, id: id };
}

// ====== 建立多筆預約（購物車批次）======
function apiCreateReservations(p) {
  requireFields(p, ['items', '借用人', '部門', '借出日', '歸還日']);
  var items = p['items'];
  if (typeof items === 'string') { try { items = JSON.parse(items); } catch (e) {} }
  if (!items || !items.length) return { ok: false, error: '清單是空的' };
  if (p['借出日'] >= p['歸還日']) return { ok: false, error: '歸還時間需晚於借出時間' };

  // 先全部驗證，任何一項不足就整批擋下，不建立半套
  var checked = [];
  for (var i = 0; i < items.length; i++) {
    var qty = Number(items[i]['數量']) || 1;
    var res = tryReserveOne(items[i]['器材id'], p['借出日'], p['歸還日'], qty, null);
    if (!res.ok) return { ok: false, error: res.error };
    checked.push({ equip: res.equip, qty: qty });
  }
  // 通過才一次建立（同一次申請共用一個批次碼）
  var batch = 'B' + new Date().getTime();
  var ids = [];
  for (var j = 0; j < checked.length; j++) {
    ids.push(writeReservation(checked[j].equip, p, checked[j].qty, batch));
  }
  notifyAdminNew(checked.map(function (c) { return { 名稱: c.equip['名稱'], 數量: c.qty }; }), p);
  return { ok: true, ids: ids, count: ids.length };
}

// 驗證單項可借（回傳 {ok, equip} 或 {ok:false, error}）
function tryReserveOne(equipId, start, end, qty, excludeId) {
  var equip = findRowById(SHEET_EQUIP, equipId);
  if (!equip.row) return { ok: false, error: '找不到器材（' + equipId + '）' };
  if (equip.data['狀態'] !== '可借用') {
    return { ok: false, error: equip.data['名稱'] + '：目前狀態為「' + equip.data['狀態'] + '」，無法預約' };
  }
  if (start >= end) return { ok: false, error: '歸還時間需晚於借出時間' };
  var a = availability(equipId, start, end, excludeId);
  if (qty > a.free) {
    return { ok: false, error: equip.data['名稱'] + '：該時段僅剩 ' + a.free + ' 份（庫存 ' + a.stock + '），無法借 ' + qty + ' 份' };
  }
  return { ok: true, equip: equip.data };
}

function writeReservation(equipData, p, qty, batch) {
  var sh = sheet(SHEET_RESV);
  var id = 'R' + new Date().getTime() + Math.floor(Math.random() * 1000);
  sh.appendRow([
    id, equipData['id'], equipData['名稱'], p['借用人'], p['部門'],
    p['聯絡方式'] || '', p['借出日'], p['歸還日'], p['用途'] || '',
    '待審核', new Date(), '', qty, batch || ''
  ]);
  return id;
}

// ====== 取消申請（公開）======
function apiCancelReservation(p) {
  requireFields(p, ['id']);
  var r = findRowById(SHEET_RESV, p['id']);
  if (!r.row) return { ok: false, error: '找不到申請' };
  if (['已歸還', '已拒絕', '已取消'].indexOf(r.data['狀態']) >= 0) {
    return { ok: false, error: '此申請已結束，無法取消' };
  }
  setCell(SHEET_RESV, r.row, '狀態', '已取消');
  return { ok: true };
}

// ====== 審核（管理端）======
function apiReview(p) {
  requireAdmin(p);
  requireFields(p, ['id', 'decision']);
  var r = findRowById(SHEET_RESV, p['id']);
  if (!r.row) return { ok: false, error: '找不到申請' };
  if (r.data['狀態'] !== '待審核') return { ok: false, error: '此申請不是待審核狀態' };

  var approved;
  if (p.decision === 'approve') {
    var qty = Number(r.data['數量']) || 1;
    var a = availability(r.data['器材id'], r.data['借出日'], r.data['歸還日'], p['id']);
    if (qty > a.free) {
      return { ok: false, error: '該時段僅剩 ' + a.free + ' 份，無法核准 ' + qty + ' 份' };
    }
    setCell(SHEET_RESV, r.row, '狀態', '已核准');
    approved = true;
  } else {
    setCell(SHEET_RESV, r.row, '狀態', '已拒絕');
    approved = false;
  }
  var note = (p['審核備註'] != null) ? p['審核備註'] : '';
  if (p['審核備註'] != null) setCell(SHEET_RESV, r.row, '審核備註', note);

  notify(r.data['聯絡方式'],
    '【器材預約】' + (approved ? '已核准' : '未通過') + '：' + r.data['器材名稱'],
    r.data['借用人'] + ' 您好，\n\n' +
    '您申請的器材預約' + (approved ? '已核准 ✅' : '未通過 ❌') + '：\n\n' +
    '器材：' + r.data['器材名稱'] + '（' + (Number(r.data['數量']) || 1) + ' 份）\n' +
    '期間：' + r.data['借出日'] + ' ~ ' + r.data['歸還日'] + '\n' +
    (note ? '備註：' + note + '\n' : '') +
    (approved ? '\n請於借出日前往領取器材。' : '\n如有疑問請聯繫器材管理人員。') + '\n\n' + APP_URL);

  return { ok: true };
}

// ====== 批次審核（管理端）：一次處理多筆、只寄一封彙整信 ======
function apiReviewBatch(p) {
  requireAdmin(p);
  requireFields(p, ['ids', 'decision']);
  var ids = p['ids'];
  if (typeof ids === 'string') { try { ids = JSON.parse(ids); } catch (e) {} }
  if (!ids || !ids.length) return { ok: false, error: '沒有項目' };

  var approved = p.decision === 'approve';
  var note = p['審核備註'] || '';
  var done = [], failed = [];
  var applicant = '', name = '', proj = '', period = '';

  for (var i = 0; i < ids.length; i++) {
    var r = findRowById(SHEET_RESV, ids[i]);
    if (!r.row || r.data['狀態'] !== '待審核') continue;
    if (approved) {
      var qty = Number(r.data['數量']) || 1;
      var a = availability(r.data['器材id'], r.data['借出日'], r.data['歸還日'], ids[i]);
      if (qty > a.free) { failed.push(r.data['器材名稱'] + '（剩 ' + a.free + '）'); continue; }
      setCell(SHEET_RESV, r.row, '狀態', '已核准');
    } else {
      setCell(SHEET_RESV, r.row, '狀態', '已拒絕');
    }
    if (note) setCell(SHEET_RESV, r.row, '審核備註', note);
    done.push({ 名稱: r.data['器材名稱'], 數量: Number(r.data['數量']) || 1 });
    applicant = r.data['聯絡方式']; name = r.data['借用人'];
    proj = r.data['用途']; period = r.data['借出日'] + ' ~ ' + r.data['歸還日'];
  }

  if (done.length) {
    var lines = done.map(function (d) { return '　• ' + d['名稱'] + ' ×' + d['數量']; }).join('\n');
    notify(applicant,
      '【器材預約】' + (approved ? '已核准' : '未通過') + '：' + (proj || name),
      name + ' 您好，\n\n您的預約' + (approved ? '已核准 ✅' : '未通過 ❌') + '：\n\n' +
      (proj ? '專案：' + proj + '\n' : '') + '期間：' + period + '\n\n' + lines + '\n' +
      (note ? '\n備註：' + note + '\n' : '') +
      (approved ? '\n請於借出日前往領取器材。' : '\n如有疑問請聯繫器材管理人員。') + '\n\n' + APP_URL);
  }
  return { ok: true, done: done.length, failed: failed };
}

// ====== 歸還（管理端）======
function apiReturn(p) {
  requireAdmin(p);
  requireFields(p, ['id']);
  var r = findRowById(SHEET_RESV, p['id']);
  if (!r.row) return { ok: false, error: '找不到申請' };
  setCell(SHEET_RESV, r.row, '狀態', '已歸還');
  return { ok: true };
}

// ====== 批次歸還（管理端）======
function apiReturnBatch(p) {
  requireAdmin(p);
  requireFields(p, ['ids']);
  var ids = p['ids'];
  if (typeof ids === 'string') { try { ids = JSON.parse(ids); } catch (e) {} }
  if (!ids || !ids.length) return { ok: false, error: '沒有項目' };
  var n = 0;
  for (var i = 0; i < ids.length; i++) {
    var r = findRowById(SHEET_RESV, ids[i]);
    if (r.row && r.data['狀態'] === '已核准') { setCell(SHEET_RESV, r.row, '狀態', '已歸還'); n++; }
  }
  return { ok: true, done: n };
}

// ====== 器材增刪改（管理端）======
function apiAddEquipment(p) {
  requireAdmin(p);
  requireFields(p, ['名稱']);
  var sh = sheet(SHEET_EQUIP);
  var id = 'E' + new Date().getTime();
  sh.appendRow([
    id, p['名稱'], p['分類'] || '', p['型號'] || '', p['序號'] || '',
    p['狀態'] || '可借用', p['照片網址'] || '', p['備註'] || '',
    p['群組'] || '', Number(p['數量']) || 1, p['附帶'] || ''
  ]);
  return { ok: true, id: id };
}

function apiUpdateEquipment(p) {
  requireAdmin(p);
  requireFields(p, ['id']);
  var r = findRowById(SHEET_EQUIP, p['id']);
  if (!r.row) return { ok: false, error: '找不到器材' };
  ['名稱','分類','型號','序號','狀態','照片網址','備註','群組','數量','附帶'].forEach(function (col) {
    if (p[col] != null) setCell(SHEET_EQUIP, r.row, col, p[col]);
  });
  return { ok: true };
}

function apiDeleteEquipment(p) {
  requireAdmin(p);
  requireFields(p, ['id']);
  var r = findRowById(SHEET_EQUIP, p['id']);
  if (!r.row) return { ok: false, error: '找不到器材' };
  sheet(SHEET_EQUIP).deleteRow(r.row);
  return { ok: true };
}

// ====== 數量感知的可借量計算 ======
// 回傳 { stock: 庫存, used: 該時段已核准佔用, free: 可借 }
function availability(equipId, start, end, excludeId) {
  var equip = findRowById(SHEET_EQUIP, equipId);
  var stock = equip.data ? (Number(equip.data['數量']) || 1) : 1;
  var used = 0;
  var rows = readSheet(SHEET_RESV);
  for (var i = 0; i < rows.length; i++) {
    var x = rows[i];
    if (x['器材id'] !== equipId) continue;
    if (excludeId && x['id'] === excludeId) continue;
    if (x['狀態'] !== '已核准') continue;
    // 接頭不算撞（A 12:00 還、B 12:00 借 → 不重疊）
    if (start < x['歸還日'] && end > x['借出日']) used += (Number(x['數量']) || 1);
  }
  return { stock: stock, used: used, free: Math.max(0, stock - used) };
}

// ====== Email 通知 ======
function getAdminEmail() {
  if (ADMIN_EMAIL) return ADMIN_EMAIL;
  try { return Session.getEffectiveUser().getEmail(); } catch (e) { return ''; }
}

function notify(to, subject, body) {
  if (!NOTIFY) return;
  if (!to || String(to).indexOf('@') < 0) return;
  try { MailApp.sendEmail({ to: String(to), subject: subject, body: body }); } catch (e) {}
}

function notifyAdminNew(items, p) {
  var lines = items.map(function (it) { return '　• ' + it['名稱'] + ' ×' + it['數量']; }).join('\n');
  notify(getAdminEmail(),
    '【器材預約】新申請待審核（' + items.length + ' 項）',
    '有新的器材預約申請待審核：\n\n' + lines + '\n\n' +
    '借用人：' + p['借用人'] + '（' + p['部門'] + '）\n' +
    '期間：' + p['借出日'] + ' ~ ' + p['歸還日'] + '\n' +
    '用途：' + (p['用途'] || '—') + '\n' +
    '聯絡：' + (p['聯絡方式'] || '—') + '\n\n' +
    '前往審核：' + APP_URL);
}

function testEmail() {
  var to = getAdminEmail();
  MailApp.sendEmail(to, '【器材預約】Email 通知測試', '收到代表通知功能正常運作。\n\n' + APP_URL);
  Logger.log('已寄出測試信至：' + to);
}

// ====== 工具 ======
function sheet(name) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('找不到工作表：' + name);
  return sh;
}

function readSheet(name) {
  var sh = sheet(name);
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0];
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var v = values[i][j];
      if (v instanceof Date && (headers[j] === '借出日' || headers[j] === '歸還日')) {
        v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
      } else if (v instanceof Date) {
        v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
      }
      obj[headers[j]] = v;
    }
    if (obj['id']) out.push(obj);
  }
  return out;
}

function findRowById(name, id) {
  var sh = sheet(name);
  var values = sh.getDataRange().getValues();
  var headers = values[0];
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      var data = {};
      for (var j = 0; j < headers.length; j++) data[headers[j]] = values[i][j];
      return { row: i + 1, data: data };
    }
  }
  return { row: 0, data: null };
}

function setCell(name, row, colName, value) {
  var sh = sheet(name);
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var col = headers.indexOf(colName) + 1;
  if (col > 0) sh.getRange(row, col).setValue(value);
}

function requireAdmin(p) {
  if (String(p.pin || '') !== String(ADMIN_PIN)) throw new Error('管理員 PIN 錯誤');
}
function requireFields(p, fields) {
  for (var i = 0; i < fields.length; i++) {
    if (!p[fields[i]] && p[fields[i]] !== 0) throw new Error('缺少必填欄位：' + fields[i]);
  }
}
function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * 一鍵建立/重置兩張表與標題列（新 schema）。跑一次後再跑 importInventory。
 */
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var eq = ss.getSheetByName(SHEET_EQUIP) || ss.insertSheet(SHEET_EQUIP);
  eq.clear();
  eq.appendRow(['id','名稱','分類','型號','序號','狀態','照片網址','備註','群組','數量','附帶']);
  eq.setFrozenRows(1);

  var rv = ss.getSheetByName(SHEET_RESV) || ss.insertSheet(SHEET_RESV);
  rv.clear();
  rv.appendRow(['id','器材id','器材名稱','借用人','部門','聯絡方式','借出日','歸還日','用途','狀態','申請時間','審核備註','數量','批次']);
  rv.setFrozenRows(1);
}

/**
 * 非破壞性升級（可重複執行）：
 *  1) 預約表補「批次」欄
 *  2) 器材表補「附帶」欄，並把看起來像配件清單的「備註」自動搬到「附帶」
 *     （判斷：備註含 、 × ＋ 這類清單符號才搬；純文字備註如「油壓雲台」「單獨」保留）
 * 不刪資料。
 */
function upgradeSchema() {
  var rv = sheet(SHEET_RESV);
  var rh = rv.getRange(1, 1, 1, rv.getLastColumn()).getValues()[0];
  if (rh.indexOf('批次') < 0) rv.getRange(1, rv.getLastColumn() + 1).setValue('批次');

  var eq = sheet(SHEET_EQUIP);
  var eh = eq.getRange(1, 1, 1, eq.getLastColumn()).getValues()[0];
  if (eh.indexOf('附帶') < 0) eq.getRange(1, eq.getLastColumn() + 1).setValue('附帶');
  eh = eq.getRange(1, 1, 1, eq.getLastColumn()).getValues()[0];
  var cNote = eh.indexOf('備註') + 1, cAcc = eh.indexOf('附帶') + 1;
  var last = eq.getLastRow();
  if (last > 1 && cNote > 0 && cAcc > 0) {
    var notes = eq.getRange(2, cNote, last - 1, 1).getValues();
    var accs  = eq.getRange(2, cAcc,  last - 1, 1).getValues();
    for (var i = 0; i < notes.length; i++) {
      var n = String(notes[i][0]);
      if (!accs[i][0] && /[、×＋]/.test(n)) { accs[i][0] = n; notes[i][0] = ''; }
    }
    eq.getRange(2, cNote, last - 1, 1).setValues(notes);
    eq.getRange(2, cAcc,  last - 1, 1).setValues(accs);
  }
  Logger.log('升級完成');
}

/**
 * 一次性：把所有「相機」拆成機身（清空附帶），並將電池/相機袋/充電器等
 * 配件建成獨立可借器材（電源/配件/腳架…）。可重複執行，同名已存在會略過。
 * 在編輯器執行即可，不需重新部署。
 */
function splitCameraKits() {
  var eq = sheet(SHEET_EQUIP);
  var values = eq.getDataRange().getValues();
  var h = values[0];
  var cName = h.indexOf('名稱'), cCat = h.indexOf('分類'), cAcc = h.indexOf('附帶');

  // 1) 清空所有相機的附帶
  for (var i = 1; i < values.length; i++) {
    if (values[i][cCat] === '相機') eq.getRange(i + 1, cAcc + 1).setValue('');
  }

  // 2) 建立獨立配件（名稱, 分類, 數量）
  var add = [
    ['Sony 電池 NP-FZ100', '電源', 5], ['Sony 充電器', '電源', 3], ['Sony 相機袋', '配件', 1],
    ['Canon 電池', '電源', 4], ['Canon 假電池', '電源', 1], ['Canon 充電器', '電源', 3], ['Canon 相機袋', '配件', 1],
    ['小相機 電池', '電源', 2], ['小相機 充電器', '電源', 1], ['小相機 備用電池(非原廠)', '電源', 1],
    ['GoPro 腳架', '腳架', 1], ['GoPro 充電器', '電源', 1], ['GoPro 兔籠', '配件', 1],
    ['Pocket3 轉接底座', '配件', 1], ['Pocket3 充電柄', '電源', 1], ['Pocket3 小腳架', '腳架', 1],
    ['Pocket3 原廠袋', '配件', 1], ['Pocket3 mic', '音訊', 1]
  ];
  var existing = {};
  for (var j = 1; j < values.length; j++) existing[values[j][cName]] = true;

  add.forEach(function (a) {
    if (existing[a[0]]) return;
    var id = 'E' + new Date().getTime() + Math.floor(Math.random() * 1000);
    eq.appendRow([id, a[0], a[1], '', '', '可借用', '', '', '', a[2], '']);
  });
  Logger.log('相機已拆成機身，配件已獨立');
}

/**
 * 一次性：把 Pocket 兩台設成套裝。
 * 移除先前 splitCameraKits 拆出的 Pocket3 獨立配件，並把配件寫回兩台 Pocket 的「附帶」。
 * 在編輯器執行即可，不需重新部署。
 */
function fixPocketKits() {
  var eq = sheet(SHEET_EQUIP);
  var values = eq.getDataRange().getValues();
  var h = values[0];
  var cName = h.indexOf('名稱'), cAcc = h.indexOf('附帶');

  // 刪掉拆出的 Pocket3 獨立配件（由下往上刪，避免列號位移）
  var del = ['Pocket3 轉接底座', 'Pocket3 充電柄', 'Pocket3 小腳架', 'Pocket3 原廠袋', 'Pocket3 mic'];
  for (var i = values.length - 1; i >= 1; i--) {
    if (del.indexOf(values[i][cName]) >= 0) eq.deleteRow(i + 1);
  }

  // 設定兩台 Pocket 的附帶（整組借）
  var acc = {
    'DJI Pocket 3 全套': '轉接底座、充電柄、小腳架、原廠袋、Pocket mic',
    'DJI Pocket 4 全套': '轉接底座、充電柄、小腳架、原廠袋、Pocket mic'
  };
  var v2 = eq.getDataRange().getValues();
  for (var j = 1; j < v2.length; j++) {
    if (acc[v2[j][cName]]) eq.getRange(j + 1, cAcc + 1).setValue(acc[v2[j][cName]]);
  }
  Logger.log('Pocket 兩台已設為套裝');
}

/**
 * 一鍵匯入真實器材清單（新 schema，含 群組/數量）。
 * 會清空「器材」表舊資料（保留標題列）並重寫。不影響「預約」表。
 * 每筆：[名稱, 分類, 型號, 序號, 狀態, 照片網址, 備註, 群組, 數量]
 */
function importInventory() {
  var rows = [
    // ===== 🆕 新增器材 =====
    ['Sony FX3',            '相機',  'ILME-FX3', '', '可借用', '', '（新）機身', '', 1],
    ['24-70mm F2.8',        '鏡頭',  '',         '', '可借用', '', '（新）', '', 1],
    ['DJI Pocket 4 全套',   '相機',  'Pocket 4', '', '可借用', '', '（新）含全部配件，整組借', '', 1],
    ['DJI Mic 3 整套',      '音訊',  'Mic 3',    '', '可借用', '', '（新）整套', '', 1],
    ['DJI RS 4 Pro 穩定器', '穩定器', 'RS 4 Pro', '', '可借用', '', '（新）整組', '', 1],

    // ===== 📷 相機系統（整組借）=====
    ['Sony A7S III 全套', '相機', 'ILCE-7SM3', '', '可借用', '', '機身、電池×5、充電器×3、相機袋', '', 1],
    ['Canon 相機A 全套',  '相機', '',          '', '可借用', '', '機身、假電池、電池×4、充電器×3、相機袋', '', 1],
    ['Sony 小相機 全套',  '相機', '',          '', '可借用', '', '小相機、配件、充電器、電池×2、備用電池(非原廠)', '', 1],
    ['GoPro 全套',        '相機', '',          '', '可借用', '', '相機、腳架、充電器(附線)、兔籠', '', 1],
    ['DJI Pocket 3 全套', '相機', 'Pocket 3',  '', '可借用', '', '相機、轉接底座、充電柄、小腳架、原廠袋、Pocket mic', '', 1],
    ['iPhone 12 Pro',     '相機', '',          '', '可借用', '', '單獨', '', 1],

    // ===== 🎙 音訊 =====
    ['DJI Mic 一代 整套', '音訊', '', '', '可借用', '', '收納盒、相機用×1、人員用×2、耳機盒(附耳機)、3.5mm線、防風兔毛', '', 1],
    ['Rode 無線麥組',     '音訊', '', '', '可借用', '', '無線麥A、無線麥B', '', 1],
    ['Rode 指向麥組',     '音訊', '', '', '可借用', '', 'Rode麥克風A、小指向麥A、小指向麥B', '', 1],
    ['Yeti 電容麥',       '音訊', '', '', '可借用', '', '單獨', '', 1],
    ['領夾麥克風',        '音訊', '', '', '可借用', '', '單獨', '', 1],

    // ===== 💡 燈光 =====
    ['Aputure 主燈 整組', '燈光', '', '', '可借用', '', '主燈、電源線×2、控制器、黑燈罩、燈架、柔光罩、柔光布×2、蜂巢網、保護罩、燈罩手提袋、手提箱', '', 1],
    ['便攜燈A 整組',      '燈光', '', '', '可借用', '', '燈、NP-F750充電器、柔光傘、電池×2', '', 1],
    ['豆腐燈',            '燈光', '', '', '可借用', '', '燈＋保護套', '', 1],
    ['背景燈C',           '燈光', '', '', '可借用', '', '單獨', '', 1],

    // ===== 🦵 腳架 =====
    ['腳架A',            '腳架', '', '', '可借用', '', '', '', 1],
    ['腳架B（油壓）',    '腳架', '', '', '可借用', '', '油壓雲台', '', 1],
    ['腳架C',            '腳架', '', '', '可借用', '', '', '', 1],
    ['腳架D',            '腳架', '', '', '可借用', '', '', '', 1],
    ['腳架E',            '腳架', '', '', '可借用', '', '', '', 1],
    ['手持/手機小腳架組', '腳架', '', '', '可借用', '', '手持三腳架A、B、手持手機腳架A、燈架A、黑色腳架帶', '', 1],

    // ===== 💾 儲存 =====
    // 記憶卡：群組收合，點開選容量
    ['SD 256G',         '儲存', 'SD',      '', '可借用', '', '', '記憶卡', 2],
    ['SD 128G',         '儲存', 'SD',      '', '可借用', '', '', '記憶卡', 2],
    ['SD 128G (300MB/s)','儲存', 'SD',     '', '可借用', '', '高速卡', '記憶卡', 1],
    ['microSD 256G',    '儲存', 'microSD', '', '可借用', '', '', '記憶卡', 1],
    ['microSD 128G',    '儲存', 'microSD', '', '可借用', '', '', '記憶卡', 2],
    // SSD 也照容量分
    ['SSD 1TB',         '儲存', 'SSD', '', '可借用', '', '', '硬碟', 1],
    ['SSD 2TB',         '儲存', 'SSD', '', '可借用', '', '', '硬碟', 1],
    ['隨身碟 128G',     '儲存', '',    '', '可借用', '', '', '', 3],
    ['記憶卡殼',        '儲存', '',    '', '可借用', '', '', '', 5],
    ['讀卡機組',        '儲存', '',    '', '可借用', '', '讀卡機A/B/C、三合一讀卡機A', '', 1],

    // ===== 🔌 電源 / 配件 =====
    ['行動電源A', '電源', '', '', '可借用', '', '', '', 1],
    ['行動電源B', '電源', '', '', '可借用', '', '', '', 1],
    ['行動電源C', '電源', '', '', '可借用', '', '', '', 1],
    ['行動電源D', '電源', '', '', '可借用', '', '', '', 1],
    ['線材轉接配件箱', '配件', '', '', '可借用', '', '動力延長線×2、多孔延長線×2、hub轉接器×5、type c hub×3、USB-C轉接頭×2、40W豆腐頭、三頭線、雙頭線、電池充電線、萬用轉接插頭、USB電源插頭、Sony假電池×2、充電袋、車用手機支架', '', 1],

    // ===== 🎬 其他 =====
    ['場記板A', '雜項', '', '', '可借用', '', '單獨', '', 1]
  ];

  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EQUIP);
  if (!sh) throw new Error('找不到「器材」工作表，請先跑 setupSheets');
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
  var out = rows.map(function (r, i) { return ['E' + (i + 1)].concat(r); });
  sh.getRange(2, 1, out.length, out[0].length).setValues(out);
  Logger.log('已匯入 ' + out.length + ' 筆器材');
}
