/**
 * 攝影器材管理與預約系統 — 後端 API
 * 部署方式：Google Apps Script Web App（見 README「部署步驟」）
 *
 * 資料表結構（在同一個 Google Sheet 裡建兩個工作表）：
 *
 * 工作表「器材」欄位（第 1 列為標題）：
 *   A:id  B:名稱  C:分類  D:型號  E:序號  F:狀態  G:照片網址  H:備註
 *   狀態值：可借用 / 維修中 / 報廢
 *
 * 工作表「預約」欄位（第 1 列為標題）：
 *   A:id  B:器材id  C:器材名稱  D:借用人  E:部門  F:聯絡方式
 *   G:借出日  H:歸還日  I:用途  J:狀態  K:申請時間  L:審核備註  M:申請人Email
 *   狀態值：待審核 / 已核准 / 已拒絕 / 已歸還 / 已取消
 *   申請人Email：系統自動抓登入者公司帳號（部署為「機構內任何人」時有效），防造假
 */

// ====== 設定 ======
// 管理員 PIN：請改成你自己的密碼，管理端操作需要它
var ADMIN_PIN = '1234';

var SHEET_EQUIP = '器材';
var SHEET_RESV  = '預約';

// ====== 入口 ======
function doGet(e) {
  return handle(e);
}
function doPost(e) {
  return handle(e);
}

function handle(e) {
  var params = {};
  if (e && e.parameter) params = e.parameter;
  // POST 帶 JSON body 時優先解析
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
      case 'list':               out = apiList(); break;
      case 'create_reservation': out = apiCreateReservation(params); break;
      case 'cancel_reservation': out = apiCancelReservation(params); break;
      // 以下為管理端（需 PIN）
      case 'review':             out = apiReview(params); break;
      case 'return':             out = apiReturn(params); break;
      case 'add_equipment':      out = apiAddEquipment(params); break;
      case 'update_equipment':   out = apiUpdateEquipment(params); break;
      case 'delete_equipment':   out = apiDeleteEquipment(params); break;
      default: out = { ok: false, error: '未知的 action: ' + action };
    }
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  return json(out);
}

// ====== 讀取（公開）======
function apiList() {
  return {
    ok: true,
    equipment: readSheet(SHEET_EQUIP),
    reservations: readSheet(SHEET_RESV)
  };
}

// ====== 建立預約申請（公開）======
function apiCreateReservation(p) {
  requireFields(p, ['器材id', '借用人', '部門', '借出日', '歸還日']);
  var equip = findRowById(SHEET_EQUIP, p['器材id']);
  if (!equip.row) return { ok: false, error: '找不到該器材' };
  if (equip.data['狀態'] !== '可借用') {
    return { ok: false, error: '該器材目前狀態為「' + equip.data['狀態'] + '」，無法預約' };
  }
  if (p['借出日'] > p['歸還日']) return { ok: false, error: '歸還日不可早於借出日' };

  // 撞期檢查：同器材、狀態為已核准且日期區間重疊
  var conflict = findConflict(p['器材id'], p['借出日'], p['歸還日'], null);
  if (conflict) {
    return { ok: false, error: '該時段已被核准借用（' + conflict['借出日'] + ' ~ ' + conflict['歸還日'] + '），請改期' };
  }

  var email = '';
  try { email = Session.getActiveUser().getEmail() || ''; } catch (e) {}

  var sh = sheet(SHEET_RESV);
  var id = 'R' + new Date().getTime();
  sh.appendRow([
    id, p['器材id'], equip.data['名稱'], p['借用人'], p['部門'],
    p['聯絡方式'] || '', p['借出日'], p['歸還日'], p['用途'] || '',
    '待審核', new Date(), '', email
  ]);
  return { ok: true, id: id };
}

// ====== 取消申請（公開，需帶申請 id）======
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
  requireFields(p, ['id', 'decision']); // decision: approve / reject
  var r = findRowById(SHEET_RESV, p['id']);
  if (!r.row) return { ok: false, error: '找不到申請' };
  if (r.data['狀態'] !== '待審核') return { ok: false, error: '此申請不是待審核狀態' };

  if (p.decision === 'approve') {
    var conflict = findConflict(r.data['器材id'], r.data['借出日'], r.data['歸還日'], p['id']);
    if (conflict) {
      return { ok: false, error: '撞期！已有核准借用（' + conflict['借出日'] + ' ~ ' + conflict['歸還日'] + '），無法核准' };
    }
    setCell(SHEET_RESV, r.row, '狀態', '已核准');
  } else {
    setCell(SHEET_RESV, r.row, '狀態', '已拒絕');
  }
  if (p['審核備註'] != null) setCell(SHEET_RESV, r.row, '審核備註', p['審核備註']);
  return { ok: true };
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

// ====== 器材增刪改（管理端）======
function apiAddEquipment(p) {
  requireAdmin(p);
  requireFields(p, ['名稱']);
  var sh = sheet(SHEET_EQUIP);
  var id = 'E' + new Date().getTime();
  sh.appendRow([
    id, p['名稱'], p['分類'] || '', p['型號'] || '', p['序號'] || '',
    p['狀態'] || '可借用', p['照片網址'] || '', p['備註'] || ''
  ]);
  return { ok: true, id: id };
}

function apiUpdateEquipment(p) {
  requireAdmin(p);
  requireFields(p, ['id']);
  var r = findRowById(SHEET_EQUIP, p['id']);
  if (!r.row) return { ok: false, error: '找不到器材' };
  ['名稱','分類','型號','序號','狀態','照片網址','備註'].forEach(function(col){
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

// ====== 工具函式 ======
function findConflict(equipId, start, end, excludeId) {
  var rows = readSheet(SHEET_RESV);
  for (var i = 0; i < rows.length; i++) {
    var x = rows[i];
    if (x['器材id'] !== equipId) continue;
    if (excludeId && x['id'] === excludeId) continue;
    if (x['狀態'] !== '已核准') continue;
    // 日期字串為 YYYY-MM-DD 可直接比較；重疊條件：start <= x.end 且 end >= x.start
    if (start <= x['歸還日'] && end >= x['借出日']) return x;
  }
  return null;
}

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
      // 日期欄位轉成 YYYY-MM-DD 字串
      if (v instanceof Date && (headers[j] === '借出日' || headers[j] === '歸還日')) {
        v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
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
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 一鍵初始化：在編輯器選這個函式按執行，會自動建立兩張表與標題列，並放入範例器材。
 */
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var eq = ss.getSheetByName(SHEET_EQUIP) || ss.insertSheet(SHEET_EQUIP);
  eq.clear();
  eq.appendRow(['id','名稱','分類','型號','序號','狀態','照片網址','備註']);
  eq.appendRow(['E1','Sony A7 IV','相機','A7M4','SN-0001','可借用','','公司主力機身']);
  eq.appendRow(['E2','24-70mm f/2.8 GM','鏡頭','SEL2470GM2','SN-0002','可借用','','標準變焦']);
  eq.appendRow(['E3','Aputure 600D','燈光','600D','SN-0003','可借用','','棚拍主燈']);
  eq.setFrozenRows(1);

  var rv = ss.getSheetByName(SHEET_RESV) || ss.insertSheet(SHEET_RESV);
  rv.clear();
  rv.appendRow(['id','器材id','器材名稱','借用人','部門','聯絡方式','借出日','歸還日','用途','狀態','申請時間','審核備註','申請人Email']);
  rv.setFrozenRows(1);
}
