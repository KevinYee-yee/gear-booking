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
 *   A:id  B:器材id  C:器材名稱  D:借用人  E:部門  F:聯絡方式(Email)
 *   G:借出日  H:歸還日  I:用途  J:狀態  K:申請時間  L:審核備註
 *   狀態值：待審核 / 已核准 / 已拒絕 / 已歸還 / 已取消
 *   聯絡方式(Email)：申請人填的 email，用來寄核准/拒絕通知
 */

// ====== 設定 ======
// 管理員 PIN：請改成你自己的密碼，管理端操作需要它
var ADMIN_PIN = '1234';

// Email 通知設定
var NOTIFY = true;              // 是否開啟 email 通知
var ADMIN_EMAIL = '';           // 收「新申請通知」的信箱；留空＝寄給指令碼擁有者（你）
var APP_URL = 'https://kevinyee-yee.github.io/gear-booking/'; // 通知信附上的系統連結

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

  var sh = sheet(SHEET_RESV);
  var id = 'R' + new Date().getTime();
  sh.appendRow([
    id, p['器材id'], equip.data['名稱'], p['借用人'], p['部門'],
    p['聯絡方式'] || '', p['借出日'], p['歸還日'], p['用途'] || '',
    '待審核', new Date(), ''
  ]);

  // 通知管理員有新申請
  notify(getAdminEmail(),
    '【器材預約】新申請待審核：' + equip.data['名稱'],
    '有一筆新的器材預約申請，請審核：\n\n' +
    '器材：' + equip.data['名稱'] + '\n' +
    '借用人：' + p['借用人'] + '（' + p['部門'] + '）\n' +
    '期間：' + p['借出日'] + ' ~ ' + p['歸還日'] + '\n' +
    '用途：' + (p['用途'] || '—') + '\n' +
    '聯絡：' + (p['聯絡方式'] || '—') + '\n\n' +
    '前往審核：' + APP_URL);

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

  var approved;
  if (p.decision === 'approve') {
    var conflict = findConflict(r.data['器材id'], r.data['借出日'], r.data['歸還日'], p['id']);
    if (conflict) {
      return { ok: false, error: '撞期！已有核准借用（' + conflict['借出日'] + ' ~ ' + conflict['歸還日'] + '），無法核准' };
    }
    setCell(SHEET_RESV, r.row, '狀態', '已核准');
    approved = true;
  } else {
    setCell(SHEET_RESV, r.row, '狀態', '已拒絕');
    approved = false;
  }
  var note = (p['審核備註'] != null) ? p['審核備註'] : '';
  if (p['審核備註'] != null) setCell(SHEET_RESV, r.row, '審核備註', note);

  // 通知申請人結果（聯絡方式須為 email）
  notify(r.data['聯絡方式'],
    '【器材預約】' + (approved ? '已核准' : '未通過') + '：' + r.data['器材名稱'],
    r.data['借用人'] + ' 您好，\n\n' +
    '您申請的器材預約' + (approved ? '已核准 ✅' : '未通過 ❌') + '：\n\n' +
    '器材：' + r.data['器材名稱'] + '\n' +
    '期間：' + r.data['借出日'] + ' ~ ' + r.data['歸還日'] + '\n' +
    (note ? '備註：' + note + '\n' : '') +
    (approved ? '\n請於借出日前往領取器材。' : '\n如有疑問請聯繫器材管理人員。') + '\n\n' + APP_URL);

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

// ====== Email 通知 ======
function getAdminEmail() {
  if (ADMIN_EMAIL) return ADMIN_EMAIL;
  try { return Session.getEffectiveUser().getEmail(); } catch (e) { return ''; }
}

function notify(to, subject, body) {
  if (!NOTIFY) return;
  if (!to || String(to).indexOf('@') < 0) return; // 非有效 email 就跳過
  try {
    MailApp.sendEmail({ to: String(to), subject: subject, body: body });
  } catch (e) {
    // 寄信失敗不影響主流程（例如配額用盡）
  }
}

/**
 * 測試 Email：在編輯器選這個函式按執行，會寄一封測試信到管理員信箱，
 * 並在第一次觸發時要求授權寄信權限（請允許）。
 */
function testEmail() {
  var to = getAdminEmail();
  MailApp.sendEmail(to, '【器材預約】Email 通知測試', '這是一封測試信，收到代表通知功能正常運作。\n\n系統：' + APP_URL);
  Logger.log('已寄出測試信至：' + to);
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 一鍵匯入真實器材清單：在編輯器選這個函式按執行。
 * 會清空「器材」表的舊資料（保留標題列）並寫入下方整理好的清單。
 * 不影響「預約」表。跑完不需重新部署。
 *
 * 每筆：[名稱, 分類, 型號, 序號, 狀態, 照片網址, 備註]
 * 「整組借」的配件寫在備註，借用時整組一起借。
 */
function importInventory() {
  var rows = [
    // ===== 🆕 新增器材（精簡，不細標編號）=====
    ['Sony FX3',            '相機',  'ILME-FX3', '', '可借用', '', '（新）機身'],
    ['24-70mm F2.8',        '鏡頭',  '',         '', '可借用', '', '（新）'],
    ['DJI Pocket 4 全套',   '相機',  'Pocket 4', '', '可借用', '', '（新）含全部配件，整組借'],
    ['DJI Mic 3 整套',      '音訊',  'Mic 3',    '', '可借用', '', '（新）整套'],
    ['DJI RS 4 Pro 穩定器', '穩定器', 'RS 4 Pro', '', '可借用', '', '（新）整組'],

    // ===== 📷 相機系統（機身＋專屬電池/充電器/袋，整組借）=====
    ['Sony A7S III 全套', '相機', 'ILCE-7SM3', '', '可借用', '', '機身、電池×5、充電器×3、相機袋'],
    ['Canon 相機A 全套',  '相機', '',          '', '可借用', '', '機身、假電池、電池×4、充電器×3、相機袋'],
    ['Sony 小相機 全套',  '相機', '',          '', '可借用', '', '小相機、配件、充電器、電池×2、備用電池(非原廠)'],
    ['GoPro 全套',        '相機', '',          '', '可借用', '', '相機、腳架、充電器(附線)、兔籠'],
    ['DJI Pocket 3 全套', '相機', 'Pocket 3',  '', '可借用', '', '相機、轉接底座、充電柄、小腳架、原廠袋、Pocket mic'],
    ['iPhone 12 Pro',     '相機', '',          '', '可借用', '', '單獨'],

    // ===== 🎙 音訊 =====
    ['DJI Mic 一代 整套', '音訊', '', '', '可借用', '', '收納盒、相機用×1、人員用×2、耳機盒(附耳機)、3.5mm線、防風兔毛'],
    ['Rode 無線麥組',     '音訊', '', '', '可借用', '', '無線麥A、無線麥B'],
    ['Rode 指向麥組',     '音訊', '', '', '可借用', '', 'Rode麥克風A、小指向麥A、小指向麥B'],
    ['Yeti 電容麥',       '音訊', '', '', '可借用', '', '單獨'],
    ['領夾麥克風',        '音訊', '', '', '可借用', '', '單獨'],

    // ===== 💡 燈光 =====
    ['Aputure 主燈 整組', '燈光', '', '', '可借用', '', '主燈、電源線×2、控制器、黑燈罩、燈架、柔光罩、柔光布×2、蜂巢網、保護罩、燈罩手提袋、手提箱'],
    ['便攜燈A 整組',      '燈光', '', '', '可借用', '', '燈、NP-F750充電器、柔光傘、電池×2'],
    ['豆腐燈',            '燈光', '', '', '可借用', '', '燈＋保護套'],
    ['背景燈C',           '燈光', '', '', '可借用', '', '單獨'],

    // ===== 🦵 腳架（個別 5 支 + 小腳架組）=====
    ['腳架A',           '腳架', '', '', '可借用', '', ''],
    ['腳架B（油壓）',   '腳架', '', '', '可借用', '', '油壓雲台'],
    ['腳架C',           '腳架', '', '', '可借用', '', ''],
    ['腳架D',           '腳架', '', '', '可借用', '', ''],
    ['腳架E',           '腳架', '', '', '可借用', '', ''],
    ['手持/手機小腳架組', '腳架', '', '', '可借用', '', '手持三腳架A、B、手持手機腳架A、燈架A、黑色腳架帶'],

    // ===== 💾 儲存（分類成組）=====
    ['SD 記憶卡組',    '儲存', '', '', '可借用', '', 'SD：256G×2、128G×2、128G/300ms×1'],
    ['microSD 記憶卡組','儲存', '', '', '可借用', '', 'microSD：256G×1、128G×2'],
    ['SSD 硬碟組',     '儲存', '', '', '可借用', '', '1TB、2TB'],
    ['隨身碟組',       '儲存', '', '', '可借用', '', '128G×3'],
    ['記憶卡殼組',     '儲存', '', '', '可借用', '', '卡殼×5'],
    ['讀卡機組',       '儲存', '', '', '可借用', '', '讀卡機A/B/C、三合一讀卡機A'],

    // ===== 🔌 電源 / 配件 =====
    ['行動電源A', '電源', '', '', '可借用', '', ''],
    ['行動電源B', '電源', '', '', '可借用', '', ''],
    ['行動電源C', '電源', '', '', '可借用', '', ''],
    ['行動電源D', '電源', '', '', '可借用', '', ''],
    ['線材轉接配件箱', '配件', '', '', '可借用', '', '動力延長線×2、多孔延長線×2、hub轉接器×5、type c hub×3、USB-C轉接頭×2、40W豆腐頭、三頭線、雙頭線、電池充電線、萬用轉接插頭、USB電源插頭、Sony假電池×2、充電袋、車用手機支架'],

    // ===== 🎬 其他 =====
    ['場記板A', '雜項', '', '', '可借用', '', '單獨']
  ];

  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EQUIP);
  if (!sh) throw new Error('找不到「器材」工作表，請先跑 setupSheets');
  // 清掉舊資料（保留標題列）
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
  // 寫入：第一欄補上 id（E1, E2, …）
  var out = rows.map(function (r, i) { return ['E' + (i + 1)].concat(r); });
  sh.getRange(2, 1, out.length, out[0].length).setValues(out);
  Logger.log('已匯入 ' + out.length + ' 筆器材');
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
  rv.appendRow(['id','器材id','器材名稱','借用人','部門','聯絡方式','借出日','歸還日','用途','狀態','申請時間','審核備註']);
  rv.setFrozenRows(1);
}
