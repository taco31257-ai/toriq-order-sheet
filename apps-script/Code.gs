/**
 * 鳥久 × 隼｜店內管理系統 — Google 試算表雲端同步（v3.2：保留單位文字的安全庫存及兩店各自商品備註）
 *
 * 第一次使用／更新程式：
 *   1. 把下面 SYNC_KEY 改成你的同步密碼（與舊版相同即可，手機不用重新輸入）
 *   2. 函式選單選「setup」→ 執行（會自動把舊版分店格式轉成共用格式，資料不會遺失）
 *   3. setup 會自動完成：品項是空的就匯入成本主表 349 個品項；現有酒款一次標成兩店都有（只做一次）
 *   4. 部署 → 管理部署作業 → 編輯 → 版本選「新版本」→ 部署（網址不變）
 */
const SYNC_KEY = '請改成你的同步密碼';

const STORES = ['隼', '鳥久'];
const TZ = 'Asia/Taipei';
const SH = { vendors: '廠商', items: '品項', wines: '酒款', log: '叫貨紀錄', stats: '月統計', help: '使用說明' };
const HEAD = {
  vendors: ['廠商ID', '廠商', '送貨星期', '注意事項'],
  items: ['品項ID', '廠商ID', '廠商', '品名', '單價', '單位', '隼使用台', '隼平日', '隼假日', '鳥久使用台', '鳥久平日', '鳥久假日', '隼商品備註', '鳥久商品備註'],
  wines: ['酒款ID', '酒名', '酒種', '容量ml', '售價', '杯價', '進價', '代理商', '料號', '介紹', '隼使用', '鳥久使用'],
  log: ['單號', '類型', '門市', '叫貨時間', '送貨日期', '廠商', '叫貨人', '品項', '數量', '單位', '單價', '小計', '原始資料']
};
const ITEM_STORE_COL = { '隼': 6, '鳥久': 9 };   // 使用台、平日、假日（0 起算）
const WINE_STORE_COL = { '隼': 10, '鳥久': 11 };
const ITEM_NOTE_COL = { '隼': 12, '鳥久': 13 };

/* ================= 初始化 ================= */
function setup() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(TZ);
  migrateV2_();
  const v = ensureSheet_('vendors'); v.hideColumns(1); v.setColumnWidth(2, 220); v.setColumnWidth(4, 340);
  const it = ensureSheet_('items'); it.hideColumns(1, 2); it.setColumnWidth(3, 180); it.setColumnWidth(4, 260);
  const w = ensureSheet_('wines'); w.hideColumns(1); w.setColumnWidth(2, 260); w.setColumnWidth(10, 320);
  const log = ensureSheet_('log');
  log.getRange('D:D').setNumberFormat('yyyy-mm-dd hh:mm'); log.setColumnWidth(8, 260); log.hideColumns(13);
  buildStats_(); buildHelp_(true);
  oneTime_();
  ['工作表1', 'Sheet1'].forEach(n => { const s = ss.getSheetByName(n); if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s); });
  ss.setActiveSheet(it);
}

function ensureSheet_(k) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SH[k]);
  if (!sh) sh = ss.insertSheet(SH[k]);
  if (sh.getLastRow() === 0) sh.setFrozenRows(1);
  // 新增欄位只更新標題，不移動或清除舊資料（品項表新增兩店商品備註）
  sh.getRange(1, 1, 1, HEAD[k].length).setValues([HEAD[k]]).setFontWeight('bold').setBackground('#e9eee7');
  return sh;
}

/* 舊版（每列有「門市」欄）轉成共用格式 */
function migrateV2_() {
  const ss = SpreadsheetApp.getActive();
  const vs = ss.getSheetByName(SH.vendors), is = ss.getSheetByName(SH.items);
  if (!vs || !is || String(vs.getRange(1, 1).getValue()) !== '門市') return;
  const oldV = vs.getLastRow() > 1 ? vs.getRange(2, 1, vs.getLastRow() - 1, 5).getValues() : [];
  const oldI = is.getLastRow() > 1 ? is.getRange(2, 1, is.getLastRow() - 1, 10).getValues() : [];
  const vByName = {}, vIdMap = {}, vRows = [];
  oldV.forEach(r => {
    const name = String(r[2]).trim(); if (!name) return;
    if (!vByName[name]) { vByName[name] = String(r[1] || Utilities.getUuid()); vRows.push([vByName[name], name, r[3], r[4]]); }
    vIdMap[String(r[1])] = vByName[name];
  });
  const iByKey = {}, iRows = [];
  oldI.forEach(r => {
    const store = r[0], vname = String(r[3]).trim(), name = String(r[5]).trim(); if (!name || !ITEM_STORE_COL[store]) return;
    const k = vname + '|' + name;
    let row = iByKey[k];
    if (!row) {
      row = [String(r[1] || Utilities.getUuid()), vIdMap[String(r[2])] || vByName[vname] || '', vname, name, r[6], r[7], '', '', '', '', '', ''];
      iByKey[k] = row; iRows.push(row);
    }
    const c = ITEM_STORE_COL[store]; row[c] = r[4] || '備品'; row[c + 1] = r[8]; row[c + 2] = r[9];
  });
  vs.clear(); is.clear();
  ensureSheet_('vendors'); ensureSheet_('items');
  if (vRows.length) vs.getRange(2, 1, vRows.length, 4).setValues(vRows);
  if (iRows.length) is.getRange(2, 1, iRows.length, 12).setValues(iRows);
}

function buildStats_() {
  const ss = SpreadsheetApp.getActive();
  if (ss.getSheetByName(SH.stats)) return;
  const sh = ss.insertSheet(SH.stats);
  sh.getRange('A1:A2').setValues([['統計月份'], ['門市']]).setFontWeight('bold');
  sh.getRange('B1').setNumberFormat('@').setValue(Utilities.formatDate(new Date(), TZ, 'yyyy-MM'));
  sh.getRange('C1').setValue('← 例 2026-11').setFontColor('#777777');
  sh.getRange('C2').setValue('← 鳥久 / 隼，空白＝兩店合計').setFontColor('#777777');
  const rng = "'" + SH.log + "'!A:L";
  const start = 'DATE(VALUE(LEFT($B$1,4)),VALUE(RIGHT($B$1,2)),1)';
  const where = " where D >= datetime '\"&TEXT(" + start + ",\"yyyy-mm-dd\")&\" 00:00:00' and D < datetime '\"&TEXT(EDATE(" + start + ",1),\"yyyy-mm-dd\")&\" 00:00:00' and H is not null\"&IF($B$2=\"\",\"\",\" and C = '\"&$B$2&\"'\")&\"";
  const q = (cols, tail) => '=IFERROR(QUERY(' + rng + ',"select ' + cols + where + tail + '",1),"本月尚無紀錄")';
  sh.getRange('A4').setValue('依廠商').setFontWeight('bold');
  sh.getRange('A5').setFormula(q('C, B, F, count(H), sum(L)', " group by C, B, F order by C, B, F label C '門市', B '類型', F '廠商', count(H) '品項筆數', sum(L) '估算金額'"));
  sh.getRange('H4').setValue('依品項（用量）').setFontWeight('bold');
  sh.getRange('H5').setFormula(q('C, F, H, J, sum(I), sum(L)', " group by C, F, H, J order by C, F, H label C '門市', F '廠商', H '品項', J '單位', sum(I) '總數量', sum(L) '估算金額'"));
  sh.getRange('P4').setValue('依叫貨人').setFontWeight('bold');
  sh.getRange('P5').setFormula(q('C, G, count(H), sum(L)', " group by C, G order by C label C '門市', G '叫貨人', count(H) '品項筆數', sum(L) '估算金額'"));
  sh.setFrozenRows(5);
}

function buildHelp_(rebuild) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SH.help);
  if (sh && !rebuild) return;
  if (sh) sh.clear(); else sh = ss.insertSheet(SH.help);
  const t = [
    ['【共用資料】「廠商」「品項」「酒款」兩店共用一份：改單價、品名、介紹，兩店同時生效。'],
    ['【品項｜各店欄位】隼使用台、鳥久使用台填站區（炒、炸、烤、冷、備品、調味，可多個，用「、」隔開）；留空＝該店不使用，網頁不會顯示。'],
    ['・平日／假日安全量各店分開填。新增品項時「品項ID」「廠商ID」留空即可，系統會依廠商名稱自動補上。'],
    ['【酒款｜各店欄位】隼使用、鳥久使用填 v 代表該店有賣；售價、杯價、進價兩店共用。'],
    ['【網頁修改】在網頁「設定」或「酒款資料管理」改資料會自動寫回這裡；在網頁刪除＝該店不再使用（共用資料保留，另一店不受影響）。'],
    ['・網頁和試算表同時修改時，以最後存檔的一方為準。'],
    ['【叫貨紀錄】一個品項一列。要刪打錯的單，請在網頁「紀錄」裡刪除；或在這裡刪掉同一單號的所有列。最後一欄「原始資料」（隱藏）請勿修改。'],
    ['【月統計】B1 輸入月份、B2 輸入門市（空白＝兩店合計）。估算金額＝叫貨當下單價 × 數量，實際以廠商帳單為準。']
  ];
  sh.getRange(1, 1, t.length, 1).setValues(t).setWrap(true);
  sh.setColumnWidth(1, 780);
}

/* ================= API ================= */
function doGet(e) { return handle_((e && e.parameter) || {}); }
function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: '格式錯誤' }); }
  return handle_(body);
}

function handle_(p) {
  try {
    if (String(p.key || '') !== SYNC_KEY || SYNC_KEY === '請改成你的同步密碼') return json_({ ok: false, error: 'auth' });
    if (p.action === 'ping') return json_({ ok: true, version: 3 });
    if (STORES.indexOf(p.store) < 0) return json_({ ok: false, error: '門市不正確' });
    switch (p.action) {
      case 'master': return json_(getMaster_(p.store));
      case 'saveMaster': return json_(withLock_(() => saveMaster_(p.store, p.vendors, p.items)));
      case 'wines': return json_(getWines_(p.store));
      case 'saveWines': return json_(withLock_(() => saveWines_(p.store, p.rows)));
      case 'addWines': return json_(withLock_(() => addWines_(p.rows, p.stores)));
      case 'history': return json_(getHistory_(p.store, p.kind, p.from, p.to));
      case 'saveOrder': return json_(withLock_(() => saveOrder_(p.store, p.kind, p.order)));
      case 'deleteOrder': return json_(withLock_(() => deleteOrder_(p.store, p.id)));
    }
    return json_({ ok: false, error: '未知的動作' });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function json_(o) { if (o && o.ok) { o.sid = SpreadsheetApp.getActive().getId(); o.v = 3; } return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function withLock_(fn) { const l = LockService.getScriptLock(); l.waitLock(25000); try { return fn(); } finally { l.releaseLock(); } }
function sheet_(k) { return SpreadsheetApp.getActive().getSheetByName(SH[k]) || ensureSheet_(k); }
function values_(k) { const sh = sheet_(k); const n = sh.getLastRow(); return n < 2 ? [] : sh.getRange(2, 1, n - 1, HEAD[k].length).getValues(); }
function writeAll_(k, rows) {
  const sh = sheet_(k), w = HEAD[k].length, n = sh.getLastRow();
  if (n > 1) sh.getRange(2, 1, n - 1, w).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, w).setValues(rows);
}
const num_ = v => (v === '' || v === null || v === undefined || isNaN(Number(v))) ? '' : Number(v);
const str_ = v => (v === null || v === undefined) ? '' : String(v);
const used_ = v => ['v', 'V', '是', 'y', 'Y', 'TRUE', 'true', '✓', '1'].indexOf(String(v).trim()) >= 0 || v === true;

/* ---------- 食材：廠商＋品項 ---------- */
function loadCatalog_() {
  const vrows = values_('vendors').filter(r => String(r[1]).trim());
  const irows = values_('items').filter(r => String(r[3]).trim());
  let changed = false;
  const byName = {};
  vrows.forEach(r => { if (!r[0]) { r[0] = Utilities.getUuid(); changed = true; } byName[String(r[1]).trim()] = String(r[0]); });
  irows.forEach(r => {
    if (!r[0]) { r[0] = Utilities.getUuid(); changed = true; }
    const vname = String(r[2]).trim();
    if (!r[1] || !vrows.some(v => String(v[0]) === String(r[1]))) {
      if (!byName[vname] && vname) { const id = Utilities.getUuid(); vrows.push([id, vname, '', '']); byName[vname] = id; }
      if (byName[vname]) { r[1] = byName[vname]; changed = true; }
    }
    const vn = vrows.find(v => String(v[0]) === String(r[1]));
    if (vn && String(r[2]) !== String(vn[1])) { r[2] = vn[1]; changed = true; }
  });
  if (changed) { writeAll_('vendors', vrows); writeAll_('items', irows); }
  return { vrows, irows };
}

function itemView_(r, store) {
  const c = ITEM_STORE_COL[store];
  return { id: str_(r[0]), vendor: str_(r[1]), name: str_(r[3]).trim(), cat: str_(r[c]).trim(), price: Number(num_(r[4]) || 0),
    unit: str_(r[5]), week: str_(r[c + 1]), holiday: str_(r[c + 2]), note: str_(r[ITEM_NOTE_COL[store]]) };
}

function getMaster_(store) {
  const { vrows, irows } = loadCatalog_();
  const c = ITEM_STORE_COL[store];
  const vendors = vrows.map(r => ({ id: str_(r[0]), name: str_(r[1]).trim(), days: str_(r[2]), note: str_(r[3]) }));
  const items = [], others = [];
  irows.forEach(r => (String(r[c]).trim() ? items : others).push(itemView_(r, store)));
  return { ok: true, vendors, items, others, stockTextSchema: 2 };
}

function saveMaster_(store, vendors, items) {
  if (!Array.isArray(vendors) || !Array.isArray(items)) return { ok: false, error: '資料不完整' };
  const { vrows, irows } = loadCatalog_();
  const c = ITEM_STORE_COL[store];
  const vById = {}; vrows.forEach(r => vById[String(r[0])] = r);
  vendors.forEach(v => {
    const row = vById[String(v.id)];
    const vals = [str_(v.id), str_(v.name).trim(), str_(v.days), str_(v.note)];
    if (row) vals.forEach((x, i) => row[i] = x); else { vrows.push(vals); vById[vals[0]] = vals; }
  });
  const iById = {}; irows.forEach(r => iById[String(r[0])] = r);
  const normItem = x => str_(x).toLowerCase().replace(/[\s　_\-－（）()]/g, '');
  const keep = new Set();
  items.forEach(i => {
    const id = str_(i.id);
    const vr = vById[str_(i.vendor)];
    const vendorName = vr ? str_(vr[1]) : '';
    // 同一商品兩店匯入時，即使早期分店 ID 不同，也優先對應共用的舊列
    let row = iById[id] || irows.find(r => normItem(r[3]) === normItem(i.name) &&
      (str_(r[1]) === str_(i.vendor) || (vendorName && normItem(r[2]) === normItem(vendorName))));
    if (!row) { row = Array(HEAD.items.length).fill(''); row[0] = id; irows.push(row); iById[id] = row; }
    keep.add(str_(row[0]));
    row[1] = str_(i.vendor); row[2] = vr ? vr[1] : row[2]; row[3] = str_(i.name).trim(); row[4] = num_(i.price); row[5] = str_(i.unit);
    row[c] = str_(i.cat).trim() || '備品'; row[c + 1] = str_(i.week); row[c + 2] = str_(i.holiday);
    row[ITEM_NOTE_COL[store]] = str_(i.note);
  });
  // 本店不再使用：清空本店欄位，共用資料保留
  irows.forEach(r => { if (!keep.has(String(r[0])) && String(r[c]).trim()) { r[c] = ''; r[c + 1] = ''; r[c + 2] = ''; r[ITEM_NOTE_COL[store]] = ''; } });
  writeAll_('vendors', vrows); writeAll_('items', irows);
  return getMaster_(store);
}

/* ---------- 酒款 ---------- */
function wineView_(r) {
  return { id: str_(r[0]), name: str_(r[1]), type: str_(r[2]) || '清酒', ml: num_(r[3]) === '' ? '' : Number(r[3]), price: num_(r[4]), cup: num_(r[5]),
    cost: num_(r[6]), supplier: str_(r[7]), sku: str_(r[8]), desc: str_(r[9]) };
}
function getWines_(store) {
  const c = WINE_STORE_COL[store];
  const rows = values_('wines').filter(r => String(r[1]).trim());
  let changed = false;
  rows.forEach(r => { if (!r[0]) { r[0] = Utilities.getUuid(); changed = true; } });
  if (changed) writeAll_('wines', rows);
  const items = [], others = [];
  rows.forEach(r => (used_(r[c]) ? items : others).push(wineView_(r)));
  return { ok: true, items, others };
}
function saveWines_(store, list) {
  if (!Array.isArray(list)) return { ok: false, error: '資料不完整' };
  const c = WINE_STORE_COL[store];
  const rows = values_('wines').filter(r => String(r[1]).trim());
  const byId = {}; rows.forEach(r => byId[String(r[0])] = r);
  const keep = new Set();
  list.forEach(w => {
    const id = str_(w.id); if (!id || !str_(w.name).trim()) return; keep.add(id);
    let row = byId[id];
    if (!row) { row = [id, '', '', '', '', '', '', '', '', '', '', '']; rows.push(row); byId[id] = row; }
    row[1] = str_(w.name); row[2] = str_(w.type); row[3] = num_(w.ml); row[4] = num_(w.price); row[5] = num_(w.cup);
    row[6] = num_(w.cost); row[7] = str_(w.supplier).trim(); row[8] = str_(w.sku); row[9] = str_(w.desc); row[c] = 'v';
  });
  rows.forEach(r => { if (!keep.has(String(r[0])) && used_(r[c])) r[c] = ''; });
  writeAll_('wines', rows);
  return { ok: true, count: keep.size };
}

/* 報價單匯入：新增或更新酒款，並標記指定門市（不會取消其他酒款的標記） */
function addWines_(list, stores) {
  if (!Array.isArray(list)) return { ok: false, error: '資料不完整' };
  stores = (Array.isArray(stores) ? stores : STORES).filter(s => WINE_STORE_COL[s] !== undefined);
  const rows = values_('wines').filter(r => String(r[1]).trim());
  const n = s => String(s || '').toLowerCase().replace(/[\s　_\-－（）()]/g, '');
  const byId = {}, byKey = {};
  rows.forEach(r => { byId[String(r[0])] = r; byKey[n(r[1]) + '|' + (Number(r[3]) || '')] = r; if (r[8]) byKey['sku:' + n(r[8])] = r; });
  let added = 0, updated = 0;
  list.forEach(w => {
    if (!str_(w.name).trim()) return;
    let row = byId[str_(w.id)] || (w.sku && byKey['sku:' + n(w.sku)]) || byKey[n(w.name) + '|' + (Number(w.ml) || '')];
    if (!row) { row = [str_(w.id) || Utilities.getUuid(), '', '', '', '', '', '', '', '', '', '', '']; rows.push(row); byId[row[0]] = row; added++; } else updated++;
    row[1] = str_(w.name); row[2] = str_(w.type) || row[2]; row[3] = num_(w.ml) === '' ? row[3] : num_(w.ml);
    if (num_(w.price) !== '' && Number(w.price) > 0) row[4] = num_(w.price);
    if (num_(w.cup) !== '' && Number(w.cup) > 0) row[5] = num_(w.cup);
    if (num_(w.cost) !== '') row[6] = num_(w.cost);
    ['', '', '', '', '', '', '', str_(w.supplier).trim(), str_(w.sku), str_(w.desc)].forEach((v, i) => { if (i >= 7 && v) row[i] = v; });
    stores.forEach(s => row[WINE_STORE_COL[s]] = 'v');
  });
  writeAll_('wines', rows);
  return { ok: true, added, updated };
}

/* 一次性整理：品項空白時匯入成本主表；既有酒款標成兩店都有 */
function oneTime_() {
  const props = PropertiesService.getScriptProperties();
  if (!values_('items').some(r => String(r[3]).trim())) seedItems();
  if (props.getProperty('winesBothStores') !== 'done') {
    const rows = values_('wines').filter(r => String(r[1]).trim());
    rows.forEach(r => { if (used_(r[WINE_STORE_COL['隼']]) || used_(r[WINE_STORE_COL['鳥久']])) { r[WINE_STORE_COL['隼']] = 'v'; r[WINE_STORE_COL['鳥久']] = 'v'; } });
    if (rows.length) writeAll_('wines', rows);
    props.setProperty('winesBothStores', 'done');
  }
}

/* ---------- 叫貨紀錄 ---------- */
function saveOrder_(store, kind, h) {
  if (!h || !h.id || !Array.isArray(h.lines) || !h.lines.length) return { ok: false, error: '資料不完整' };
  const sh = sheet_('log'), last = sh.getLastRow();
  if (last >= 2) {
    const ids = sh.getRange(2, 1, last - 1, 1).getValues();
    if (ids.some(r => String(r[0]) === String(h.id))) return { ok: true, dup: true };
  }
  const type = kind === 'wine' ? '酒類' : '食材';
  const time = h.date ? new Date(h.date) : new Date();
  const raw = JSON.stringify(h);
  if (raw.length > 49000) throw new Error('單筆紀錄過大');
  const rows = h.lines.map((l, idx) => {
    let name, qty, unit, price;
    if (kind === 'wine') { const d = l.d || {}; name = str_(d.name).replace(/\n/g, ' ') + (d.ml ? ' ' + d.ml + 'ml' : ''); qty = num_(l.qty); unit = '瓶'; price = num_(d.cost); }
    else { name = str_(l.name); qty = num_(l.qty); unit = str_(l.unit); price = num_(l.price); }
    const sub = (qty !== '' && price !== '') ? Math.round(qty * price) : '';
    return [str_(h.id), type, store, time, str_(h.delivery), str_(h.vendor), str_(h.who), name, qty, unit, price, sub, idx === 0 ? raw : ''];
  });
  sh.getRange(last + 1, 1, rows.length, HEAD.log.length).setValues(rows);
  return { ok: true, rows: rows.length };
}

function deleteOrder_(store, id) {
  const sh = sheet_('log'), n = sh.getLastRow();
  if (n < 2) return { ok: true, deleted: 0 };
  const vals = sh.getRange(2, 1, n - 1, 3).getValues();
  let del = 0;
  for (let i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][0]) === String(id) && vals[i][2] === store) { sh.deleteRow(i + 2); del++; }
  }
  return { ok: true, deleted: del };
}

function getHistory_(store, kind, from, to) {
  const type = kind === 'wine' ? '酒類' : '食材';
  const ok = s => /^\d{4}-\d{2}$/.test(s || '');
  const f = ok(from) ? from : '0000-00', t = ok(to) ? to : '9999-99';
  const orders = [];
  values_('log').forEach(r => {
    if (r[2] !== store || r[1] !== type || !r[12]) return;
    let h; try { h = JSON.parse(r[12]); } catch (e) { return; }
    const m = str_(h.date).slice(0, 7);
    if (m < f || m > t) return;
    orders.push(h);
  });
  return { ok: true, orders: orders };
}

/* ================= 選用：匯入成本主表品項（品項分頁是空的才會執行） ================= */
function seedItems() {
  if (values_('items').some(r => String(r[3]).trim())) throw new Error('品項分頁已經有資料，為避免重複，這次不匯入。');
  const vmap = {}, vrows = [], irows = [];
  SEED.vendors.forEach(v => { const id = Utilities.getUuid(); vmap[v[0]] = id; vrows.push([id, v[0], '', v[1]]); });
  SEED.items.forEach(i => irows.push([Utilities.getUuid(), vmap[i[0]], i[0], i[1], i[2], i[3], i[4], i[5], i[6], '', '', '']));
  writeAll_('vendors', vrows); writeAll_('items', irows);
  SpreadsheetApp.getActive().toast('已匯入 ' + vrows.length + ' 家廠商、' + irows.length + ' 個品項（隼的使用台已依成本主表填入；鳥久請自行勾填）');
}

/* 2026-09-07 成本主表；欄位：廠商, 品名, 單價, 單位, 隼使用台（空白＝隼未使用）, 隼平日, 隼假日 */
const SEED = {
 vendors: [["大業蔬菜", "每日送貨"], ["富帆-隼（含稅）", "生鮮漁貨空運三六到台灣，每日送貨"], ["富帆-鳥久（免稅）", "生鮮漁貨空運三六到台灣，每日送貨"], ["丸立鮮魚商", "每日送貨"], ["新永和魚商", "周一休市"], ["涅普頓有限公司（海鮮）", "周一休市"], ["鱻濤有限公司（海鮮）", "每日送貨"], ["億東", "每日送貨，請司機退冰"], ["裕賀", "每日送貨，請司機退冰"], ["美食家", "app下單，需指定日期，周日不送"], ["全國食品", ""], ["金媽媽泡菜", "周末不一定送"], ["海利貿易公司（海鮮）", "目前無合作"], ["惡海", "每日送貨"], ["JFC(TAIPEI)", "需提前預訂，2000元免運"], ["元榆雞-鳥久（免稅）", "叫貨後2-3天到，4000元免運，可少量叫貨到門市"], ["三商-鳥久", "每日送貨"], ["日本食研", "叫貨後2-3天到，1500元免運，冷凍、常溫分開計算"], ["雞肉大業", "周一休市"], ["勤益蛋", "三五送貨，需放現金支付"], ["錦品拉麵-鳥久（免稅）", "二五送貨"], ["米", ""], ["思橙(久利)", ""], ["勵碩企業股份有限公司（欣伯）", "需提前預訂，3000元免運，或2箱薯條"], ["雜支食材（免稅）", ""], ["鳥久", ""], ["隼", ""]],
 items: [
["大業蔬菜", "無籽檸檬", 170, "台斤", "烤", 1, 1],
["富帆-隼（含稅）", "蒟蒻-黑板（海苔蒟蒻）180gm", 30, "包", "備品", "", ""],
["富帆-鳥久（免稅）", "橙醋1800ml", 288, "罐", "冷、烤", 3, 3],
["丸立鮮魚商", "鮭魚卵(紅鱒)-日本神內無盒-醬油漬250g*2盒", 1050, "盒", "冷、炒", 1, 1],
["新永和魚商", "鮭魚卵(鱒)-日本(北食)無盒-醬油漬500/g", 1050, "盒", "冷、炒", 1, 1],
["涅普頓有限公司（海鮮）", "林屋海苔-半形（100張）", 950, "包", "冷", 1, 1],
["鱻濤有限公司（海鮮）", "鮭魚肚條", 370, "公斤", "烤", 2, 5],
["億東", "台灣豬梅花肉0.2袋裝", 235, "公斤", "炒", 1, 2],
["裕賀", "台灣豬五花火鍋片（1kg/10盒）", 320, "公斤", "備品", 1, 2],
["美食家", "母指粗粒黑胡椒600G/20入", 300, "包", "調味", 1, 1],
["金媽媽泡菜", "韓式泡菜", 350, "公斤", "備品", 1, 1],
["海利貿易公司（海鮮）", "山葵章魚1/kg", 767, "盒", "調味", 1, 2],
["惡海", "蛤蠣", 150, "台斤", "備品", 4, 6],
["JFC(TAIPEI)", "AMARI 一味唐辛粉 1KG", 560, "公斤", "調味", 1, 1],
["元榆雞-鳥久（免稅）", "雞皮 1kg", 80, "公斤", "備品", 3, 5],
["三商-鳥久", "丸金特級淡口醬油1.8L 6入/箱", 190, "罐", "調味", 1, 1],
["日本食研", "天婦羅粉1.5kg", 200, "包", "調味", 0.5, 0.5],
["雞肉大業", "去骨雞腿", 135, "台斤", "備品", 6, 9],
["勤益蛋", "勤憶雞蛋240顆/箱", 1920, "顆", "備品", 1, 1],
["錦品拉麵-鳥久（免稅）", "生拉麵120克/40包", 360, "箱", "備品", 1, 1],
["思橙(久利)", "山形正宗濃郁梅酒", 1700, "罐", "調味", 1, 1],
["勵碩企業股份有限公司（欣伯）", "3/8\"裏粉脆薯", 300, "包", "調味", 1, 1],
["雜支食材（免稅）", "魔杖麵包片", 65, "條", "備品", 1, 2],
["鳥久", "水", 0, "cc", "備品", "", ""],
["隼", "胡麻醬(隼）", 152.56777777777776, "公斤", "調味", "", ""],
["大業蔬菜", "台灣白Ｋ", 15, "台斤", "冷、炒、炸", 1, 1],
["富帆-隼（含稅）", "擺飾花（綜合）", 80, "盒", "備品", "", ""],
["富帆-鳥久（免稅）", "日本紫蘇大葉 100/枚（富凡）", 385, "盒", "冷", 1, 1],
["丸立鮮魚商", "竹葉（聯）100P裝/包", 160, "包", "冷、烤", 1, 1],
["新永和魚商", "午仔魚20P", 1800, "箱", "烤", 5, 5],
["億東", "台灣手修松阪豬", 480, "公斤", "烤、炒", 12, 12],
["裕賀", "美國特選翼板肉（冷凍）", 800, "公斤", "烤", 1, 2],
["美食家", "福泉味醂1.8L/6入", 230, "公升", "調味", 5, 5],
["元榆雞-鳥久（免稅）", "生鮮雞心 1kg", 250, "公斤", "備品", 1, 2],
["三商-鳥久", "大武郎清酒2000cc 紙盒/6盒/箱", 480, "罐", "調味", 2, 3],
["日本食研", "炸雞粉1.5kg", 250, "包", "調味", 0.5, 1],
["雞肉大業", "雞軟骨", 155, "台斤", "備品", 1.5, 2],
["雜支食材（免稅）", "白葡萄酒", 750, "箱", "調味", 1, 1],
["鳥久", "醃製乳酪", 267, "公斤", "備品", 1, 1],
["隼", "瑩烏賊醃製醬（隼）", 441.5, "支", "調味", "", ""],
["大業蔬菜", "進口白Ｋ", 20, "台斤", "冷、炒、炸", 1, 1],
["富帆-隼（含稅）", "乾燥朴葉20/入", 270, "包", "備品", "", ""],
["富帆-鳥久（免稅）", "(大/500gm)鹽昆布500gm/包 10包/箱", 550, "包", "冷", 1, 1],
["丸立鮮魚商", "山葵泥100%（第一分道）250g/包", 108, "包", "冷、烤", 1, 2],
["裕賀", "美國特選翼板牛（冷藏）", 860, "公斤", "烤", 1, 2],
["美食家", "味全中華調味醬647G/12入", 83, "罐", "調味", 2, 2],
["元榆雞-鳥久（免稅）", "母雞二節翅 1kg(45-60g)", 200, "公斤", "備品", 1.5, 2],
["日本食研", "冷凍炸牡蠣Ｎ", 330, "包", "備品", 1, 2],
["雞肉大業", "雞佛", 650, "台斤", "備品", 0.5, 0.5],
["鳥久", "和風醬（鳥久）", 92.4201194181523, "公斤", "調味", 1, 1],
["隼", "雞肝醃製（隼）", 350.8733333333333, "公斤", "備品", "", ""],
["大業蔬菜", "進口洋蔥（20斤）", 20, "台斤", "冷、炒", 1, 1],
["富帆-隼（含稅）", "活生蠔（熊本品種）約120g", 115, "顆", "備品", "", ""],
["富帆-鳥久（免稅）", "活生蠔（熊本）", 115, "顆", "冷", 8, 10],
["丸立鮮魚商", "青甘片（日本）", 920, "公斤", "冷", 1, 1],
["裕賀", "美國條肉", 480, "公斤", "烤", 20, 20],
["美食家", "味全義大利調味醬605G/12入", 83, "罐", "調味", 2, 2],
["元榆雞-鳥久（免稅）", "去骨雞屁股", 450, "公斤", "備品", 1, 2],
["雞肉大業", "二節翅", 120, "台斤", "備品", 1, 1.5],
["鳥久", "毛豆（鳥久）", 99.39, "公斤", "備品", 1, 1],
["大業蔬菜", "中華豆腐", 14, "盒", "冷、炒", 1, 1],
["富帆-隼（含稅）", "冰鮮明蝦(100-110g)/隻", 1600, "公斤", "備品", "", ""],
["富帆-鳥久（免稅）", "紅薑絲1kg/包", 160, "包", "炒", 1, 1],
["丸立鮮魚商", "鮭魚(帶頭)78鮭。整尾", 520, "公斤", "冷", 8, 8],
["美食家", "Q比胡麻醬1L/9入", 230, "罐", "調味", 3, 3],
["鳥久", "大根（鳥久）", 251.86666666666667, "公斤", "備品", 1, 1],
["大業蔬菜", "原味豆干", 33, "台斤", "冷、炒", 2, 2],
["富帆-隼（含稅）", "黑潮冷凍軟殼龍蝦180-200g/隻", 777, "支", "備品", "", ""],
["富帆-鳥久（免稅）", "乾昆布600g", 480, "包", "炒", 1, 1],
["丸立鮮魚商", "鮪魚(中上)B", 570, "公斤", "冷", 2, 2],
["美食家", "鉅利冰梅醬3KG/4入", 510, "罐", "調味", 1, 1],
["鳥久", "冰梅醬（鳥久）", 273.8888888888889, "公斤", "調味", 1, 1],
["大業蔬菜", "大蒜仁", 95, "台斤", "冷、炒", 1, 1],
["富帆-隼（含稅）", "鱈魚子1/kg", 1550, "公斤", "備品", "", ""],
["富帆-鳥久（免稅）", "炒麵醬汁1200/gm", 175, "罐", "炒", 1, 2],
["丸立鮮魚商", "日本紫蘇大葉 100/枚（丸力）", 310, "盒", "冷", 1, 1],
["美食家", "S&B柚子青辣椒醬43G/100入", 62, "罐", "調味", 1, 1],
["鳥久", "酒鬼底料（鳥久）", 261.86333333333334, "公斤", "調味", 1, 1],
["大業蔬菜", "大辣椒", 120, "台斤", "冷、炒", 1, 1],
["富帆-隼（含稅）", "冷凍青海苔（日本產）150gm/盒", 110, "盒", "備品", "", ""],
["富帆-鳥久（免稅）", "中濃醬 多福 2.1/kg", 275, "罐", "炸", 1, 1],
["丸立鮮魚商", "(3S)冷凍生食干貝約41-50/1kg/盒（丸力）", 1790, "盒", "冷", 1, 1],
["美食家", "工研酢(業務用)5L/4入", 220, "桶", "調味", 1, 1],
["鳥久", "名古屋醬汁（10倍）", 140.97222222222223, "公斤", "調味", 1, 1],
["大業蔬菜", "蒜泥", 90, "台斤", "冷、烤、炸", 1, 1],
["富帆-隼（含稅）", "鰹魚酒盜280gm/瓶", 700, "罐", "調味", "", ""],
["富帆-鳥久（免稅）", "冰鮮班節蝦（100-110g)隻", 1150, "公斤", "烤", 2, 2],
["丸立鮮魚商", "松葉蟹膏（日本越河）300g/包", 440, "包", "冷", 1, 2],
["美食家", "工研業務烏醋5L/4入", 220, "桶", "調味", 1, 1],
["鳥久", "烤鹽（鳥久）", 185, "公斤", "調味", 1, 1],
["大業蔬菜", "薑泥", 85, "台斤", "冷、烤、炸", 0.5, 0.5],
["富帆-隼（含稅）", "松葉蟹碎肉黑盒大（400-500gm)生鮮", 1040, "盒", "備品", "", ""],
["富帆-鳥久（免稅）", "鱈魚子1/kg", 1550, "公斤", "烤", 1, 1],
["丸立鮮魚商", "老船長鱈魚肝120g/盒", 75, "罐", "冷", 7, 10],
["美食家", "菊鶴信州味噌(非基改)500G/8入(藏)", 60, "盒", "調味", 4, 4],
["鳥久", "辣味增（鳥久）", 247.81, "公斤", "備品", 1, 1],
["大業蔬菜", "子薑", 100, "台斤", "冷、烤、炸", 0.5, 0.5],
["富帆-隼（含稅）", "蒲燒鰻20-25隻/10kg/箱", 780, "公斤", "備品", "", ""],
["富帆-鳥久（免稅）", "煉海膽醬-煉1/kg", 2200, "包", "烤", 1, 1],
["丸立鮮魚商", "冷凍鮭魚8/9(智利）", 355, "公斤", "炒", 17, 17],
["美食家", "福松脫皮辣450G/12入", 80, "罐", "調味", 12, 12],
["鳥久", "剁椒醬（鳥久）", 271.12666666666667, "公斤", "調味", 1, 1],
["大業蔬菜", "紫洋蔥", 40, "台斤", "冷", 2, 3],
["富帆-隼（含稅）", "冷藏/煮湯味增/日本神州/1kg/包", 130, "包", "調味", "", ""],
["丸立鮮魚商", "蟹味棒（力二味）500G-50P", 170, "包", "炒", 1, 2],
["美食家", "小磨坊七味唐辛子300G/20入", 125, "包", "調味", 2, 2],
["鳥久", "辣奶油（鳥久）", 105.24603174603175, "公斤", "調味", 1, 1],
["大業蔬菜", "過貓（1斤）", 100, "台斤", "冷", 1, 2],
["富帆-隼（含稅）", "彩膳胡麻450/g", 750, "罐", "調味", "", ""],
["丸立鮮魚商", "L戎の魚店鯖魚片200g±30g/片19-21片", 1660, "箱", "烤", 1, 1],
["美食家", "大橋元氣一番好米5KG", 285, "包", "備品", 4, 4],
["鳥久", "鮭魚鬆（鳥久）", "", "", "備品", 1, 1],
["大業蔬菜", "香菜", 100, "台斤", "冷", 0.5, 1],
["富帆-隼（含稅）", "茶泡飯用（米果）300gm/包", 250, "包", "備品", "", ""],
["丸立鮮魚商", "明太子(雙葉)500g/包", 420, "包", "烤", 1, 1],
["美食家", "福祿壽美食料理米酒600ML/24入", 25, "罐", "調味", 24, 24],
["鳥久", "味噌湯底（鳥久）", 29.766666666666666, "公斤", "調味", 1, 1],
["大業蔬菜", "小黃瓜", 70, "台斤", "冷", 1, 2],
["富帆-隼（含稅）", "(綠色/小罐)山椒粉(朝倉/富士屋)", 850, "罐", "調味", "", ""],
["美食家", "美食家大豆沙拉油2,6L/6入", 155, "罐", "調味", 4, 4],
["鳥久", "炸雞漿（鳥久）", 116.66666666666666, "cc", "備品", 1, 1],
["大業蔬菜", "牛蕃茄", 60, "台斤", "冷", 1, 1.5],
["富帆-隼（含稅）", "胡麻調味醬（白）900gm/12罐/箱", 800, "罐", "調味", "", ""],
["美食家", "福壽耐炸油18L", 880, "桶", "調味", 2, 2],
["鳥久", "天婦羅漿（鳥久）", 80, "cc", "備品", 1, 1],
["大業蔬菜", "萊姆", 150, "台斤", "冷", 0.5, 0.5],
["富帆-隼（含稅）", "宇治抹茶粉（向春園）300gm", 770, "包", "調味", "", ""],
["美食家", "福壽香油3L/6入", 275, "罐", "調味", 1, 1],
["鳥久", "天汁醬（鳥久）", 65.17166666666667, "cc", "調味", 1, 1],
["大業蔬菜", "粉蔥", 40, "台斤", "冷", 0.5, 0.5],
["富帆-隼（含稅）", "(晶凍）柚子醋醬汁300gm/包", 140, "包", "調味", "", ""],
["美食家", "福壽福麻油3L/6入", 550, "罐", "調味", 1, 1],
["鳥久", "牛蒡粉（鳥久）", 215.36666666666667, "公斤", "調味", 1, 1],
["大業蔬菜", "北蔥", 35, "台斤", "冷", 0.5, 0.5],
["富帆-隼（含稅）", "夏季黑松露醬（5%松露）500gm", 640, "罐", "調味", "", ""],
["美食家", "福壽辣椒油3L/6入", 260, "罐", "調味", 1, 1],
["鳥久", "明太子麵包醬（鳥久）", 609.0833333333333, "公斤", "調味", 1, 1],
["大業蔬菜", "青龍", 120, "台斤", "冷", 0.5, 0.5],
["富帆-隼（含稅）", "（東）白醬油1.8公升", 285, "罐", "調味", "", ""],
["美食家", "低筋麵粉(義峰)3KG/8入", 105, "包", "調味", 1, 1],
["鳥久", "烤醬（鳥久）", 183.8111111111111, "公斤", "調味", 1, 1],
["大業蔬菜", "吻仔魚", 250, "台斤", "冷", 1, 1],
["富帆-隼（含稅）", "鹽麴500gm/包 20包/箱", 165, "包", "調味", "", ""],
["美食家", "樹薯（粗）澱粉3KG", 140, "包", "調味", 1, 1],
["鳥久", "雞肉丸半成品（鳥久）", 3182.291666666666, "公斤", "備品", 1, 1],
["大業蔬菜", "皮蛋", 13, "顆", "冷", 10, 15],
["富帆-隼（含稅）", "三溫糖1kg/包", 140, "公斤", "調味", "", ""],
["美食家", "樹薯（細）澱粉3KG", 130, "包", "調味", 1, 1],
["鳥久", "特調鰹魚醬油（鳥久）", 22.15, "cc", "調味", 1, 1],
["大業蔬菜", "乾辣椒絲（大業）", 240, "包", "冷", 1, 1],
["美食家", "燦和生鮮麵包粉（F白-4/41)1KG", 100, "包", "調味", 3, 3],
["鳥久", "明太子醬（鳥久）", 696.1111111111111, "公斤", "調味", 1, 1],
["大業蔬菜", "雪翠高山高麗菜", 20, "台斤", "炒、炸", 1, 2],
["美食家", "4兩包式耐熱袋400G/60入", 35, "包", "備品", 3, 3],
["鳥久", "鹹豬肉（鳥久）", 1845.5920247933886, "公斤", "備品", 1, 1],
["大業蔬菜", "平地高麗菜", 15, "台斤", "炒、炸", 1, 2],
["美食家", "1斤捲式耐熱袋5LBS/12入", 200, "捲", "備品", 2, 2],
["鳥久", "橙醋（鳥久）", 506.8166666666666, "cc", "調味", 1, 1],
["大業蔬菜", "牛蒡", 40, "支", "炒、炸", 3, 4],
["美食家", "3斤捲式耐熱袋5LBS/12入", 200, "捲", "備品", 1, 1],
["鳥久", "午仔魚（鳥久）", 11497.13, "支", "備品", 1, 1],
["大業蔬菜", "日本淮山", 150, "台斤", "炒、烤", 2, 3],
["美食家", "半斤捲式耐熱袋5LBS/12入", 190, "捲", "備品", 2, 2],
["鳥久", "飯糰（鳥久）", 68.9, "公斤", "備品", 1, 1],
["大業蔬菜", "水蓮（4兩）", 25, "包", "炒", 4, 6],
["美食家", "5斤捲式耐熱袋", 200, "捲", "備品", 1, 1],
["鳥久", "海膽醬（鳥久）", 1334, "公斤", "調味", 1, 1],
["大業蔬菜", "紅Ｋ", 17, "台斤", "炒", 1, 1],
["美食家", "湯杯蓋520(50只）/20入", 35, "條", "備品", 1, 1],
["大業蔬菜", "蓮藕", 70, "台斤", "炒", 1, 1],
["美食家", "空白紙湯杯520(50只）/20入", 80, "條", "備品", 1, 1],
["大業蔬菜", "剝皮玉米", 35, "台斤", "炒", 1, 1],
["美食家", "南亞家用保鮮膜(30CM*200尺)/60入", 55, "支", "備品", 4, 4],
["大業蔬菜", "吉頭菜", 30, "顆", "炒", 1, 1],
["美食家", "竹叉6吋/20入", 85, "包", "備品", 2, 2],
["大業蔬菜", "麻竹筍", 95, "台斤", "炒", 2, 2],
["美食家", "3M菜瓜布(有海綿/2入)/50入", 38, "包", "備品", 2, 2],
["大業蔬菜", "桂竹筍", 80, "台斤", "炒", 1, 1],
["美食家", "白毛巾（12條）/50入", 210, "包", "備品", 1, 1],
["大業蔬菜", "山苦瓜", 90, "台斤", "炒", 1.5, 2],
["美食家", "鰹魚烹大師1KG/12入", 310, "包", "冷、炒、炸", 1, 1],
["大業蔬菜", "白苦瓜", 75, "台斤", "炒", 1, 1],
["美食家", "台糖精製細砂1KG/20入", 38, "包", "冷、炒、烤、炸", 2, 3],
["大業蔬菜", "冬瓜(切片)", 22, "台斤", "炒", 2, 2],
["美食家", "小魚乾600G/50入（凍）", 140, "包", "冷、炒", 1, 1],
["大業蔬菜", "老薑", 100, "台斤", "炒", 1, 1],
["美食家", "味豐味素1KG/12入", 80, "包", "冷、炒", 1, 1],
["大業蔬菜", "小辣椒", 80, "台斤", "炒", 1, 1],
["美食家", "三魚嚴選柴魚絲3兩裝80G/40入", 105, "包", "冷、炒", 6, 6],
["大業蔬菜", "九層塔", 85, "台斤", "炒", 0.2, 0.2],
["美食家", "金蘭純釀造醬油（非基改）5L/4入", 320, "桶", "冷、烤、炒、炸", 1, 1],
["大業蔬菜", "美白菇100g", 35, "包", "炒", 1, 1],
["美食家", "小磨坊傳香特級胡椒粉T(#3)600G/12入", 100, "盒", "冷、烤、炒", 2, 2],
["大業蔬菜", "鴻禧菇100g", 35, "包", "炒", 1, 1],
["美食家", "士平熟白芝麻粒(真空包)600G/30入", 110, "包", "冷、烤、炒", 2, 2],
["大業蔬菜", "真空金針菇200g", 15, "包", "炒", 2, 3],
["美食家", "味滋康鰹魚調味露(2倍)1L/12入", 155, "公升", "冷、烤、炒", 3, 3],
["大業蔬菜", "木耳", 95, "台斤", "炒", 0.5, 0.5],
["美食家", "竹葉100張/30入", 230, "包", "冷、烤", 1, 1],
["大業蔬菜", "大茂黑瓜", 40, "罐", "炒", 2, 2],
["美食家", "雅瑪珂1/2切海苔（梅）50/枚", 250, "包", "冷、烤", 2, 2],
["大業蔬菜", "酸菜仁", 40, "台斤", "炒", 0.5, 0.5],
["美食家", "毛豆莢1KG/10入(凍)", 70, "包", "冷", 3, 5],
["大業蔬菜", "枸杞（大）", 280, "台斤", "炒", 1, 1],
["美食家", "民生特製蔭鼓（黑豆）600G/40入", 40, "包", "冷", 1, 1],
["大業蔬菜", "紅棗（大）", 90, "台斤", "炒", 1, 1],
["美食家", "母指八角粒600G/15入", 280, "包", "冷", 1, 1],
["大業蔬菜", "芋頭", 48, "台斤", "炸", 1, 1.5],
["美食家", "小磨坊花椒粒500G/20入", 500, "包", "冷", 1, 1],
["大業蔬菜", "雞蛋豆腐", 30, "盒", "炸", 3, 5],
["美食家", "野生甘草片600G/30入", 270, "包", "冷", 1, 1],
["大業蔬菜", "中香菇", 130, "台斤", "烤、炒", 1, 2],
["美食家", "海山原色話梅500G/36入", 150, "包", "冷", 1, 1],
["大業蔬菜", "大沙拉（500g)", 55, "包", "烤、炸", 3, 3],
["美食家", "元氣柚子皮100G/30入（凍）", 260, "包", "冷", 1, 1],
["大業蔬菜", "帶殼皎白筍", 110, "台斤", "烤", 2, 3],
["美食家", "Sungold鮮奶油乳酪2KG/6入", 500, "盒", "冷", 1, 1],
["大業蔬菜", "帶皮玉米筍（12兩）", 40, "包", "烤", 6, 8],
["美食家", "海山原色話梅500G/36入", 150, "包", "冷", 1, 1],
["大業蔬菜", "綠節瓜", 80, "台斤", "烤", 1, 2],
["美食家", "鉅利冰梅醬3KG/4入", 510, "罐", "冷", 1, 1],
["大業蔬菜", "五花肉", 180, "台斤", "烤", 20, 20],
["美食家", "士平蒜蓉花生(真空包)600G/30入", 175, "包", "冷", 2, 2],
["美食家", "金蘭珍品油膏(非基改)5L/4入", 245, "桶", "冷", 1, 1],
["美食家", "EBARA烤鰻魚醬1600G/6入", 330, "公升", "冷", 2, 2],
["美食家", "台鹽高級碘鹽(含碘)1KG/24入", 17, "包", "炒、烤", 1, 1],
["美食家", "元氣蝦米600G/30入（凍）", 140, "包", "炒", 1, 1],
["美食家", "錦油蔥酥（紅）600G/50入", 60, "包", "炒", 1, 2],
["美食家", "紅棗600G/30入", 85, "包", "炒", 1, 1],
["美食家", "讚岐烏龍麵180G*5/6入(凍)", 100, "包", "炒", 12, 12],
["美食家", "康寶雞粉1KG/6入", 280, "罐", "炒", 1, 1],
["美食家", "元氣燒海苔細絲100G/40入", 135, "包", "炒", 1, 1],
["美食家", "橘平屋海苔巧絲100G/40入", 135, "包", "炒", 1, 1],
["美食家", "乾海帶芽600G/24入", 250, "包", "炒", 1, 1],
["美食家", "士平熟黃豆粉600G/30入", 65, "包", "炸", 1, 1],
["美食家", "小磨坊孜然粉(馬芹子粉Cumin)1KG/12入", 490, "包", "炸", 1, 1],
["美食家", "飛燕煉乳1.3KG/12入", 140, "罐", "炸", 1, 1],
["美食家", "士平熟黑芝麻粒(真空包)600G/30入", 130, "包", "炸", 1, 1],
["美食家", "可果美番茄醬(軟管)500G/12入", 68, "罐", "炸", 2, 2],
["美食家", "青海苔粉300G/12入", 330, "包", "烤、炒、炸", 1, 1],
["美食家", "味島香鬆鰹節香味500G/25入", 455, "包", "烤、炒", 2, 2],
["美食家", "松井芥末椒鹽600G/20入", 190, "盒", "烤、炸", 2, 2],
["美食家", "小磨坊香辣椒鹽粉T(#1)600G/12入", 100, "盒", "烤、炸", 2, 2],
["美食家", "去肚魷魚", 3300, "箱", "烤", 1, 1],
["美食家", "永良冰糖（細）3KG/8入", 165, "包", "烤", 1, 1],
["美食家", "新光玫瑰鹽(細鹽)1KG/20入", 165, "包", "烤", 2, 2],
["美食家", "美味深礦鹽(岩鹽/無碘)1KG/20入", 20, "包", "烤", 2, 2],
["大業蔬菜", "山東白菜", 28, "台斤", "", "", ""],
["大業蔬菜", "馬K", 28, "台斤", "", "", ""],
["大業蔬菜", "紅皮馬Ｋ", 110, "台斤", "", "", ""],
["大業蔬菜", "馬吉", 90, "台斤", "", "", ""],
["大業蔬菜", "茄子", 35, "台斤", "", "", ""],
["大業蔬菜", "日本茄子（米茄）", 100, "台斤", "", "", ""],
["大業蔬菜", "玉米筍（100g)", 30, "盒", "", "", ""],
["大業蔬菜", "進口洋蔥", 24, "台斤", "", "", ""],
["大業蔬菜", "秋葵", 160, "台斤", "", "", ""],
["大業蔬菜", "奶油白菜", 40, "台斤", "", "", ""],
["大業蔬菜", "京都水菜", 65, "包", "", "", ""],
["大業蔬菜", "菠菜（把）", 25, "把", "", "", ""],
["大業蔬菜", "空心菜（把）", 30, "把", "", "", ""],
["大業蔬菜", "小松菜", 60, "台斤", "", "", ""],
["大業蔬菜", "台灣青花菜", 40, "台斤", "", "", ""],
["大業蔬菜", "羽衣甘藍", 190, "台斤", "", "", ""],
["大業蔬菜", "綠捲鬚", 190, "台斤", "", "", ""],
["大業蔬菜", "紅捲鬚", 200, "台斤", "", "", ""],
["大業蔬菜", "生豆包", 70, "台斤", "", "", ""],
["大業蔬菜", "魚豆腐", 110, "台斤", "", "", ""],
["大業蔬菜", "小蘋果", 20, "公斤", "", "", ""],
["大業蔬菜", "小蕃茄", 75, "台斤", "", "", ""],
["大業蔬菜", "金桔", 90, "台斤", "", "", ""],
["大業蔬菜", "蒜末", 90, "台斤", "", "", ""],
["大業蔬菜", "秀珍菇", 95, "台斤", "", "", ""],
["大業蔬菜", "杏鮑菇Ａ（1kg)", 120, "包", "", "", ""],
["大業蔬菜", "洋菇", 70, "盒", "", "", ""],
["大業蔬菜", "雲耳", 60, "台斤", "", "", ""],
["大業蔬菜", "梅花肉", 170, "台斤", "", "", ""],
["大業蔬菜", "雞蛋（白蛋,盒）", 65, "顆", "", "", ""],
["大業蔬菜", "鳥蛋", 140, "台斤", "", "", ""],
["大業蔬菜", "脆瓜", 65, "桶", "", "", ""],
["大業蔬菜", "老乾媽辣椒210/g", 125, "罐", "", "", ""],
["富帆-鳥久（免稅）", "海藻沙拉-五色乾海藻100g", 360, "包", "", "", ""],
["富帆-鳥久（免稅）", "(3S)冷凍生食干貝約41-50/1kg/盒（富凡）", 1700, "盒", "", "", ""],
["丸立鮮魚商", "真鯛（日本）", 570, "公斤", "", "", ""],
["丸立鮮魚商", "迦納（澎湖）2K上", 580, "公斤", "", "", ""],
["丸立鮮魚商", "香魚6P（母）950g/盒", 450, "盒", "", "", ""],
["丸立鮮魚商", "熟凍瑩烏賊(日本兵庫)80-90隻裝 300克/盒", "", "", "", "", ""],
["丸立鮮魚商", "冰鮮瑩烏賊（熟）日本300g/盒", 430, "盒", "", "", ""],
["丸立鮮魚商", "松葉蟹碎肉（佳）500g/袋裝", 590, "包", "", "", ""],
["丸立鮮魚商", "鮭魚鬆(日本製)1kg/包", 670, "包", "", "", ""],
["丸立鮮魚商", "鮭魚卵(鱒)-日本丸和無盒-醬油漬250g*2盒", 1300, "盒", "", "", ""],
["丸立鮮魚商", "鮭魚卵(秋)-日本極洋無盒-醬油漬500g", 1550, "盒", "", "", ""],
["裕賀", "美國特選無骨牛小排", 1050, "公斤", "", "", ""],
["裕賀", "美國冷凍外側橫隔膜", 320, "公斤", "", "", ""],
["裕賀", "紐西蘭牛舌", 550, "公斤", "", "", ""],
["裕賀", "紐西蘭牛舌尖(FF)3k-牛角", 100, "公斤", "", "", ""],
["裕賀", "美國帶油胸腹肉火鍋片", 280, "公斤", "", "", ""],
["裕賀", "安格斯牛腹肉片（500g/盒）", 280, "公斤", "", "", ""],
["裕賀", "澳洲冷凍和牛大腸", 600, "公斤", "", "", ""],
["美食家", "允億金梅粉600G/30入", 130, "包", "", "", ""],
["美食家", "元氣傳統炸蒜片500G/25入", 800, "包", "", "", ""],
["美食家", "卜蜂原味溏心蛋(10粒)/20入(藏)", 130, "包", "", "", ""],
["美食家", "日式醬香糖心蛋(10粒)/20入(藏)", 125, "包", "", "", ""],
["美食家", "Simplot帶皮楔形-薯條(原8切)2.27KG/6入(凍)", 255, "包", "", "", ""],
["美食家", "Cavcndish3/8脆薯(65090白)2KG/6入", 235, "包", "", "", ""],
["美食家", "辛妮迪瑪斯卡碰500G/6入", 205, "罐", "", "", ""],
["美食家", "元氣新西蘭奶油（無鹽）1LB/24入（凍）", 190, "包", "", "", ""],
["美食家", "安佳鮮奶油1L", 225, "罐", "", "", ""],
["美食家", "TATUA噴式鮮奶油400G", 190, "罐", "", "", ""],
["美食家", "華冠濃醇乳香雙色乳酪絲(綠膠)1KG/12入", 260, "包", "", "", ""],
["美食家", "TATUA切達起司醬1KG", 235, "包", "", "", ""],
["美食家", "寶譽粘書師父起司脆酥粉1KG/12入", 80, "包", "", "", ""],
["美食家", "老公仔標A五香粉600G/30入", 210, "盒", "", "", ""],
["美食家", "台糖二號砂糖1KG/20入", 38, "包", "", "", ""],
["美食家", "沖繩黑糖蜜500G(350ML)/12入", 90, "罐", "", "", ""],
["美食家", "Niscko一加侖六兵衛抹茶冰淇淋/3入(凍)", 435, "桶", "", "", ""],
["美食家", "小美特級黃桶香草冰淇淋3GAL(凍）", 450, "桶", "", "", ""],
["美食家", "老公仔標胡椒鹽粉（棕盒）600g/30入", 210, "盒", "", "", ""],
["美食家", "小磨坊粗黑胡椒(#71)600G/20入", 300, "盒", "", "", ""],
["美食家", "三魚嚴選柴魚絲1兩裝(36G*4)/20入", 190, "包", "", "", ""],
["美食家", "EBARA-EB豚骨拉麵湯頭1,8L/6入", 540, "罐", "", "", ""],
["美食家", "十全辣豆瓣醬(非基改)2800G/6入", 140, "罐", "", "", ""],
["美食家", "瓜哇咖哩辣味185G", 95, "盒", "", "", ""],
["美食家", "Hana日式厚麻糬(1.6*4*6cm*8塊)/12入", 135, "包", "", "", ""],
["美食家", "大橘越光米5KG/6入", 400, "包", "", "", ""],
["美食家", "福壽健味香油3L/6入", 255, "罐", "", "", ""],
["美食家", "6兩包式耐熱袋400G/60入", 36, "包", "", "", ""],
["美食家", "好潔鋁箔紙(30CM*75尺)/24入", 115, "支", "", "", ""],
["美食家", "小熊球刷40G/480入", 15, "包", "", "", ""],
["全國食品", "最中小丸菊（花形糯米殼） 204入", 1000, "箱", "", "", ""],
["JFC(TAIPEI)", "TOUKON乾燥海藻100G", 235, "克", "", "", ""],
["JFC(TAIPEI)", "DEMER 醋漬鯖魚片 110G", 160, "克", "", "", ""],
["JFC(TAIPEI)", "S&B七味粉300G chili powder", 133, "克", "", "", ""],
["JFC(TAIPEI)", "飛馬 條和蜂蜜糖漿3KG", 267, "克", "", "", ""],
["JFC(TAIPEI)", "晨欣 水煮蓮藕(0.8CM) 550G Renkon", 235, "克", "", "", ""],
["JFC(TAIPEI)", "阿根廷魷魚一夜干 200G", 135, "克", "", "", ""],
["JFC(TAIPEI)", "OBAKO 煙燻蘿蔔 200G", 156, "克", "", "", ""],
["元榆雞-鳥久（免稅）", "雞頸肉 500g", 550, "公斤", "", "", ""],
["元榆雞-鳥久（免稅）", "生鮮雞肝300G", 55, "克", "", "", ""],
["元榆雞-鳥久（免稅）", "三節翅 1kg", 200, "公斤", "", "", ""],
["元榆雞-鳥久（免稅）", "雞牡蠣 500g", 550, "公斤", "", "", ""],
["三商-鳥久", "九鬼 純正山七胡麻油600G 10瓶/箱", 360, "罐", "", "", ""],
["三商-鳥久", "丸金特級濃口醬油1.8L 6入/箱", 190, "罐", "", "", ""],
["三商-鳥久", "KEWPIE美乃滋 450G", 100, "罐", "", "", ""],
["三商-鳥久", "丹麥康門貝爾乾酪125G/12入", 120, "盒", "", "", ""],
["三商-鳥久", "蒜香風味蝦仁300G/20入", 220, "包", "", "", ""],
["三商-鳥久", "MARUDEN冷凍北海道唐揚章魚500G", 385, "包", "", "", ""],
["三商-鳥久", "KEWPIE SNOW MAN 料亭風玉子燒300G/包", 150, "包", "", "", ""],
["日本食研", "冷凍蝦排60g", 1425, "箱", "", "", ""],
["米", "瑞昌米行", 780, "台斤", "", "", ""]
 ]
};
