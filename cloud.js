/* 鳥久 × 隼｜雲端同步模組（v3）
 * - 廠商、品項、酒款：兩店共用一份目錄（Google 試算表），每筆標記各店是否使用
 * - 食材叫貨紀錄、酒類叫貨紀錄：各店分開
 * 原本的功能程式只做了少量修改：資料寫入手機後自動上傳，進入店別時自動下載。
 */
(function () {
  const CLOUD_URL = 'https://script.google.com/macros/s/AKfycbwvcq-_dXucjans6i0FyDOvRaKd1kyEkDIHHdlql4lg0uEFQio3LqsJNlduKM8sJ3eM/exec'; // Apps Script 網頁應用程式網址
  const API_VERSION = 3;
  const STORES = ['隼', '鳥久'];

  const LS = window.localStorage;
  const rawSet = Storage.prototype.setItem;
  const rawGet = Storage.prototype.getItem;
  const KEYS = { food: 'ordersV2:', wine: 'wineOrdersV3:', wines: 'wineLabelV1' };
  let applying = false, needKeyAsked = false, lastOk = null, lastErr = '';
  const chains = {}, timers = {}, fetchedMonths = {};

  /* ---------- 小工具 ---------- */
  const get = k => { try { return JSON.parse(rawGet.call(LS, k)); } catch (e) { return null; } };
  const put = (k, v) => { try { applying = true; rawSet.call(LS, k, JSON.stringify(v)); } catch (e) { } finally { applying = false; } };
  const getRaw = k => rawGet.call(LS, k);
  const putRaw = (k, v) => { try { applying = true; rawSet.call(LS, k, v); } catch (e) { } finally { applying = false; } };
  const key = () => getRaw('cloud:key') || '';
  const uuid = () => (window.crypto && crypto.randomUUID) ? crypto.randomUUID() : 'w' + Date.now().toString(36) + Math.random().toString(36).slice(2, 10);
  const norm = s => String(s || '').toLowerCase().replace(/[\s　_\-－（）()]/g, '');
  const monthsAgo = n => { const d = new Date(); d.setUTCDate(1); d.setUTCMonth(d.getUTCMonth() - n); return d.toISOString().slice(0, 7); };
  const hash = s => { let h = 5381; for (let i = 0; i < s.length; i++) h = ((h << 5) + h + s.charCodeAt(i)) | 0; return String(h >>> 0); };
  const masterHash = o => hash(JSON.stringify({ v: o.vendors || [], i: o.items || [] }));
  const wineHash = rows => hash(JSON.stringify(rows || []));
  const readFood = s => get(KEYS.food + s) || { vendors: [], items: [], history: [] };
  const readList = (kind, s) => kind === 'food' ? (readFood(s).history || []) : (get(KEYS.wine + s) || []);
  const syncedSet = (kind, s) => new Set(get('cloud:synced:' + kind + ':' + s) || []);
  const saveSynced = (kind, s, set) => put('cloud:synced:' + kind + ':' + s, [...set]);
  const isCurrent = s => typeof current !== 'undefined' && current === s;
  function serial(name, fn) { chains[name] = (chains[name] || Promise.resolve()).then(fn, fn); return chains[name]; }

  function writeList(kind, s, list) {
    if (kind === 'food') { const o = readFood(s); o.history = list; put(KEYS.food + s, o); syncMemory(s); }
    else put(KEYS.wine + s, list);
  }
  // 寫入手機後立刻更新畫面用的資料，避免之後以舊資料覆蓋
  function syncMemory(s) {
    if (!isCurrent(s)) return;
    od = getOrder();
    if (typeof vendor !== 'undefined' && !od.vendors.some(v => v.id === vendor)) vendor = od.vendors[0] ? od.vendors[0].id : '';
  }
  // 酒款資料：手機上是 {隼:[...], 鳥久:[...]}，以畫面上的 db 為準
  function wineDb() { if (typeof db !== 'undefined' && db) return db; return get(KEYS.wines) || {}; }
  function saveWineDb() { put(KEYS.wines, wineDb()); }
  function ensureWineIds(s) {
    const rows = wineDb()[s] || []; let changed = false;
    rows.forEach(r => { if (r && !r.id) { r.id = uuid(); changed = true; } });
    if (changed) saveWineDb();
    return rows;
  }

  /* ---------- 連線 ---------- */
  async function api(params, post) {
    if (!CLOUD_URL) throw new Error('off');
    if (!key()) throw new Error('nokey');
    const body = Object.assign({ key: key() }, params);
    const res = post
      ? await fetch(CLOUD_URL, { method: 'POST', body: JSON.stringify(body) })
      : await fetch(CLOUD_URL + '?' + new URLSearchParams(body), { cache: 'no-store' });
    const j = await res.json();
    if (!j.ok && j.error === 'auth') { try { LS.removeItem('cloud:key'); } catch (e) { } throw new Error('auth'); }
    if (!j.ok) throw new Error(j.error || '同步失敗');
    if (j.v !== API_VERSION) throw new Error('old');
    if (j.sid) checkSheet(j.sid);
    return j;
  }
  // 換了一份試算表時，清掉「已同步」標記，避免把手機紀錄誤判為已刪除
  function checkSheet(sid) {
    if (getRaw('cloud:sid') === sid) return;
    Object.keys(LS).filter(k => /^cloud:(synced|master|wine):/.test(k)).forEach(k => LS.removeItem(k));
    putRaw('cloud:sid', sid);
  }

  /* ---------- 上傳 ---------- */
  function pushOrders(kind, s) {
    return serial(kind + ':' + s, async () => {
      if (!CLOUD_URL || !key() || !s) return;
      setBadge('sync');
      try {
        const list = readList(kind, s), S = syncedSet(kind, s), ids = new Set(list.map(h => String(h.id)));
        for (const h of list) {
          if (S.has(String(h.id))) continue;
          await api({ action: 'saveOrder', store: s, kind, order: h }, true);
          S.add(String(h.id)); saveSynced(kind, s, S);
        }
        for (const id of [...S]) {
          if (ids.has(id)) continue;
          await api({ action: 'deleteOrder', store: s, kind, id }, true);
          S.delete(id); saveSynced(kind, s, S);
        }
        if (kind === 'food') await pushMasterNow(s);
        lastOk = new Date(); lastErr = '';
      } catch (e) { lastErr = e.message; }
      setBadge();
    });
  }
  // 只上傳這支手機已經和雲端對過的店；第一次對照在 pullMaster 進行
  async function pushMasterNow(s, first) {
    const o = readFood(s), h = masterHash(o), last = getRaw('cloud:master:' + s);
    if (h === last || (!last && !first)) return;
    const r = await api({ action: 'saveMaster', store: s, vendors: o.vendors || [], items: o.items || [] }, true);
    // 避免舊版 GAS 將「6包／半斤」誤轉成空值，或漏存兩店商品備註
    const localById = new Map((o.items || []).map(it => [String(it.id), it]));
    const returned = new Map((r.items || []).map(it => [String(it.id), it]));
    const invalid = [...localById.entries()].filter(([id, it]) => {
      const remote = returned.get(id);
      if (!remote) return true;
      for (const field of ['week', 'holiday', 'note']) {
        if (String(it[field] ?? '').trim() !== String(remote[field] ?? '').trim()) return true;
      }
      return false;
    });
    if (invalid.length) throw new Error('雲端試算表程式未保留 '+invalid.length+' 筆安全庫存／備註，請更新 Apps Script v3.2 並重新部署。已保留本機資料');
    if (r.others) { put('cloud:others:' + s, r.others); put('cloud:vendors', r.vendors); }
    putRaw('cloud:master:' + s, h);
  }
  async function pushWinesNow(s, first) {
    const last = getRaw('cloud:wine:' + s);
    if (!last && !first) return;
    const rows = ensureWineIds(s), h = wineHash(rows);
    if (h === last) return;
    await api({ action: 'saveWines', store: s, rows }, true);
    putRaw('cloud:wine:' + s, h);
  }
  function pushWines(s) {
    return serial('wines:' + s, async () => {
      if (!CLOUD_URL || !key() || !s) return;
      setBadge('sync');
      try { await pushWinesNow(s); lastOk = new Date(); lastErr = ''; } catch (e) { lastErr = e.message; }
      setBadge();
    });
  }
  function pushSoon(name, fn) { clearTimeout(timers[name]); timers[name] = setTimeout(fn, 700); }

  /* ---------- 下載：紀錄合併 ---------- */
  function merge(kind, s, cloudOrders, from, to) {
    const L = readList(kind, s), S = syncedSet(kind, s);
    const C = new Map(cloudOrders.map(h => [String(h.id), h]));
    const inRange = h => { const m = String(h.date || '').slice(0, 7); return m >= from && m <= to; };
    const kept = L.filter(h => !inRange(h) || (!C.has(String(h.id)) && !S.has(String(h.id))));
    const next = kept.concat([...C.values()]).sort((a, b) => String(a.date).localeCompare(String(b.date)));
    C.forEach((_, id) => S.add(id));
    L.forEach(h => { if (inRange(h) && S.has(String(h.id)) && !C.has(String(h.id))) S.delete(String(h.id)); });
    saveSynced(kind, s, S);
    writeList(kind, s, next);
  }

  /* ---------- 下載：共用品項 ---------- */
  async function pullMaster(s) {
    const m = await api({ action: 'master', store: s });
    put('cloud:vendors', m.vendors);
    const o = readFood(s), last = getRaw('cloud:master:' + s);
    if (last && masterHash(o) !== last) { await pushMasterNow(s); return; } // 這支手機有尚未上傳的修改
    const vname = id => (m.vendors.find(v => v.id === id) || {}).name || '';
    const items = m.items.slice(), others = m.others.slice(), vendorIds = new Set();
    // 雲端舊版可能捨棄文字型安全庫存。先保留本機仍有效的平假日及備註。
    const localVendor = id => ((o.vendors || []).find(v => v.id === id) || {}).name || '';
    const byName = new Map((o.items || []).map(it => [norm(localVendor(it.vendor))+'|'+norm(it.name), it]));
    let protectedFields=0;
    items.forEach(it=>{
      const old=(o.items||[]).find(d=>String(d.id)===String(it.id))||byName.get(norm(vname(it.vendor))+'|'+norm(it.name));
      if(!old)return;
      for(const field of ['week','holiday','note']){
        if(!String(it[field]??'').trim()&&String(old[field]??'').trim()){it[field]=old[field];protectedFields++}
      }
    });
    let added = false;
    if (!last && (o.items || []).length) {
      // 這支手機第一次同步：保留手機上雲端沒有的品項
      const lvname = id => ((o.vendors || []).find(v => v.id === id) || {}).name || '';
      for (const it of o.items) {
        const k = norm(lvname(it.vendor)) + '|' + norm(it.name);
        if (m.items.some(c => norm(vname(c.vendor)) + '|' + norm(c.name) === k)) continue;
        const oi = others.findIndex(c => norm(vname(c.vendor)) + '|' + norm(c.name) === k);
        if (oi >= 0) { const c = others.splice(oi, 1)[0]; items.push(Object.assign({}, c, { cat: it.cat || '備品', week: it.week, holiday: it.holiday })); added = true; continue; }
        let cv = m.vendors.find(v => norm(v.name) === norm(lvname(it.vendor)));
        if (!cv) { cv = (o.vendors || []).find(v => v.id === it.vendor); if (cv) m.vendors.push(cv); }
        items.push(Object.assign({}, it, { vendor: cv ? cv.id : it.vendor })); added = true;
      }
    }
    items.forEach(i => vendorIds.add(i.vendor));
    const catalogVendors = new Set(m.items.concat(m.others).map(i => i.vendor));
    o.vendors = m.vendors.filter(v => vendorIds.has(v.id) || !catalogVendors.has(v.id));
    o.items = items;
    put(KEYS.food + s, o); put('cloud:others:' + s, others); syncMemory(s);
    if (added || protectedFields) await pushMasterNow(s, true); else putRaw('cloud:master:' + s, masterHash(o));
  }

  /* ---------- 下載：共用酒款 ---------- */
  async function pullWines(s) {
    const r = await api({ action: 'wines', store: s });
    const local = ensureWineIds(s), last = getRaw('cloud:wine:' + s);
    if (last && wineHash(local) !== last) { await pushWinesNow(s); put('cloud:wothers:' + s, r.others); return; }
    const key2 = w => (w.sku ? 'sku:' + norm(w.sku) : norm(w.name) + '|' + (Number(w.ml) || ''));
    const items = r.items.slice(), others = r.others.slice();
    let added = false;
    if (!last) {
      for (const w of local) {
        const k = key2(w), k2 = norm(w.name) + '|' + (Number(w.ml) || '');
        const hit = c => key2(c) === k || norm(c.name) + '|' + (Number(c.ml) || '') === k2;
        if (r.items.some(hit)) continue;
        const oi = others.findIndex(hit);
        if (oi >= 0) { items.push(others.splice(oi, 1)[0]); added = true; continue; }
        items.push(w); added = true;
      }
    }
    const d = wineDb(); d[s] = items; saveWineDb();
    put('cloud:wothers:' + s, others);
    if (added) await pushWinesNow(s, true); else putRaw('cloud:wine:' + s, wineHash(items));
  }

  /* ---------- 進入店別時的完整同步 ---------- */
  async function pull(s) {
    if (!CLOUD_URL || !s) return;
    if (!key()) { askKey(); if (!key()) { setBadge(); return; } }
    setBadge('sync');
    try {
      await serial('food:' + s, () => pullMaster(s));
      await serial('wines:' + s, () => pullWines(s));
      const from = monthsAgo(2), to = '9999-99';
      for (const kind of ['food', 'wine']) {
        const r = await api({ action: 'history', store: s, kind, from, to });
        await serial(kind + ':' + s, () => merge(kind, s, r.orders || [], from, to));
        fetchedMonths[kind + s] = from;
      }
      lastOk = new Date(); lastErr = '';
    } catch (e) { lastErr = e.message; }
    setBadge();
    refreshView(s);
    await flushAddWines();
    await pushOrders('food', s); await pushOrders('wine', s); await pushWines(s);
  }

  async function pullMonth(kind, s, month) {
    if (!CLOUD_URL || !key() || !month || month >= (fetchedMonths[kind + s] || '9999-99')) return false;
    const done = 'm:' + kind + s + month;
    if (fetchedMonths[done]) return false;
    try {
      const r = await api({ action: 'history', store: s, kind, from: month, to: month });
      merge(kind, s, r.orders || [], month, month); fetchedMonths[done] = true; return true;
    } catch (e) { lastErr = e.message; setBadge(); return false; }
  }

  function refreshView(s) {
    if (!isCurrent(s)) return;
    syncMemory(s);
    const h = (document.getElementById('heading') || {}).textContent || '';
    if (h.includes('叫貨管理') && typeof mode !== 'undefined' && mode !== 'order') drawOrder();
    else if (h.includes('叫貨管理') && document.getElementById('oi')) drawItems();
    if (h.includes('酒類叫貨紀錄') && document.getElementById('whl')) drawWineHist();
    const w = document.getElementById('wine');
    if (w && w.style.display !== 'none' && typeof render === 'function') { try { render(); } catch (e) { } }
  }

  /* ---------- 加入另一店的共用品項／酒款 ---------- */
  function adoptFoodPanel() {
    const ob = document.getElementById('ob'); if (!ob) return;
    const others = get('cloud:others:' + current) || [];
    const vendors = get('cloud:vendors') || [];
    const vn = id => (vendors.find(v => v.id === id) || {}).name || '';
    const box = document.createElement('div');
    box.innerHTML = '<h2 style="margin-top:24px">共用品項（本店未使用）</h2><p class="muted">另一店在用、或試算表裡本店未勾選的品項。點一下可勾選使用台、填安全量後加入本店。</p><input id="adoptQ" placeholder="搜尋品名或廠商"><div id="adoptL"></div>';
    ob.appendChild(box);
    const draw = () => {
      const q = norm(document.getElementById('adoptQ').value), L = document.getElementById('adoptL'); L.replaceChildren();
      const list = others.filter(o => !q || norm(o.name + vn(o.vendor)).includes(q));
      if (!list.length) { L.innerHTML = '<p class="muted">' + (others.length ? '沒有符合的品項' : '目前沒有其他可加入的品項') + '</p>'; return; }
      list.slice(0, 80).forEach(o => {
        const d = document.createElement('div'); d.className = 'item';
        d.textContent = vn(o.vendor) + '｜' + o.name + '｜' + o.price + '/' + o.unit;
        d.onclick = () => {
          if (!od.vendors.some(v => v.id === o.vendor)) { const v = vendors.find(v => v.id === o.vendor); if (v) od.vendors.push(Object.assign({}, v)); }
          editItem(o.id, Object.assign({}, o, { cat: '' }));
        };
        L.appendChild(d);
      });
      if (list.length > 80) L.insertAdjacentHTML('beforeend', '<p class="muted">還有 ' + (list.length - 80) + ' 筆，請用搜尋縮小範圍</p>');
    };
    document.getElementById('adoptQ').oninput = draw; draw();
  }
  function adoptWinePage() {
    show('landing');
    document.getElementById('heading').textContent = current + '居酒屋｜加入共用酒款';
    const others = get('cloud:wothers:' + current) || [];
    const L = document.getElementById('landing');
    L.innerHTML = '<button class="alt" onclick="wineHub()">返回酒類管理</button><h2 style="margin-top:14px">共用酒款（本店未使用）</h2><p class="muted">另一店有賣、本店還沒上架的酒款。加入後售價、杯價、介紹與另一店共用。</p><input id="wq" placeholder="搜尋酒名或代理商"><div id="wal"></div>';
    const draw = () => {
      const q = norm(document.getElementById('wq').value), box = document.getElementById('wal'); box.replaceChildren();
      const list = others.filter(w => !q || norm(w.name + w.supplier).includes(q));
      if (!list.length) { box.innerHTML = '<p class="muted">' + (others.length ? '沒有符合的酒款' : '目前沒有其他可加入的酒款') + '</p>'; return; }
      list.slice(0, 80).forEach(w => {
        const d = document.createElement('div'); d.className = 'item';
        const b = document.createElement('button'); b.textContent = '加入本店'; b.style.marginTop = '6px';
        d.textContent = w.name.replace(/\n/g, ' ') + '｜' + (w.ml || '') + 'ml｜' + (w.supplier || '未填代理商') + '｜售價 ' + (w.price || '—');
        d.appendChild(document.createElement('br')); d.appendChild(b);
        b.onclick = () => {
          const dbx = wineDb(); (dbx[current] = dbx[current] || []).push(Object.assign({}, w));
          others.splice(others.indexOf(w), 1); put('cloud:wothers:' + current, others);
          persist(); draw();
        };
        box.appendChild(d);
      });
    };
    document.getElementById('wq').oninput = draw; draw();
  }

  /* ---------- 同步密碼與狀態標示 ---------- */
  function askKey(force) {
    if (!CLOUD_URL || (needKeyAsked && !force)) return;
    needKeyAsked = true;
    const k = prompt('請輸入雲端同步密碼（向店長索取）。\n輸入一次後這支手機會記住。');
    if (k && k.trim()) putRaw('cloud:key', k.trim());
  }
  function pendingCount(s) {
    if (!s) return 0;
    let n = 0;
    for (const kind of ['food', 'wine']) { const S = syncedSet(kind, s); n += readList(kind, s).filter(h => !S.has(String(h.id))).length; }
    return n;
  }
  function setBadge(state) {
    let b = document.getElementById('cloudBadge');
    const top = document.getElementById('top');
    if (!b && top) {
      b = document.createElement('button'); b.id = 'cloudBadge'; b.type = 'button';
      b.style.cssText = 'margin-top:8px;font-size:12px;padding:4px 10px;border-radius:999px;border:1px solid #666;background:transparent;color:#ddd;cursor:pointer';
      b.onclick = () => { if (!CLOUD_URL) return alert('尚未設定雲端同步網址，資料只存在這支手機。'); if (!key() || lastErr === 'auth') askKey(true); const s = typeof current !== 'undefined' ? current : ''; if (s) pull(s); };
      top.appendChild(document.createElement('br')); top.appendChild(b);
    }
    if (!b) return;
    const s = typeof current !== 'undefined' ? current : '';
    const p = pendingCount(s);
    let t;
    if (!CLOUD_URL) t = '本機模式（未連雲端）';
    else if (!key()) t = '☁ 點此輸入同步密碼';
    else if (state === 'sync') t = '☁ 同步中…';
    else if (lastErr === 'auth') t = '☁ 同步密碼錯誤，點此重新輸入';
    else if (lastErr === 'old') t = '☁ 試算表程式需要更新（請店長更新 Apps Script）';
    else if (lastErr) t = '☁ 暫時無法同步' + (p ? '・' + p + ' 筆待上傳' : '') + '（點此重試）';
    else if (p) t = '☁ ' + p + ' 筆待上傳';
    else t = '☁ 已同步' + (lastOk ? ' ' + lastOk.toTimeString().slice(0, 5) : '');
    b.textContent = t;
  }

  /* ---------- 接上原本的功能 ---------- */
  Storage.prototype.setItem = function (k, v) {
    rawSet.call(this, k, v);
    if (applying || this !== LS) return;
    const m = /^(ordersV2|wineOrdersV3):(.+)$/.exec(k);
    if (m) { const kind = m[1] === 'ordersV2' ? 'food' : 'wine', s = m[2]; pushSoon(kind + s, () => pushOrders(kind, s)); }
    if (k === KEYS.wines) STORES.forEach(s => pushSoon('wines' + s, () => pushWines(s)));
  };
  if (typeof enter === 'function') {
    const _enter = enter;
    enter = function (s) { const r = _enter.apply(this, arguments); setBadge(); pull(s); return r; };
  }
  if (typeof home === 'function') {
    const _home = home;
    home = function () { const r = _home.apply(this, arguments); setBadge(); return r; };
  }
  if (typeof drawOrder === 'function') {
    const _do = drawOrder;
    drawOrder = function () { const r = _do.apply(this, arguments); if (typeof mode !== 'undefined' && mode === 'settings' && CLOUD_URL) adoptFoodPanel(); return r; };
  }
  if (typeof wineHub === 'function') {
    const _wh = wineHub;
    wineHub = function () {
      const r = _wh.apply(this, arguments);
      if (CLOUD_URL) {
        const row = document.createElement('div'); row.className = 'row'; row.style.marginTop = '12px';
        const b = document.createElement('button'); b.textContent = '從共用酒款加入'; b.onclick = adoptWinePage;
        row.appendChild(b); document.getElementById('landing').appendChild(row);
      }
      return r;
    };
  }
  if (typeof wineMessage === 'function') {
    const _wm = wineMessage;
    wineMessage = function () { const h = _wm.apply(this, arguments); if (h && typeof wWho !== 'undefined') h.who = wWho; return h; };
  }
  if (typeof historyDraw === 'function') {
    const _hd = historyDraw;
    historyDraw = function () {
      const r = _hd.apply(this, arguments), mo = (document.getElementById('month') || {}).value;
      pullMonth('food', current, mo).then(ok => { if (ok && document.getElementById('hist')) { syncMemory(current); _hd(); } });
      return r;
    };
  }
  if (typeof drawWineHist === 'function') {
    const _dw = drawWineHist;
    drawWineHist = function () {
      const r = _dw.apply(this, arguments), mo = (document.getElementById('whm') || {}).value;
      pullMonth('wine', current, mo).then(ok => { if (ok && document.getElementById('whl')) _dw(); });
      return r;
    };
  }
  // 報價單匯入：新增／更新酒款並標記兩店；失敗時保留，下次同步重送
  window.cloudAddWines = function (list) {
    const box = (get('cloud:pendingAddWines') || []).concat(list || []);
    put('cloud:pendingAddWines', box);
    return flushAddWines();
  };
  function flushAddWines() {
    return serial('addWines', async () => {
      const box = get('cloud:pendingAddWines') || [];
      if (!box.length || !CLOUD_URL || !key()) return;
      try { await api({ action: 'addWines', store: STORES[0], stores: STORES, rows: box }, true); put('cloud:pendingAddWines', []); lastErr = ''; }
      catch (e) { lastErr = e.message; }
      setBadge();
    });
  }
  /* 匯入 Excel 後立即同步店別主檔，避免尚未建立 cloud:master 標記時跳過資料 */
  window.cloudSyncImportedFood = async function (shops) {
    if (!CLOUD_URL) return { state: 'local', message: '此網站尚未設定雲端網址，資料目前只保存在本機' };
    if (!key()) return { state: 'login', message: '尚未輸入雲端同步密碼，請點擊雲端狀態按鈕登入後再同步' };
    const done = [];
    try {
      for (const s of shops) {
        if (!STORES.includes(s)) continue;
        await serial('food:' + s, () => pushMasterNow(s, true));
        done.push(s);
      }
      lastOk = new Date();
      lastErr = '';
      setBadge();
      return { state: 'ok', stores: done, message: done.join('、') + ' 的品項及安全庫存已送交雲端主檔' };
    } catch (e) {
      lastErr = e.message;
      setBadge();
      return { state: 'error', stores: done, message: '雲端更新失敗：' + e.message + '。本機匯入資料仍保留，請稍後重試同步' };
    }
  };
  window.addEventListener('online', () => { if (typeof current !== 'undefined' && current) { pushOrders('food', current); pushOrders('wine', current); pushWines(current); } });
  setBadge();
})();
