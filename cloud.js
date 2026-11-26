/* 鳥久 × 隼｜雲端同步模組
 * 把每支手機的叫貨資料（廠商、品項、食材叫貨紀錄、酒類叫貨紀錄）同步到 Google 試算表。
 * 原本的功能程式不需修改：本模組在資料寫入手機後自動上傳，進入店別時自動下載。
 */
(function () {
  const CLOUD_URL = ''; // ← Apps Script「網頁應用程式網址」（結尾 /exec）

  const LS = window.localStorage;
  const rawSet = Storage.prototype.setItem;
  const rawGet = Storage.prototype.getItem;
  const KEYS = { food: 'ordersV2:', wine: 'wineOrdersV3:' };
  let applying = false, needKeyAsked = false, lastOk = null, lastErr = '';
  const chains = {}, timers = {}, fetchedMonths = {};

  /* ---------- 小工具 ---------- */
  const get = k => { try { return JSON.parse(rawGet.call(LS, k)); } catch (e) { return null; } };
  const put = (k, v) => { try { applying = true; rawSet.call(LS, k, JSON.stringify(v)); } catch (e) { } finally { applying = false; } };
  const key = () => rawGet.call(LS, 'cloud:key') || '';
  const putRaw = (k, v) => { try { applying = true; rawSet.call(LS, k, v); } catch (e) { } finally { applying = false; } };
  const monthsAgo = n => { const d = new Date(); d.setUTCDate(1); d.setUTCMonth(d.getUTCMonth() - n); return d.toISOString().slice(0, 7); };
  const hash = s => { let h = 5381; for (let i = 0; i < s.length; i++) h = ((h << 5) + h + s.charCodeAt(i)) | 0; return String(h >>> 0); };
  const masterHash = o => hash(JSON.stringify({ v: o.vendors || [], i: o.items || [] }));
  const readFood = s => get(KEYS.food + s) || { vendors: [], items: [], history: [] };
  const readList = (kind, s) => kind === 'food' ? (readFood(s).history || []) : (get(KEYS.wine + s) || []);
  const syncedSet = (kind, s) => new Set(get('cloud:synced:' + kind + ':' + s) || []);
  const saveSynced = (kind, s, set) => put('cloud:synced:' + kind + ':' + s, [...set]);
  function writeList(kind, s, list) {
    if (kind === 'food') { const o = readFood(s); o.history = list; put(KEYS.food + s, o); syncMemory(s); }
    else put(KEYS.wine + s, list);
  }
  // 寫入手機後立刻更新畫面用的資料，避免之後以舊資料覆蓋
  function syncMemory(s) {
    if (typeof current === 'undefined' || current !== s) return;
    od = getOrder();
    if (typeof vendor !== 'undefined' && !od.vendors.some(v => v.id === vendor)) vendor = od.vendors[0] ? od.vendors[0].id : '';
  }
  function serial(name, fn) { chains[name] = (chains[name] || Promise.resolve()).then(fn, fn); return chains[name]; }

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
    if (j.sid) checkSheet(j.sid);
    return j;
  }
  // 換了一份試算表時，清掉「已同步」標記，避免把手機紀錄誤判為已刪除
  function checkSheet(sid) {
    const old = rawGet.call(LS, 'cloud:sid');
    if (old === sid) return;
    Object.keys(LS).filter(k => k.startsWith('cloud:synced:') || k.startsWith('cloud:master:')).forEach(k => LS.removeItem(k));
    applying = true; try { rawSet.call(LS, 'cloud:sid', sid); } finally { applying = false; }
  }

  /* ---------- 上傳：比對手機與已同步清單，補傳新增、同步刪除 ---------- */
  function push(kind, s) {
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
        if (kind === 'food') {
          const o = readFood(s), h = masterHash(o);
          if (h !== rawGet.call(LS, 'cloud:master:' + s)) {
            await api({ action: 'saveMaster', store: s, vendors: o.vendors || [], items: o.items || [] }, true);
            putRaw('cloud:master:' + s, h);
          }
        }
        lastOk = new Date(); lastErr = '';
      } catch (e) { lastErr = e.message; }
      setBadge();
    });
  }
  function pushSoon(kind, s) {
    clearTimeout(timers[kind + s]);
    timers[kind + s] = setTimeout(() => push(kind, s), 700);
  }

  /* ---------- 下載：雲端紀錄與手機合併 ---------- */
  function merge(kind, s, cloudOrders, from, to) {
    const L = readList(kind, s), S = syncedSet(kind, s);
    const C = new Map(cloudOrders.map(h => [String(h.id), h]));
    const inRange = h => { const m = String(h.date || '').slice(0, 7); return m >= from && m <= to; };
    // 範圍內：以雲端為準，另外保留尚未上傳的手機紀錄；範圍外：維持手機原樣
    const kept = L.filter(h => !inRange(h) || (!C.has(String(h.id)) && !S.has(String(h.id))));
    const next = kept.concat([...C.values()]).sort((a, b) => String(a.date).localeCompare(String(b.date)));
    C.forEach((_, id) => S.add(id));
    L.forEach(h => { if (inRange(h) && S.has(String(h.id)) && !C.has(String(h.id))) S.delete(String(h.id)); });
    saveSynced(kind, s, S);
    writeList(kind, s, next);
  }

  async function pull(s) {
    if (!CLOUD_URL || !s) return;
    if (!key()) { askKey(); if (!key()) { setBadge(); return; } }
    setBadge('sync');
    try {
      // 1. 廠商、品項
      const m = await api({ action: 'master', store: s });
      const o = readFood(s), localH = masterHash(o), last = rawGet.call(LS, 'cloud:master:' + s);
      const cloudHas = (m.vendors.length || m.items.length), localHas = (o.vendors.length || o.items.length);
      const dirty = last && localH !== last;
      let adopt = false;
      if (!dirty && cloudHas) {
        if (!last && localHas && localH !== masterHash({ vendors: m.vendors, items: m.items })) {
          adopt = confirm('雲端已有「' + s + '」的 ' + m.items.length + ' 個品項，這支手機也有 ' + o.items.length + ' 個。\n\n確定：改用雲端資料（建議）\n取消：用這支手機的資料覆蓋雲端');
        } else adopt = true;
      }
      if (adopt) { o.vendors = m.vendors; o.items = m.items; put(KEYS.food + s, o); syncMemory(s); putRaw('cloud:master:' + s, masterHash(o)); }
      // 2. 近三個月的叫貨紀錄
      const from = monthsAgo(2), to = '9999-99';
      for (const kind of ['food', 'wine']) {
        const r = await api({ action: 'history', store: s, kind, from, to });
        merge(kind, s, r.orders || [], from, to);
        fetchedMonths[kind + s] = from;
      }
      lastOk = new Date(); lastErr = '';
    } catch (e) { lastErr = e.message; }
    setBadge();
    refreshView(s);
    await push('food', s); await push('wine', s);
  }

  // 查詢更早月份時，另外下載那個月
  async function pullMonth(kind, s, month) {
    if (!CLOUD_URL || !key() || !month || month >= (fetchedMonths[kind + s] || '9999-99')) return false;
    const done = 'cloud:month:' + kind + s + month;
    if (fetchedMonths[done]) return false;
    try {
      const r = await api({ action: 'history', store: s, kind, from: month, to: month });
      merge(kind, s, r.orders || [], month, month); fetchedMonths[done] = true; return true;
    } catch (e) { lastErr = e.message; setBadge(); return false; }
  }

  function refreshView(s) {
    if (typeof current === 'undefined' || current !== s) return;
    syncMemory(s);
    const h = (document.getElementById('heading') || {}).textContent || '';
    if (h.includes('叫貨管理') && typeof mode !== 'undefined' && mode !== 'order') drawOrder();
    else if (h.includes('叫貨管理') && document.getElementById('oi')) drawItems();
    if (h.includes('酒類叫貨紀錄') && document.getElementById('whl')) drawWineHist();
  }

  /* ---------- 同步密碼與狀態標示 ---------- */
  function askKey(force) {
    if (!CLOUD_URL || (needKeyAsked && !force)) return;
    needKeyAsked = true;
    const k = prompt('請輸入雲端同步密碼（向店長索取）。\n輸入一次後這支手機會記住。');
    if (k && k.trim()) { applying = true; try { rawSet.call(LS, 'cloud:key', k.trim()); } finally { applying = false; } }
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
    if (m) pushSoon(m[1] === 'ordersV2' ? 'food' : 'wine', m[2]);
  };
  if (typeof enter === 'function') {
    const _enter = enter;
    enter = function (s) { const r = _enter.apply(this, arguments); setBadge(); pull(s); return r; };
  }
  if (typeof home === 'function') {
    const _home = home;
    home = function () { const r = _home.apply(this, arguments); setBadge(); return r; };
  }
  if (typeof wineMessage === 'function') {
    const _wm = wineMessage;
    wineMessage = function () { const h = _wm.apply(this, arguments); if (h && typeof wWho !== 'undefined') h.who = wWho; return h; };
  }
  if (typeof historyDraw === 'function') {
    const _hd = historyDraw;
    historyDraw = function () {
      const r = _hd.apply(this, arguments), mo = (document.getElementById('month') || {}).value;
      pullMonth('food', current, mo).then(ok => { if (ok && document.getElementById('hist')) { od = getOrder(); _hd(); } });
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
  window.addEventListener('online', () => { if (typeof current !== 'undefined' && current) { push('food', current); push('wine', current); } });
  setBadge();
})();
