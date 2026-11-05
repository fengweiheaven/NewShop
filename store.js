/* OrderFlow store —— 存储适配层。
 * local：IndexedDB；cloudbase：腾讯云开发 PostgreSQL（workspaces 表 + rev 乐观锁）。
 * 配置见 config.js / config.local.js。
 */
(function () {
  'use strict';

  const CFG = Object.assign({ mode: 'local', envId: '', collection: 'workspaces', docId: 'main' }, window.ORDERFLOW_CONFIG || {});
  const DB_NAME = 'orderflow-workbench-v1';
  const STORE_NAME = 'workspace';
  const KEY = 'state';

  let dbPromise = null;
  let lastRev = 0;

  function openDb() {
    if (dbPromise) return dbPromise;
    dbPromise = new Promise((resolve, reject) => {
      const req = indexedDB.open(DB_NAME, 1);
      req.onupgradeneeded = () => req.result.createObjectStore(STORE_NAME);
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
    return dbPromise;
  }

  async function localLoad() {
    const db = await openDb();
    const value = await new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_NAME).objectStore(STORE_NAME).get(KEY);
      tx.onsuccess = () => resolve(tx.result);
      tx.onerror = () => reject(tx.error);
    });
    return value ? { state: value, rev: 0, source: 'local' } : null;
  }

  async function localSave(state) {
    writeSnapshot(state);
    const db = await openDb();
    await new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_NAME, 'readwrite');
      tx.objectStore(STORE_NAME).put(state, KEY);
      tx.oncomplete = resolve;
      tx.onerror = () => reject(tx.error);
      tx.onabort = () => reject(tx.error);
    });
  }

  const SNAP_DB = 'orderflow-snapshots-v1';
  const SNAP_STORE = 'snaps';
  const SNAP_LIMIT = 30;
  let snapPromise = null;
  function snapDb() {
    if (snapPromise) return snapPromise;
    snapPromise = new Promise((resolve, reject) => {
      const req = indexedDB.open(SNAP_DB, 1);
      req.onupgradeneeded = () => req.result.createObjectStore(SNAP_STORE);
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
    return snapPromise;
  }
  async function writeSnapshot(state) {
    try {
      const db = await snapDb();
      const key = new Date().toISOString();
      await new Promise((resolve, reject) => {
        const tx = db.transaction(SNAP_STORE, 'readwrite');
        tx.objectStore(SNAP_STORE).put(JSON.parse(JSON.stringify(state)), key);
        tx.oncomplete = resolve;
        tx.onerror = () => reject(tx.error);
      });
      const keys = await snapshotKeys(db);
      for (const k of keys.slice(SNAP_LIMIT)) {
        await new Promise(res => { const tx = db.transaction(SNAP_STORE, 'readwrite'); tx.objectStore(SNAP_STORE).delete(k); tx.oncomplete = res; tx.onerror = res; });
      }
    } catch (_) { }
  }
  async function snapshotKeys(db) {
    return new Promise((resolve, reject) => {
      const req = db.transaction(SNAP_STORE).objectStore(SNAP_STORE).getAllKeys();
      req.onsuccess = () => resolve((req.result || []).slice().sort().reverse());
      req.onerror = () => reject(req.error);
    });
  }
  async function snapshots() {
    const db = await snapDb();
    const keys = await snapshotKeys(db);
    const out = [];
    for (const k of keys) {
      const value = await new Promise((resolve, reject) => {
        const req = db.transaction(SNAP_STORE).objectStore(SNAP_STORE).get(k);
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
      });
      out.push({ key: k, orders: value && value.orders ? value.orders.length : 0, batches: value && value.batches ? value.batches.length : 0, state: value });
    }
    return out;
  }
  function sql() {
    const app = window.OrderAuth && window.OrderAuth.app;
    if (!app) throw new Error('尚未登录，无法访问云端数据。');
    return app.rdb();
  }

  function uid() {
    const u = window.OrderAuth && window.OrderAuth.user;
    return (u && u.uid) || '';
  }

  async function cloudLoad() {
    const { data, error } = await sql().from(CFG.collection).select('*').eq('id', CFG.docId).limit(1);
    if (error) throw new Error(error.message || '读取云端数据失败。');
    const row = Array.isArray(data) ? data[0] : null;
    if (!row) return null;
    lastRev = Number(row.rev) || 0;
    return { state: row.data, rev: lastRev, source: 'cloudbase' };
  }

  async function cloudSave(state) {
    const db = sql();
    const payload = {
      data: state,
      rev: lastRev + 1,
      updated_at: new Date().toISOString(),
      updated_by: uid()
    };
    if (lastRev === 0) {
      const { error } = await db.from(CFG.collection).upsert(Object.assign({ id: CFG.docId }, payload));
      if (error) throw new Error(error.message || '保存失败。');
    } else {
      const { data, error } = await db.from(CFG.collection).update(payload).eq('id', CFG.docId).eq('rev', lastRev).select();
      if (error) throw new Error(error.message || '保存失败。');
      if (!data || !data.length) { const err = new Error('数据已被他人修改，请刷新后重试。'); err.conflict = true; throw err; }
    }
    lastRev += 1;
    try {
      await db.from('audit_logs').insert({ uid: uid(), action: 'save', detail: { orders: (state.orders || []).length, rev: lastRev } });
    } catch (_) { }
    writeSnapshot(state);
  }

  const OrderStore = {
    get mode() { return CFG.mode; },
    get envId() { return CFG.envId; },
    get rev() { return lastRev; },
    async load() { return CFG.mode === 'cloudbase' ? cloudLoad() : localLoad(); },
    async save(state) { return CFG.mode === 'cloudbase' ? cloudSave(state) : localSave(state); },
    readLocal: localLoad,
    snapshots,
    async restoreSnapshot(key) {
      const list = await snapshots();
      const hit = list.find(x => x.key === key);
      return hit ? hit.state : null;
    }
  };

  window.OrderStore = OrderStore;
})();
