/* OrderFlow store —— 存储适配层。
 * 本地模式用 IndexedDB（与旧版行为一致）；云端模式接腾讯云开发 CloudBase。
 * 通过 window.ORDERFLOW_CONFIG 切换：{ mode: 'local' | 'cloudbase', envId: '...' }。
 */
(function () {
  'use strict';

  const CFG = Object.assign({ mode: 'local', envId: '', collection: 'workspaces', docId: 'main' }, window.ORDERFLOW_CONFIG || {});
  const DB_NAME = 'orderflow-workbench-v1';
  const STORE_NAME = 'workspace';
  const KEY = 'state';
  const SDK_URL = 'https://static.cloudbase.net/cloudbase-js-sdk/2.7.1/cloudbase.full.js';

  let dbPromise = null;
  let appPromise = null;
  let lastRev = 0;
  let watchStarted = false;
  const listeners = new Set();

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
    const db = await openDb();
    await new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_NAME, 'readwrite');
      tx.objectStore(STORE_NAME).put(state, KEY);
      tx.oncomplete = resolve;
      tx.onerror = () => reject(tx.error);
      tx.onabort = () => reject(tx.error);
    });
  }

  function loadScript(src) {
    return new Promise((resolve, reject) => {
      if (document.querySelector('script[data-orderflow-sdk]')) return resolve();
      const el = document.createElement('script');
      el.src = src;
      el.async = true;
      el.dataset.orderflowSdk = '1';
      el.onload = resolve;
      el.onerror = () => reject(new Error('云开发 SDK 加载失败，请检查网络。'));
      document.head.append(el);
    });
  }

  async function cloudApp() {
    if (appPromise) return appPromise;
    if (!CFG.envId) throw new Error('缺少 CloudBase 环境 ID（ORDERFLOW_CONFIG.envId）。');
    appPromise = (async () => {
      if (!window.cloudbase) await loadScript(SDK_URL);
      if (!window.cloudbase) throw new Error('云开发 SDK 不可用。');
      return window.cloudbase.init({ env: CFG.envId });
    })();
    return appPromise;
  }

  async function cloudAuth() {
    const app = await cloudApp();
    const auth = app.auth({ persistence: 'local' });
    if (!auth.hasLoginState || !(await auth.hasLoginState())) throw new Error('未登录');
    return auth;
  }

  async function cloudLoad() {
    const app = await cloudApp();
    await cloudAuth();
    const res = await app.database().collection(CFG.collection).doc(CFG.docId).get();
    const doc = res && res.data && res.data[0];
    if (!doc) return null;
    lastRev = Number(doc.rev) || 0;
    return { state: doc.data, rev: lastRev, source: 'cloudbase' };
  }

  async function cloudSave(state) {
    const app = await cloudApp();
    await cloudAuth();
    const col = app.database().collection(CFG.collection);
    const rev = lastRev + 1;
    const payload = { data: state, rev, updatedAt: new Date().toISOString(), updatedBy: (app.auth().currentUser || {}).uid || '' };
    if (lastRev === 0) {
      await col.add(Object.assign({ _id: CFG.docId }, payload));
    } else {
      const res = await col.where({ _id: CFG.docId, rev: lastRev }).update(payload);
      const updated = (res && (res.updated || res.stats && res.stats.updated)) || 0;
      if (!updated) throw new Error('数据已被他人修改，请刷新后重试。');
    }
    lastRev = rev;
  }

  const OrderStore = {
    get mode() { return CFG.mode; },
    get envId() { return CFG.envId; },
    get rev() { return lastRev; },
    get available() { return true; },
    async load() { return CFG.mode === 'cloudbase' ? cloudLoad() : localLoad(); },
    async save(state) { return CFG.mode === 'cloudbase' ? cloudSave(state) : localSave(state); },
    onExternalChange(cb) { listeners.add(cb); },
    async startWatch() {
      if (watchStarted || CFG.mode !== 'cloudbase') return;
      watchStarted = true;
      try {
        const app = await cloudApp();
        await cloudAuth();
        app.database().collection(CFG.collection).doc(CFG.docId).watch({
          onChange: snap => {
            const doc = snap && snap.docs && snap.docs[0];
            if (!doc || !doc.data) return;
            const rev = Number(doc.rev) || 0;
            if (rev <= lastRev) return;
            lastRev = rev;
            listeners.forEach(cb => { try { cb(doc.data, rev); } catch (_) {} });
          },
          onError: () => {}
        });
      } catch (_) { watchStarted = false; }
    }
  };

  window.OrderStore = OrderStore;
})();
