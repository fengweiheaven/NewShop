/* OrderFlow auth —— 登录门与角色。
 * 本地模式不启用登录；云端模式在页面加载后要求先登录（用户名 = 手机号）。
 */
(function () {
  'use strict';

  const CFG = window.ORDERFLOW_CONFIG || {};
  const cloudMode = CFG.mode === 'cloudbase';
  const state = { app: null, auth: null, user: null, role: '', ready: null };
  let resolveReady;
  state.ready = new Promise(res => { resolveReady = res; });

  function loadSdk() {
    if (window.cloudbase) return Promise.resolve();
    return new Promise((resolve, reject) => {
      const el = document.createElement('script');
      el.src = 'vendor/cloudbase.iife.js';
      el.onload = resolve;
      el.onerror = () => reject(new Error('云开发 SDK 加载失败'));
      document.head.append(el);
    });
  }

  async function ensureApp() {
    if (state.app) return state.app;
    await loadSdk();
    state.app = window.cloudbase.init({ env: CFG.envId, region: CFG.region, accessKey: CFG.accessKey });
    state.auth = state.app.auth();
    return state.app;
  }

  function currentUser() {
    try {
      const s = state.auth && state.auth.hasLoginState && state.auth.hasLoginState();
      return (s && (s.user || s)) || state.user || null;
    } catch (_) { return state.user || null; }
  }

  async function loadRole() {
    const user = currentUser();
    if (!user || !user.uid) return '';
    try {
      const { data } = await state.app.rdb().from('app_users').select('*').eq('uid', user.uid).limit(1);
      const row = Array.isArray(data) && data[0];
      return row ? row.role : '';
    } catch (_) { return ''; }
  }

  function renderGate(message) {
    let gate = document.getElementById('auth-gate');
    if (!gate) {
      gate = document.createElement('div');
      gate.id = 'auth-gate';
      document.body.append(gate);
    }
    gate.innerHTML = `
      <form class="auth-card" id="auth-form" autocomplete="on">
        <div class="auth-brand"><span class="brand-mark">序</span><div><strong>序单</strong><small>商贸订单工作台</small></div></div>
        <h2>登录</h2>
        <p class="muted">用管理员分配的手机号与密码登录。</p>
        <label class="field"><span>手机号</span><input name="username" inputmode="tel" autocomplete="username" required placeholder="11 位手机号"></label>
        <label class="field"><span>密码</span><input name="password" type="password" autocomplete="current-password" required placeholder="密码"></label>
        <p class="form-error" id="auth-error">${message || ''}</p>
        <button class="btn" type="submit" id="auth-submit">登录</button>
        <p class="auth-foot">账号由管理员在云开发控制台创建</p>
      </form>`;
    document.body.classList.add('auth-locked');
    const form = document.getElementById('auth-form');
    const err = document.getElementById('auth-error');
    const btn = document.getElementById('auth-submit');
    form.onsubmit = async e => {
      e.preventDefault();
      const f = new FormData(form);
      const username = String(f.get('username') || '').trim();
      const password = String(f.get('password') || '');
      if (!username || !password) { err.textContent = '请填写手机号和密码。'; return; }
      btn.disabled = true; btn.textContent = '登录中…'; err.textContent = '';
      try {
        const res = await state.auth.signInWithPassword({ username, password });
        if (res && res.error) throw new Error(res.error.message || '登录失败');
        state.user = currentUser();
        state.role = await loadRole();
        document.body.classList.remove('auth-locked');
        gate.remove();
        resolveReady({ user: state.user, role: state.role });
      } catch (e) {
        err.textContent = (e && (e.message || e.error_description)) || '登录失败，请检查手机号和密码。';
        btn.disabled = false; btn.textContent = '登录';
      }
    };
    form.querySelector('[name=username]')?.focus();
  }

  async function start() {
    if (!cloudMode) { resolveReady({ user: null, role: 'local' }); return; }
    try {
      await ensureApp();
      const user = currentUser();
      if (user && user.uid) {
        state.user = user;
        state.role = await loadRole();
        resolveReady({ user, role: state.role });
        return;
      }
      renderGate('');
    } catch (e) {
      renderGate((e && e.message) || '初始化失败');
    }
  }

  window.OrderAuth = {
    get mode() { return cloudMode ? 'cloudbase' : 'local'; },
    get user() { return currentUser(); },
    get role() { return state.role; },
    ready: state.ready,
    async signOut() {
      try { await state.auth.signOut(); } catch (_) { }
      location.reload();
    }
  };

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start);
  else start();
})();
