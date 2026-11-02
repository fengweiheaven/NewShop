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

  const ROLE_NAMES = { admin: '管理员', clerk: '客服', warehouse: '仓库', finance: '财务' };
  let profile = null;
  async function loadProfile() {
    const u = currentUser();
    if (!u) return null;
    let row = null;
    try {
      const { data } = await state.app.rdb().from('app_users').select('*').eq('uid', u.uid).limit(1);
      row = Array.isArray(data) && data[0];
    } catch (_) { }
    return {
      uid: u.uid,
      name: (row && row.name) || u.nickname || u.name || u.username || u.phone || '',
      phone: (row && row.phone) || u.phone || '',
      role: (row && row.role) || ''
    };
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
        profile = await loadProfile();
        state.role = profile ? profile.role : '';
        document.body.classList.remove('auth-locked');
        gate.remove();
        decorateHeader();
        resolveReady({ user: state.user, role: state.role });
      } catch (e) {
        err.textContent = (e && (e.message || e.error_description)) || '登录失败，请检查手机号和密码。';
        btn.disabled = false; btn.textContent = '登录';
      }
    };
    form.querySelector('[name=username]')?.focus();
  }

  function el(tag, cls, html) { const n = document.createElement(tag); if (cls) n.className = cls; if (html != null) n.innerHTML = html; return n; }
  function closeMenu() { const m = document.getElementById('user-menu'); if (m) m.remove(); }
  function showMenu(anchor, p) {
    closeMenu();
    const menu = el('div', 'user-menu');
    menu.id = 'user-menu';
    menu.innerHTML = '<div class="user-menu-head"><strong>' + (p.name || p.phone || '已登录') + '</strong><small>' + (p.phone || '') + '</small></div>'
      + '<div class="user-menu-role">角色：' + (ROLE_NAMES[p.role] || '未分配') + '</div>'
      + '<div class="user-menu-version">版本 ' + ((document.querySelector('meta[name=orderflow-build]') || {}).content || 'local') + '</div>'
      + '<button type="button" class="user-menu-item plain" id="nickname-btn">修改昵称</button>'
      + '<button type="button" class="user-menu-item" id="logout-btn">退出登录</button>';
    document.body.append(menu);
    const r = anchor.getBoundingClientRect();
    menu.style.top = (r.bottom + 8) + 'px';
    menu.style.right = Math.max(12, window.innerWidth - r.right) + 'px';
    menu.addEventListener('click', e => e.stopPropagation());
    document.getElementById('nickname-btn').onclick = () => { closeMenu(); openNicknameDialog(); };
    document.getElementById('logout-btn').onclick = async () => {
      try { await state.auth.signOut(); } catch (_) { }
      location.reload();
    };
    setTimeout(() => document.addEventListener('click', closeMenu, { once: true }), 0);
  }
  function openNicknameDialog() {
    const p = profile || {};
    const wrap = el('div', 'mini-dialog');
    wrap.id = 'nickname-dialog';
    wrap.innerHTML = '<form class="mini-card"><h3>修改昵称</h3>'
      + '<p class="muted">昵称只用于界面显示，不改登录手机号。</p>'
      + '<label class="field"><span>昵称</span><input name="nickname" value="' + (p.name && p.name !== p.phone ? p.name : '') + '" placeholder="例如：老板 / 小张" required></label>'
      + '<p class="form-error" id="nickname-error"></p>'
      + '<div class="mini-actions"><button type="button" class="btn secondary" id="nickname-cancel">取消</button><button type="submit" class="btn" id="nickname-save">保存</button></div></form>';
    document.body.append(wrap);
    const form = wrap.querySelector('form');
    wrap.querySelector('#nickname-cancel').onclick = () => wrap.remove();
    form.onsubmit = async e => {
      e.preventDefault();
      const value = String(new FormData(form).get('nickname') || '').trim();
      if (!value) { wrap.querySelector('#nickname-error').textContent = '请填写昵称。'; return; }
      const btn = wrap.querySelector('#nickname-save');
      btn.disabled = true; btn.textContent = '保存中…';
      try {
        const uid = (currentUser() || {}).uid || '';
        const { error } = await state.app.rdb().from('app_users').upsert({ uid, name: value, phone: p.phone || '', role: p.role || 'clerk' });
        if (error) throw new Error(error.message || '保存失败');
        profile = Object.assign({}, p, { name: value, role: p.role || 'clerk' });
        state.role = profile.role;
        updateChip();
        wrap.remove();
      } catch (err) {
        wrap.querySelector('#nickname-error').textContent = (err && err.message) || '保存失败';
        btn.disabled = false; btn.textContent = '保存';
      }
    };
    form.querySelector('[name=nickname]').focus();
  }
  function updateChip() {
    const chip = document.getElementById('user-chip');
    if (!chip || !profile) return;
    const initial = (profile.name || profile.phone || '?').slice(0, 1);
    chip.querySelector('.user-avatar').textContent = initial;
    chip.querySelector('.user-meta strong').textContent = profile.name || profile.phone || '已登录';
    chip.querySelector('.user-meta small').textContent = ROLE_NAMES[profile.role] || '未分配角色';
  }
  function decorateHeader() {
    const badge = document.querySelector('.save-badge');
    if (badge) badge.innerHTML = '<i></i>已连接云端 · 自动保存';
    const actions = document.querySelector('.topbar-actions');
    if (!actions || document.getElementById('user-chip')) return;
    const p = profile || { name: '', phone: '', role: '' };
    const initial = (p.name || p.phone || '?').slice(0, 1);
    const chip = el('button', 'user-chip');
    chip.type = 'button';
    chip.id = 'user-chip';
    chip.innerHTML = '<span class="user-avatar">' + initial + '</span><span class="user-meta"><strong>' + (p.name || p.phone || '已登录') + '</strong><small>' + (ROLE_NAMES[p.role] || '未分配角色') + '</small></span>';
    chip.onclick = e => { e.stopPropagation(); showMenu(chip, p); };
    const divider = actions.querySelector('.topbar-divider');
    actions.insertBefore(chip, divider || null);
  }
  async function start() {
    if (!cloudMode) { resolveReady({ user: null, role: 'local' }); return; }
    try {
      await ensureApp();
      const user = currentUser();
      if (user && user.uid) {
        state.user = user;
        profile = await loadProfile();
        state.role = profile ? profile.role : '';
        decorateHeader();
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
    get app() { return state.app; },
    get auth() { return state.auth; },
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
