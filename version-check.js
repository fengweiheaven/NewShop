/* 新版本检测：通过 GitHub API 读取 gh-pages 最新提交，
 * 与页面内嵌的 build sha 比较，发现更新就提示刷新。
 */
(function () {
  'use strict';
  const meta = name => (document.querySelector('meta[name="' + name + '"]') || {}).content || '';
  const buildSha = meta('orderflow-sha');
  if (!buildSha) return;
  const REPO = meta('orderflow-repo') || 'fengweiheaven/NewShop';

  function banner(version) {
    if (document.getElementById('update-banner')) return;
    const box = document.createElement('div');
    box.id = 'update-banner';
    box.className = 'update-banner';
    box.innerHTML = '<span>有新版本可用（' + (version || '最新') + '）</span>'
      + '<button type="button" class="btn small" id="update-reload">立即刷新</button>'
      + '<button type="button" class="update-close" aria-label="关闭">×</button>';
    document.body.append(box);
    box.querySelector('#update-reload').onclick = () => location.reload();
    box.querySelector('.update-close').onclick = () => box.remove();
  }

  async function check() {
    try {
      const r = await fetch('https://api.github.com/repos/' + REPO + '/branches/gh-pages', { cache: 'no-store' });
      if (!r.ok) return;
      const j = await r.json();
      const sha = (j.commit && j.commit.sha || '').slice(0, 7);
      if (sha && sha !== buildSha.slice(0, 7)) banner(sha);
    } catch (_) { }
  }

  setTimeout(check, 5000);
  setInterval(check, 60000);
})();
