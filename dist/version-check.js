/* 新版本检测：读取 gh-pages 分支里的 version.json（走 GitHub API，不吃 CDN 缓存），
 * 与页面内嵌的 build 版本比较，发现更新就提示刷新。
 */
(function () {
  'use strict';
  const meta = name => (document.querySelector('meta[name="' + name + '"]') || {}).content || '';
  const build = meta('orderflow-build');
  const repo = meta('orderflow-repo') || 'fengweiheaven/NewShop';
  if (!build || build === 'local') return;

  function banner(version) {
    if (document.getElementById('update-banner')) return;
    const box = document.createElement('div');
    box.id = 'update-banner';
    box.className = 'update-banner';
    box.innerHTML = '<span>有新版本可用（' + version + '）</span>'
      + '<button type="button" class="btn small" id="update-reload">立即刷新</button>'
      + '<button type="button" class="update-close" aria-label="关闭">×</button>';
    document.body.append(box);
    box.querySelector('#update-reload').onclick = () => location.reload();
    box.querySelector('.update-close').onclick = () => box.remove();
  }

  async function check() {
    try {
      const url = 'https://api.github.com/repos/' + repo + '/contents/version.json?ref=gh-pages';
      const r = await fetch(url, { cache: 'no-store', headers: { Accept: 'application/vnd.github.raw' } });
      if (!r.ok) return;
      const text = await r.text();
      let data = null;
      try { data = JSON.parse(text); } catch (_) { return; }
      if (data && data.version && data.version !== build) banner(data.version);
    } catch (_) { }
  }

  setTimeout(check, 6000);
  setInterval(check, 300000);
})();
