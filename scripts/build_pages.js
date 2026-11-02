// Build the publishable site into ./public for GitHub Pages.
// Injects cache-busting versions and the runtime config from environment variables.
const fs = require('fs');
const path = require('path');

const SRC = path.join(__dirname, '..', 'dist');
const OUT = path.join(__dirname, '..', 'public');
const ASSETS = ['style.css', 'extra.css', 'seed.js', 'core.js', 'config.js', 'config.local.js', 'store.js', 'auth.js', 'app.js', 'vendor/jszip.min.js'];
const version = (process.env.GITHUB_SHA || Date.now().toString(36)).slice(0, 7);

function copyDir(src, dest) {
  fs.mkdirSync(dest, { recursive: true });
  for (const e of fs.readdirSync(src, { withFileTypes: true })) {
    if (e.name === 'config.local.js' || e.name === 'config.local.js.example') continue;
    const s = path.join(src, e.name);
    const d = path.join(dest, e.name);
    if (e.isDirectory()) copyDir(s, d);
    else fs.copyFileSync(s, d);
  }
}

fs.rmSync(OUT, { recursive: true, force: true });
copyDir(SRC, OUT);

const env = process.env.CB_ENV_ID || '';
const region = process.env.CB_REGION || 'ap-shanghai';
const key = process.env.CB_KEY || '';
if (!key) {
  console.error('CB_KEY is missing; refusing to build a site that cannot log in.');
  process.exit(1);
}
fs.writeFileSync(path.join(OUT, 'config.local.js'),
  `window.ORDERFLOW_CONFIG = Object.assign({}, window.ORDERFLOW_CONFIG, {\n  mode: 'cloudbase',\n  envId: '${env}',\n  region: '${region}',\n  accessKey: '${key}'\n});\n`);

const indexPath = path.join(OUT, 'index.html');
let html = fs.readFileSync(indexPath, 'utf8');
for (const asset of ASSETS) {
  const re = new RegExp('(src|href)="' + asset.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '"', 'g');
  html = html.replace(re, `$1="${asset}?v=${version}"`);
}
html = html.replace('</head>', `<meta name="orderflow-build" content="${version}"></head>`);
fs.writeFileSync(indexPath, html);
fs.writeFileSync(path.join(OUT, '.nojekyll'), '');
console.log('built public/ version', version);
