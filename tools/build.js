// src/index.html の /*INLINE:パス*/ をファイルの中身に、/*TAILWIND*/ を Tailwind で作ったCSSに置き換え、1つのHTMLにまとめる
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const root = path.join(__dirname, '..');
const css = execFileSync(path.join(root, 'node_modules/.bin/tailwindcss'),
  ['-c', path.join(root, 'tailwind.config.js'), '-i', path.join(root, 'src/tailwind.css'), '--minify'],
  { cwd: root, encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'] });
let html = fs.readFileSync(path.join(root, 'src/index.html'), 'utf8');
html = html.replace('/*TAILWIND*/', () => css);
html = html.replace(/\/\*INLINE:([^*]+)\*\//g, (_, p) => {
  const body = fs.readFileSync(path.join(root, p.trim()), 'utf8');
  return body.replace(/<\/script/gi, '<\\/script');
});
fs.mkdirSync(path.join(root, 'dist'), { recursive: true });
const out = path.join(root, 'dist', '保護者会日程作成ツール.html');
fs.writeFileSync(out, html);
console.log(out, (html.length / 1024).toFixed(0) + 'KB', 'css', (css.length / 1024).toFixed(0) + 'KB');
