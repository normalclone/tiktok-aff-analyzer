// Gộp 2 thư viện + app.js vào app.html -> 1 file HTML offline duy nhất.
// Chạy: node build/build.mjs
import { readFileSync, writeFileSync, statSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const here = dirname(fileURLToPath(import.meta.url));
const read = (p) => readFileSync(join(here, p), 'utf8');
// Tránh chuỗi "</script>" bên trong lib làm đóng thẻ <script> sớm
const safe = (s) => s.replaceAll('</script>', '<\\/script>');

let html = read('app.html');
for (const lib of ['lib/xlsx.full.min.js', 'lib/chart.umd.min.js']) {
  html = html.replace(`<script src="${lib}"></script>`, `<script>\n${safe(read(lib))}\n</script>`);
}
html = html.replace('<script src="app.js"></script>', `<script>\n${read('app.js')}\n</script>`);

const out = join(here, '..', 'TikTok_Affiliate_Analyzer.html');
writeFileSync(out, html);
console.log(`Đã build: ${out} (${(statSync(out).size / 1024 / 1024).toFixed(2)} MB)`);
