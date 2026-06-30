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
// Dùng HÀM thay thế (không phải chuỗi) để JS không diễn giải các mẫu "$&", "$`"...
// có sẵn trong code minified của thư viện (lỗi này từng chèn nhầm thẻ <script> vào giữa SheetJS).
for (const lib of ['lib/xlsx.full.min.js', 'lib/chart.umd.min.js']) {
  const inline = `<script>\n${safe(read(lib))}\n</script>`;
  html = html.replace(`<script src="${lib}"></script>`, () => inline);
}
const appInline = `<script>\n${read('app.js')}\n</script>`;
html = html.replace('<script src="app.js"></script>', () => appInline);

const out = join(here, '..', 'TikTok_Affiliate_Analyzer.html');
writeFileSync(out, html);
console.log(`Đã build: ${out} (${(statSync(out).size / 1024 / 1024).toFixed(2)} MB)`);
