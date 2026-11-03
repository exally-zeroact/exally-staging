/* hakaru-oboetara-hayaku-naru-ka.mjs
 *   -- ★「覚えて おけば 速く なる」を ★測って から★ 言う★（116）（2026-09-25）
 *
 *  ★★なぜ★★
 *    114（CPUの記録）... ★87.4% が `lib/shiki-hyou.js`★
 *    115（回数）...     `名から番地` ★141,920,112 回★ ／ ★別々の 字は 53,365 種類★
 *                       ＝★同じ 字を 平均 2,659回 読み直して います★
 *    ⇒★「覚えて おけば 効く はず」は ★まだ 見立てです★★
 *    ⇒★外して 測るまで 見立て★＝★実際に 覚えさせて 秒を 測ります★
 *
 *  ★★測る 3通り★★（★同じ 台・同じ 材料・同じ 道具★）
 *    ⓪★そのまま★            ... ★元の 字の まま★（★物差し★）
 *    ①★名から番地だけ 覚える★ ... 字 → 番地 を 覚える
 *    ②★番地から名も 覚える★   ... ①＋ 番地 → 字 も 覚える
 *    ⇒★①と②を 分けるのは 「どちらが 効いたか」を 言える 為★
 *
 *  ★★安全か（★先に 確かめました★）★★
 *    `名から番地` を 呼ぶ 所 ... ★16か所★
 *    その 返り物を 書き換えて いる 所 ... ★0件★（`.板=` `.行=` `.列=` で 探して 0）
 *    ⇒★覚えた 物を そのまま 返しても 安全です★
 *    ⇒★機械でも 毎回 確かめます★（0件で ない なら exit 5）
 *
 *  ★★この 道具は 直しません★★
 *    ＝★repo の `lib/shiki-hyou.js` は 1バイトも 触りません★
 *    ＝★配る 時だけ 字を 包みます★＝★手で 書き換えた 物では ありません★
 *    ⇒★直すのは Exally1 です★。★これは 「直したら どれだけ 縮むか」の 数だけを 出します★
 *
 *  ★★速さだけで 決めません★★
 *    ＝★3通りの 答え（マス数・式数・誤の数・数の数・字の数・数の 足し算）を 突き合わせます★
 *    ＝★1つでも 違えば その 手は 使えません★（exit 6）
 *
 *  ★★出さない 物★★ ... ★マスの 値／シートの 名／式の 字★
 *
 *  使い方: node docs/measured/hakaru-oboetara-hayaku-naru-ka.mjs "<xlsb>" [--秒 300]
 */
import fs from 'node:fs';
import path from 'node:path';
import http from 'node:http';
import crypto from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { borrow, launch } from '../../scripts/_borrow-playwright.mjs';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..', '..');
const 材料 = process.argv[2];
const 秒 = process.argv.includes('--秒') ? Number(process.argv[process.argv.indexOf('--秒') + 1]) : 300;
/* ★★何回ずつ 測るか★★（2026-09-25・追記）
     ★1回ずつでは 「揺れ」と 「効いた」が 分けられません★
     ＝1回目は ⓪62.1秒／①40.6秒／②42.8秒 でした。
       ⓪→① の 21.4秒は 大きい。★但し ①→② の 2.1秒は 揺れかも しれません★
     ⇒★揺れ幅（同じ 通りを 何回も 測った 時の 幅）を 先に 出します★
     ⇒★幅より 小さい 差を 「効いた」と 言っては いけません★ */
const 回 = process.argv.includes('--回') ? Number(process.argv[process.argv.indexOf('--回') + 1]) : 1;
if (!材料 || !fs.existsSync(材料)) { console.log('★材料が 在りません★'); process.exit(3); }
const 中 = fs.readFileSync(材料);
console.log('★材料★ ' + path.basename(材料) + '（' + 中.length.toLocaleString() + ' バイト）');
console.log('  sha256 ' + crypto.createHash('sha256').update(中).digest('hex'));

const 表の道 = path.join(ROOT, 'lib', 'shiki-hyou.js');
const 元字 = fs.readFileSync(表の道, 'utf8');
const 元印 = crypto.createHash('sha256').update(元字).digest('hex');

/* ── ★門①★ 包む 相手が 1つずつ しか 無いか ── */
const 印1 = 'function 名から番地(名) {';
const 印2 = 'function 番地から名(行, 列, 板) {';
for (const 組 of [['名から番地', 印1], ['番地から名', 印2]]) {
  const c = 元字.split(組[1]).length - 1;
  console.log('★包む 相手★ ' + 組[0].padEnd(6) + ' ... ファイルに ' + c + '個');
  if (c !== 1) { console.log('★★1つで は ありません＝包みません★★'); process.exit(4); }
}
/* ── ★門②★ 返り物を 書き換えて いないか ── */
const 書き換え = (元字.match(/\.(?:板|行|列)\s*=[^=]/g) || []).length;
console.log('★返り物を 書き換えて いる 所★ ... ' + 書き換え + '件');
if (書き換え !== 0) { console.log('★★書き換えて います＝覚えては いけません★★'); process.exit(5); }

function 包む(どれ) {
  let s = 元字;
  if (どれ >= 1) {
    s = s.replace(印1,
      'var __覚1 = new Map();\n'
      + '  function 名から番地(名) {\n'
      + '    var k = String(名); var v = __覚1.get(k);\n'
      + '    if (v !== undefined) return v;\n'
      + '    v = __名から番地の元(k); __覚1.set(k, v); return v;\n'
      + '  }\n'
      + '  function __名から番地の元(名) {');
  }
  if (どれ >= 2) {
    s = s.replace(印2,
      'var __覚2 = new Map();\n'
      + '  function 番地から名(行, 列, 板) {\n'
      + '    var k = 行 + "|" + 列 + "|" + 板; var v = __覚2.get(k);\n'
      + '    if (v !== undefined) return v;\n'
      + '    v = __番地から名の元(行, 列, 板); __覚2.set(k, v); return v;\n'
      + '  }\n'
      + '  function __番地から名の元(行, 列, 板) {');
  }
  return s;
}
/* ★包んだ 字が ★本当に 変わって いるか★ を 先に 見ます★
   ＝`replace` が 外れても 黙って 元の まま 走り ★「速く ならなかった」の 嘘★に なる為 */
for (let i = 1; i <= 2; i++) {
  const t = 包む(i);
  if (t === 元字 || t.length <= 元字.length) {
    console.log('★★包めて いません（' + i + '）★★'); process.exit(4);
  }
}

let 配る字 = 元字;
const 型 = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8', '.json': 'application/json', '.png': 'image/png',
  '.svg': 'image/svg+xml', '.ico': 'image/x-icon' };
const server = http.createServer((req, res) => {
  const u = decodeURIComponent(req.url.split('?')[0]);
  const p = path.join(ROOT, u === '/' ? 'book.html' : u.replace(/^\//, ''));
  if (!p.startsWith(ROOT) || !fs.existsSync(p) || fs.statSync(p).isDirectory()) { res.writeHead(404); res.end(); return; }
  if (p === 表の道) { res.writeHead(200, { 'Content-Type': 型['.js'] }); res.end(配る字); return; }
  res.writeHead(200, { 'Content-Type': 型[path.extname(p)] || 'application/octet-stream' });
  res.end(fs.readFileSync(p));
});
await new Promise((ok) => server.listen(0, '127.0.0.1', ok));
const もと = 'http://127.0.0.1:' + server.address().port;

const 札 = ['⓪そのまま', '①名から番地だけ 覚える', '②番地から名も 覚える'];
const 出た = [];
/* ★★順に 回さず 交互に 回します★★
     ＝★台が 温まる／冷える 向きが 1つの 通りに 偏らない 為★
     （⓪⓪⓪①①①②②② だと ⓪が 全部 「冷えた 回」に なりえます） */
const 順 = [];
for (let k = 0; k < 回; k++) for (let d = 0; d <= 2; d++) 順.push(d);
for (const どれ of 順) {
  配る字 = 包む(どれ);
  const ck = await borrow('oboetara', 'chromium');
  const browser = await launch('oboetara', ck, {}, 'chromium');
  const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
  await page.goto(もと + '/book.html', { waitUntil: 'load', timeout: 120000 });
  await page.evaluate(() => {
    document.body.classList.remove('exally-locked');
    const ov = document.getElementById('loginOv');
    if (ov) { ov.classList.remove('open'); ov.style.display = 'none'; }
    window.FileOut = { deliver: () => Promise.reject(new Error('★書き出しては いけません★')) };
  });
  const t0 = Date.now();
  await page.setInputFiles('#bookFileInput', 材料);
  let 開いた = 0;
  for (let t = 0; t < 秒; t += 2) {
    await new Promise((r) => setTimeout(r, 2000));
    const い = await page.evaluate(() => {
      const ss = window.sheets || [];
      let m = 0; ss.forEach((sh) => { m += Object.keys((sh && sh.data) || {}).length; });
      return m;
    }).catch(() => -1);
    if (い > 0) { 開いた = Date.now() - t0; break; }
  }
  /* ★★答えが 変わって いないかを 見ます★★（★速くても 答えが 違えば 失格★）
       ＝★値そのものは 出しません★＝★数だけ★ */
  const 中身 = await page.evaluate(() => {
    const ss = window.sheets || [];
    let マス = 0, 式 = 0, 誤 = 0, 数 = 0, 字 = 0, 足 = 0;
    ss.forEach((sh) => {
      const d = (sh && sh.data) || {};
      Object.keys(d).forEach((k) => {
        マス++;
        const c = d[k]; if (!c || typeof c !== 'object') return;
        if (c.f) 式++;
        const v = c.v;
        if (typeof v === 'string' && v.charAt(0) === '#') 誤++;
        else if (typeof v === 'number') { 数++; 足 += v; }
        else if (typeof v === 'string') 字++;
      });
    });
    return { マス: マス, 式: 式, 誤: 誤, 数: 数, 字: 字, 足: 足 };
  }).catch(() => null);
  出た.push({ 号: どれ, 札: 札[どれ], ms: 開いた, 中: 中身 });
  console.log('  ' + 札[どれ].padEnd(24) + ' ... ' + (開いた ? String(開いた).padStart(7) + ' ms' : '★開きません★'));
  await browser.close();
}
server.close();

console.log('');
console.log('★★結果★★（' + 回 + '回ずつ）');
const 束 = [0, 1, 2].map((d) => 出た.filter((x) => x.号 === d).map((x) => x.ms));
const 中央 = (a) => { const b = a.slice().sort((p, q) => p - q); return b[(b.length - 1) >> 1]; };
const 基 = 中央(束[0]) || 1;
console.log('  ' + '通り'.padEnd(24) + '中ほど'.padStart(10) + '一番速い'.padStart(10)
  + '一番遅い'.padStart(10) + '★幅★'.padStart(9) + '   ' + '物差し比'.padStart(8) + '   各回');
[0, 1, 2].forEach((d) => {
  const a = 束[d];
  const 小 = Math.min.apply(null, a), 大 = Math.max.apply(null, a), 中 = 中央(a);
  console.log('  ' + 札[d].padEnd(24) + String(中.toLocaleString()).padStart(10)
    + String(小.toLocaleString()).padStart(10) + String(大.toLocaleString()).padStart(10)
    + String((大 - 小).toLocaleString()).padStart(9)
    + '   ' + ((中 / 基 * 100).toFixed(0) + '%').padStart(8)
    + '   ' + a.join(' / '));
});
/* ★★差が 揺れより 大きいか★★ */
const 幅 = Math.max.apply(null, [0, 1, 2].map((d) => Math.max.apply(null, 束[d]) - Math.min.apply(null, 束[d])));
console.log('');
console.log('★★一番 大きい 揺れ幅 ... ' + 幅.toLocaleString() + ' ms★★');
console.log('  ⇒★この 幅より 小さい 差を 「効いた」と 言っては いけません★');
[[0, 1], [0, 2], [1, 2]].forEach((組) => {
  const 差 = 中央(束[組[0]]) - 中央(束[組[1]]);
  console.log('  ' + (札[組[0]] + ' → ' + 札[組[1]]).padEnd(46)
    + String(差.toLocaleString()).padStart(9) + ' ms   '
    + (Math.abs(差) > 幅 ? '★幅より 大きい＝差が 在ります★' : '★幅の 中＝差は 言えません★'));
});
console.log('');
console.log('★★答えは 変わって いないか★★（★速くても 答えが 違えば 失格★）');
console.log('  ' + '通り'.padEnd(24) + 'マス'.padStart(8) + '式'.padStart(8) + '誤'.padStart(7)
  + '数'.padStart(8) + '字'.padStart(8) + '   数の 足し算');
let 揃った = true;
出た.forEach((x) => {
  if (出た.filter((y) => y.号 === x.号).indexOf(x) !== 0 && 回 > 1) { /* 出すのは 1回目だけ */ }
  const c = x.中 || {};
  console.log('  ' + x.札.padEnd(24) + String(c.マス).padStart(8) + String(c.式).padStart(8)
    + String(c.誤).padStart(7) + String(c.数).padStart(8) + String(c.字).padStart(8)
    + '   ' + (typeof c.足 === 'number' ? c.足.toFixed(6) : '?'));
  const b = 出た[0].中 || {};
  if (!x.中 || c.マス !== b.マス || c.式 !== b.式 || c.誤 !== b.誤 || c.数 !== b.数
    || c.字 !== b.字 || Math.abs(c.足 - b.足) > 1e-6) 揃った = false;
});
console.log('  ⇒' + (揃った ? '★3通り とも 揃って います★' : '★★揃って いません＝この 手は 使えません★★'));
console.log('');
const 後印 = crypto.createHash('sha256').update(fs.readFileSync(表の道)).digest('hex');
console.log('★repo の `lib/shiki-hyou.js` は 触って いません★');
console.log('  （走らせる 前 ' + 元印.slice(0, 16) + ' ／ 後 ' + 後印.slice(0, 16) + '）');
if (元印 !== 後印) { console.log('★★触れて います★★'); process.exit(7); }
if (!揃った) process.exit(6);
