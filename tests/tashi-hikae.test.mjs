/* tashi-hikae.test.mjs — ★幅・高さを 書いた 列と 行の 足し算の 控え★（2026-11-07）
 *
 *  前は colX/rowY を 呼ぶ たびに colW/rowH の 鍵を 全部 回して いた
 *  ＝取り込んだ xlsx の 行の 高さを 読むと 鍵が 何千も 来る（方眼紙）
 *  実WebKit で 測った（直す前）：rowH の 鍵 5000 ＝ 描き 186ms・行の 境目を 探す 1924ms
 *  ⇒ book.html に「並べた 鍵＋前からの 和」の 控えを 入れた（描く・境目・位置→列行 の 間だけ）
 *
 *  ここで 見る 事：
 *   ① 控えあり（_控えて の 中）と 控えなし（外）で colX/rowY が ★全部 同じ 数★
 *      ＝隠し 3種・倍率 7つ・ずれ・小数の 高さ・遠い 行（100万）まで
 *   ② xToC/yToR が 控えなし（_xToCの中 を 直に 呼ぶ）と 同じ 答え
 *   ③ 境目を 探す 2つ が 前の 歩き方（0 から 歩く）と 同じ 答え
 *   ④ 控えは 外に 出たら 捨てる（深さ 0 で null）
 *
 *  使い方: node tests/tashi-hikae.test.mjs
 *          node tests/tashi-hikae.test.mjs --self-test   （わざと 壊して 赤に なるか）
 */
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const 本 = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf-8');

function 切る(src, 名) {
  const i = src.indexOf('function ' + 名 + '(');
  if (i < 0) throw new Error('★' + 名 + ' が 無い★');
  let 深 = 0;
  for (let k = src.indexOf('{', i); k < src.length; k++) {
    if (src[k] === '{') 深++;
    else if (src[k] === '}') { 深--; if (!深) return src.slice(i, k + 1); }
  }
  throw new Error('★' + 名 + ' を 切り出せない★');
}
function 一行(src, re, 名) {
  const m = re.exec(src);
  if (!m) throw new Error('★' + 名 + ' が 無い★');
  return m[0];
}

const 名たち = ['cW', 'rH', '_隠しが同じ', '_隠しの並び', '_隠れた数', '_隠れ行の並び', '_隠れ列の並び',
  '_控えて', '_足しの控え', '_控えの和', 'colX', 'rowY',
  'xToC', '_xToCの中', 'yToR', '_yToRの中',
  'getColResizeBorder', '_getColResizeBorderの中', 'getRowResizeBorder', '_getRowResizeBorderの中'];

/* ★前の 境目の 探し方（0 から 歩く）★＝比べる 相手（2026-11-07 直す前の book.html を そのまま 写した） */
const 前の境目 = `
function 前の列の境目(x, y, threshold){
  threshold = threshold || 4;
  if(y >= HDR_H || x < HDR_W) return -1;
  var cw = sheets[activeSheet].colW;
  var start = Object.keys(cw).length === 0
    ? Math.max(0, Math.floor((x + scrollLeft - HDR_W) / (COL_W * scale)) - 2)
    : 0;
  for(var c = start; c < COLS; c++){
    var bx = colX(c) + cW(c);
    if(bx < HDR_W - threshold) continue;
    if(bx > x + threshold + 100) break;
    if(Math.abs(x - bx) <= threshold) return c;
  }
  return -1;
}
function 前の行の境目(x, y, threshold){
  threshold = threshold || 4;
  if(x >= HDR_W || y < HDR_H) return -1;
  var rh = sheets[activeSheet].rowH;
  var start = Object.keys(rh).length === 0
    ? Math.max(0, Math.floor((y + scrollTop - HDR_H) / (ROW_H * scale)) - 2)
    : 0;
  for(var r = start; r < ROWS; r++){
    var by = rowY(r) + rH(r);
    if(by < HDR_H - threshold) continue;
    if(by > y + threshold + 100) break;
    if(Math.abs(y - by) <= threshold) return r;
  }
  return -1;
}`;

function 台を組む(src) {
  const 素 = [
    一行(src, /var ROW_H = \d+, COL_W = \d+, HDR_W = \d+, HDR_H = \d+, ROWS = \d+, COLS = \d+;/, '大きさ'),
    一行(src, /var _隠し = \{[^\n]*\};/, '_隠し'),
    一行(src, /var _足し控え = \{[^}]*\};/, '_足し控え'),
    'var sheets = [], activeSheet = 0, scale = 1, scrollTop = 0, scrollLeft = 0;',
  ].concat(名たち.map((n) => 切る(src, n))).join('\n') + 前の境目;
  const 箱 = {};
  vm.createContext(箱);
  vm.runInContext(素, 箱);
  return 箱;
}

/* 決まった 乱数（毎回 同じ 板） */
function 乱数(種) { let s = 種 >>> 0; return () => { s = (s * 1664525 + 1013904223) >>> 0; return s / 4294967296; }; }

function 板を作る(種) {
  const R = 乱数(種);
  const 板 = { name: '板' + 種, colW: {}, rowH: {}, hiddenRows: {}, hiddenCols: {}, kirikaeRows: {}, kirikaeCols: {},
    filterHidden: {}, freezeRow: 0, freezeCol: 0, 既定の列幅: 64 };
  for (let i = 0; i < 400; i++) { const r = Math.floor(R() * 3000); 板.rowH[r] = R() < 0.2 ? 1 + R() * 80 : Math.round(1 + R() * 80); }
  for (let i = 0; i < 150; i++) { const c = Math.floor(R() * 600); 板.colW[c] = Math.round(4 + R() * 200); }
  for (let i = 0; i < 120; i++) 板.hiddenRows[Math.floor(R() * 3000)] = true;
  for (let i = 0; i < 40; i++) 板.filterHidden[Math.floor(R() * 3000)] = true;
  for (let i = 0; i < 20; i++) 板.kirikaeRows[Math.floor(R() * 3000)] = true;
  for (let i = 0; i < 40; i++) 板.hiddenCols[Math.floor(R() * 600)] = true;
  for (let i = 0; i < 10; i++) 板.kirikaeCols[Math.floor(R() * 600)] = true;
  板.rowH[1048570] = 50; 板.colW[16380] = 30;   /* 遠い 所 */
  return 板;
}

const 倍率たち = [1, 0.5, 0.75, 1.25, 0.1, 2, 0.9];

function 見る(src) {
  const 違い = [];
  const t = 台を組む(src);
  let 数 = 0;
  for (const 種 of [1, 2]) {
    const 板 = 板を作る(種);
    t.sheets = [板]; t.activeSheet = 0;
    for (const sc of 倍率たち) {
      t.scale = sc;
      for (const [sT, sL] of [[0, 0], [12345, 3210]]) {
        t.scrollTop = sT; t.scrollLeft = sL;
        /* 行は 2つおき＋書いた 鍵と 隠した 行の 前後は 全部（境目で ずれが 出る 所） */
        const 行の組 = new Set(); for (let r = 0; r <= 3100; r += 2) 行の組.add(r);
        for (const o of [板.rowH, 板.hiddenRows, 板.filterHidden, 板.kirikaeRows]) for (const k in o) { const n = +k; for (const d of [0, 1, 2]) if (n + d < 1048576) 行の組.add(n + d); }
        行の組.add(1048575);
        const 行たち = [...行の組].sort((a, b) => a - b);
        const 列たち = []; for (let c = 0; c <= 650; c++) 列たち.push(c); 列たち.push(16380, 16381, 16383);
        const 外 = { 行: 行たち.map((r) => t.rowY(r)), 列: 列たち.map((c) => t.colX(c)) };
        const 中 = t._控えて(() => ({ 行: 行たち.map((r) => t.rowY(r)), 列: 列たち.map((c) => t.colX(c)) }), null, []);
        for (let i = 0; i < 行たち.length; i++) { 数++; if (外.行[i] !== 中.行[i]) 違い.push('rowY(' + 行たち[i] + ') 倍率' + sc + ' 外' + 外.行[i] + ' 中' + 中.行[i]); }
        for (let i = 0; i < 列たち.length; i++) { 数++; if (外.列[i] !== 中.列[i]) 違い.push('colX(' + 列たち[i] + ') 倍率' + sc + ' 外' + 外.列[i] + ' 中' + 中.列[i]); }
        /* ② 位置 → 列・行 */
        for (let p = 0; p < 900; p += 7) {
          数 += 2;
          const a = t.xToC(p), b = t._xToCの中(p);
          if (a !== b) 違い.push('xToC(' + p + ') 倍率' + sc + ' 控え' + a + ' 無し' + b);
          const c2 = t.yToR(p), d = t._yToRの中(p);
          if (c2 !== d) 違い.push('yToR(' + p + ') 倍率' + sc + ' 控え' + c2 + ' 無し' + d);
        }
        /* ③ 境目 */
        for (let p = 0; p < 700; p += 3) {
          数 += 2;
          /* 前の 歩き方も 控えの 中で 回す（控えの 中と 外が 同じ 数なのは ①で 見た）＝速さだけの ため */
          const a = t.getRowResizeBorder(10, p), b = t._控えて(t.前の行の境目, null, [10, p]);
          if (a !== b) 違い.push('行の境目 y=' + p + ' 倍率' + sc + ' 今' + a + ' 前' + b);
          const c3 = t.getColResizeBorder(p, 10), d = t._控えて(t.前の列の境目, null, [p, 10]);
          if (c3 !== d) 違い.push('列の境目 x=' + p + ' 倍率' + sc + ' 今' + c3 + ' 前' + d);
        }
        if (違い.length > 5) return { 違い, 数, t };
      }
    }
  }
  /* 書いた 鍵が 無い 板（速い 道）も 同じか */
  t.sheets = [{ name: '空', colW: {}, rowH: {}, hiddenRows: {}, hiddenCols: {} }]; t.scale = 1; t.scrollTop = 500; t.scrollLeft = 300;
  for (let p = 0; p < 700; p++) {
    数 += 2;
    if (t.getRowResizeBorder(10, p) !== t.前の行の境目(10, p)) 違い.push('空の板 行の境目 y=' + p);
    if (t.getColResizeBorder(p, 10) !== t.前の列の境目(p, 10)) 違い.push('空の板 列の境目 x=' + p);
  }
  return { 違い, 数, t };
}

let pass = 0, fail = 0;
const T = (n, f) => { try { f(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); } };

console.log('');
console.log('[tashi-hikae] ★幅・高さを 書いた 列と 行の 足し算の 控え★');

let 結果;
T('★★控えあり／なしで colX・rowY・xToC・yToR・境目 が 全部 同じ★★', () => {
  結果 = 見る(本);
  if (結果.違い.length) throw new Error('★' + 結果.違い.length + '本 違う★  ' + 結果.違い.slice(0, 5).join(' ／ '));
  console.log('      … ' + 結果.数 + ' 個 比べて 全部 同じ（板2枚 × 倍率7 × ずれ2）');
});

T('★控えは 外に 出たら 捨てる（古い 控えを 使わない）★', () => {
  const t = 結果 ? 結果.t : 台を組む(本);
  t.sheets = [板を作る(9)]; t.scale = 1; t.scrollTop = 0; t.scrollLeft = 0;
  t._控えて(() => { t.rowY(500); t.colX(300); }, null, []);
  const 箱 = vm.runInContext('_足し控え', t);
  if (箱.深さ !== 0 || 箱.行 !== null || 箱.列 !== null) throw new Error('★外に 出ても 控えが 残って いる★');
  /* 入れ子＋転けても 捨てる */
  try { t._控えて(() => t._控えて(() => { t.rowY(10); throw new Error('わざと'); }, null, []), null, []); } catch (e) { /* わざと */ }
  if (箱.深さ !== 0 || 箱.行 !== null) throw new Error('★転けた 後に 控えが 残って いる★');
  /* 控えの 外で 高さを 変えたら 次の 描きで 新しい 数 */
  const 前 = t._控えて(() => t.rowY(600), null, []);
  t.sheets[0].rowH[3] = (t.sheets[0].rowH[3] || 24) + 100;
  const 後 = t._控えて(() => t.rowY(600), null, []);
  if (後 - 前 !== 100) throw new Error('★高さを 変えても 数が 変わらない（' + 前 + '→' + 後 + '）★');
  /* 隠しも 同じ＝控えの 外で 隠したら 次の 描きで 詰まる */
  const 前2 = t._控えて(() => t.rowY(2000), null, []);
  let 隠す行 = 1500; while (t.sheets[0].hiddenRows[隠す行] || t.sheets[0].filterHidden[隠す行] || t.sheets[0].kirikaeRows[隠す行]) 隠す行++;
  const 高さ = t.rH(隠す行);
  t.sheets[0].hiddenRows[隠す行] = true;
  const 後2 = t._控えて(() => t.rowY(2000), null, []);
  if (Math.abs((前2 - 後2) - 高さ) > 1) throw new Error('★隠しても 詰まらない（' + 前2 + '→' + 後2 + '・高さ ' + 高さ + '）★');
  if (箱.隠行 !== null || 箱.隠列 !== null) throw new Error('★隠しの 並びの 控えが 残って いる★');
});

T('★描く・境目・位置→列行 の 5つ が 控えの 中で 回る★', () => {
  for (const n of ['_doRender', 'xToC', 'yToR', 'getColResizeBorder', 'getRowResizeBorder']) {
    const 中 = n.charAt(0) === '_' ? n + 'の中' : '_' + n + 'の中';
    const re = new RegExp('function ' + n + '\\([^)]*\\)\\{ return _控えて\\(' + 中 + ', this, arguments\\); \\}');
    if (!re.test(本)) throw new Error('★' + n + ' が 控えの 包みに 入って いない★');
  }
});

if (process.argv.includes('--self-test')) {
  console.log('');
  console.log('--self-test: ★わざと 壊して 赤に なるか★（book.html は 1字も 触りません）');
  const 壊し方 = [
    { 名: '和に 既定を 引かない（控えの 数が ずれる）', 前: '書いた[k] - 既定', 後: '書いた[k]' },
    { 名: '隠した 鍵も 足す', 前: 'if(!isNaN(kk) && 見える(kk) > 0)', 後: 'if(!isNaN(kk))' },
    { 名: '控えの 二分探索が n 以下を 数える', 前: 'if(並[m] < n) lo = m + 1;', 後: 'if(並[m] <= n) lo = m + 1;' },
    { 名: '行の 境目の 始めが 1つ 先', 前: 'if(rowY(m) + rH(m) < HDR_H - threshold) lo = m + 1; else hi = m; }\n  var start = lo;', 後: 'if(rowY(m) + rH(m) < HDR_H - threshold) lo = m + 1; else hi = m; }\n  var start = lo + 1;' },
    { 名: '列の 境目の 始めが 1つ 先', 前: 'if(colX(m) + cW(m) < HDR_W - threshold) lo = m + 1; else hi = m; }\n  var start = lo;', 後: 'if(colX(m) + cW(m) < HDR_W - threshold) lo = m + 1; else hi = m; }\n  var start = lo + 1;' },
    { 名: '外に 出ても 控えを 捨てない', 前: '_足し控え.列 = null; _足し控え.行 = null;', 後: '' },
  ];
  let 悪 = 0;
  for (const こ of 壊し方) {
    if (本.indexOf(こ.前) < 0) { 悪++; console.log('  ✗ ' + こ.名 + ' — ★壊す 所が 見つからない★（形が 変わった＝見張りを 直す）'); continue; }
    const 壊 = 本.replace(こ.前, こ.後);
    let 赤 = false;
    try {
      const r = 見る(壊);
      if (r.違い.length) 赤 = true;
      if (!赤) {
        const t = r.t; t.sheets = [板を作る(9)];
        t._控えて(() => { t.rowY(500); }, null, []);
        const 箱 = vm.runInContext('_足し控え', t);
        if (箱.行 !== null) 赤 = true;
        else {
          const 前 = t._控えて(() => t.rowY(600), null, []);
          t.sheets[0].rowH[3] = (t.sheets[0].rowH[3] || 24) + 100;
          if (t._控えて(() => t.rowY(600), null, []) - 前 !== 100) 赤 = true;
        }
      }
    } catch (e) { 赤 = true; }
    if (赤) { pass++; console.log('  ✓ ' + こ.名 + ' → 赤'); } else { 悪++; console.log('  ✗ ' + こ.名 + ' → ★緑のまま（見張りが 空振り）★'); }
  }
  fail += 悪;
}

console.log('');
console.log('  ' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
