/* kakushi-shirushi.test.mjs — ★隠した行・列の 並びの 覚えが 古いまま 使われないか★（2026-11-07）
 *
 *  前は「隠した物が 変わったか」を 鍵の 数と 和だけで 見て いた（_隠しの印・08-28 から）
 *  ＝{1,4} と {2,3} は 数2・和5 で 同じ 印 ⇒ 絞り込みで 隠す 行が 1・4 から 2・3 に 変わっても
 *    古い 並び [1,4] の まま rowY を 数え、行が 違う 所に 描かれた
 *  ⇒ 前の 鍵を そのまま 覚えて 1つずつ 比べる（_隠しが同じ）。違えば 作り直す
 *
 *  ここで 見る 事：
 *   ① {1,4}→{2,3}（行・絞り込み）／{0,5}→{2,3}（列・切り替え）で 新しい 並びが 返る
 *   ② 決まった 乱数で 5000 回 隠す物を 変え、毎回 1から 数えた 答えと 全部 同じ
 *   ③ 偽（0・false）は 隠していない／隠す物が 同じなら 作り直さない（並びの 箱が 同じ 物）
 *
 *  使い方: node tests/kakushi-shirushi.test.mjs
 *          node tests/kakushi-shirushi.test.mjs --self-test   （前の 数と和の 印に 戻すと 赤に なるか）
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

const 名たち = ['_隠しが同じ', '_隠しの並び', '_隠れた数', '_隠れ行の並び', '_隠れ列の並び'];

/* ★前の 印（数と和）★＝壊しの 試験で これに 戻すと 赤に なる事を 見る */
const 前の並び = `
function _隠しの並び(どっち, たち){
  var 印 = どっち + '|' + (sheets[activeSheet].name||'') + '|';
  for(var i=0;i<たち.length;i++){
    var o = たち[i] || {}; var n = 0, 和 = 0;
    for(var k in o){ if(Object.prototype.hasOwnProperty.call(o,k) && o[k]){ n++; 和 += (+k)||0; } }
    印 += n + ':' + 和 + '/';
  }
  var 箱 = _隠し[どっち];
  if(箱.印 === 印) return 箱.並び;
  var 印付き = {};
  for(var i=0;i<たち.length;i++){ var o = たち[i] || {};
    for(var k in o){ if(Object.prototype.hasOwnProperty.call(o,k) && o[k]) 印付き[+k] = 1; } }
  var 並び = Object.keys(印付き).map(Number).sort(function(a,b){ return a-b; });
  箱.印 = 印; 箱.並び = 並び; return 並び;
}`;

function 台を組む(src, 壊す) {
  let 関数 = 名たち.map((n) => 切る(src, n)).join('\n');
  if (壊す) 関数 = 関数.replace(切る(src, '_隠しの並び'), 前の並び);
  const 素 = [
    一行(src, /var _隠し = \{[^\n]*\};/, '_隠し'),
    一行(src, /var _足し控え = \{[^}]*\};/, '_足し控え'),
    'var sheets = [], activeSheet = 0;',
    関数,
  ].join('\n');
  const 箱 = {};
  vm.createContext(箱);
  vm.runInContext(素, 箱);
  return 箱;
}

function 乱数(種) { let s = 種 >>> 0; return () => { s = (s * 1664525 + 1013904223) >>> 0; return s / 4294967296; }; }
function 本当(たち) {
  const s = new Set();
  for (const o of たち) for (const k of Object.keys(o || {})) if (o[k]) s.add(+k);
  return [...s].sort((a, b) => a - b);
}
const 同じ = (a, b) => a.length === b.length && a.every((v, i) => v === b[i]);

function 走る(壊す) {
  const 台 = 台を組む(本, 壊す);
  let 赤 = 0, 緑 = 0;
  const 見る = (名, 出, 正) => {
    if (同じ(Array.from(出), 正)) 緑++;
    else { 赤++; if (赤 <= 5) console.log('  ✗ ' + 名 + ' 出=[' + Array.from(出).slice(0, 10) + '] 正=[' + 正.slice(0, 10) + ']'); }
  };
  const 板 = { name: 'S', hiddenRows: {}, hiddenCols: {}, kirikaeRows: {}, kirikaeCols: {}, filterHidden: {} };
  台.sheets.push(板);

  /* ① 数も和も 同じで 中身が 違う */
  板.filterHidden = { 1: true, 4: true };
  見る('① 行 {1,4}', 台._隠れ行の並び(), [1, 4]);
  板.filterHidden = { 2: true, 3: true };
  見る('① 行 {1,4}→{2,3}', 台._隠れ行の並び(), [2, 3]);
  板.kirikaeCols = { 0: 1, 5: 1 };
  見る('① 列 {0,5}', 台._隠れ列の並び(), [0, 5]);
  板.kirikaeCols = { 2: 1, 3: 1 };
  見る('① 列 {0,5}→{2,3}', 台._隠れ列の並び(), [2, 3]);
  /* 同じ 箱の 中を 書き換える（物は 同じ・中身だけ 変わる） */
  delete 板.hiddenRows[9]; 板.hiddenRows[7] = true; 板.hiddenRows[13] = true;
  見る('① 行 手で隠す 足し', 台._隠れ行の並び(), [2, 3, 7, 13]);
  delete 板.hiddenRows[7]; delete 板.hiddenRows[13]; 板.hiddenRows[10] = true; 板.hiddenRows[10.5] = false; 板.hiddenRows[10] = true;
  板.hiddenRows[8] = true; delete 板.hiddenRows[10]; 板.hiddenRows[12] = true;
  見る('① 行 7,13→8,12（数2・和20 同じ）', 台._隠れ行の並び(), [2, 3, 8, 12]);

  /* ③ 偽は 隠していない・同じなら 作り直さない */
  板.hiddenRows = { 8: true, 12: false, 20: 0 };
  見る('③ 偽は数えない', 台._隠れ行の並び(), [2, 3, 8]);
  const 一回目 = 台._隠れ行の並び();
  const 二回目 = 台._隠れ行の並び();
  if (壊す || 一回目 === 二回目) 緑++; else { 赤++; console.log('  ✗ ③ 同じなのに 作り直した'); }

  /* ② 乱数で 5000 回 */
  const R = 乱数(20261107);
  const 行の箱 = ['hiddenRows', 'kirikaeRows', 'filterHidden'], 列の箱 = ['hiddenCols', 'kirikaeCols'];
  for (const k of [...行の箱, ...列の箱]) 板[k] = {};
  for (let t = 0; t < 5000; t++) {
    const 行 = R() < 0.6;
    const 名 = (行 ? 行の箱 : 列の箱)[Math.floor(R() * (行 ? 3 : 2))];
    const o = 板[名], u = R();
    if (u < 0.1) 板[名] = {};
    else if (u < 0.2) { /* 同じ 数・同じ 和の 入れ替え（a,d → b,c） */
      const a = Math.floor(R() * 20), d = a + 3 + Math.floor(R() * 5);
      delete o[a]; delete o[d]; o[a + 1] = true; o[d - 1] = true;
    } else if (u < 0.6) o[Math.floor(R() * 30)] = R() < 0.85 ? true : 0;
    else delete o[Math.floor(R() * 30)];
    if (R() < 0.05) 台.activeSheet = 0;
    const 正 = 本当(行 ? 行の箱.map((k) => 板[k]) : 列の箱.map((k) => 板[k]));
    見る('② ' + t + ' ' + 名, 行 ? 台._隠れ行の並び() : 台._隠れ列の並び(), 正);
  }
  return { 赤, 緑 };
}

if (process.argv.includes('--self-test')) {
  const r = 走る(true);
  console.log('壊し（前の 数と和の 印）: 赤 ' + r.赤 + ' / 緑 ' + r.緑);
  if (r.赤 === 0) { console.log('★壊しても 赤に ならない＝試験が 見ていない★'); process.exit(1); }
  console.log('OK（壊すと 赤）');
} else {
  const r = 走る(false);
  console.log('隠しの並び: 緑 ' + r.緑 + ' / 赤 ' + r.赤);
  process.exit(r.赤 ? 1 : 0);
}
