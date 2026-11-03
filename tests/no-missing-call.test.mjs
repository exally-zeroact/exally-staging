/* no-missing-call.test.mjs — ★呼んでいるのに 無い 働き★を 見つける 見張り 2026-08-30
 *
 *  ★なぜ 要るか（08-30 に 実際に 起きた）★
 *    「改ページを足す」の 中で `colName(c)` を 呼んだが、うちの 名前は `colLetter`
 *    ＝★colName は どこにも 無い★。
 *    ・字を 見るだけの 試験は ★緑のまま★（`function 改ページを足す(` は 在るので）
 *    ・実ブラウザで 押して はじめて `ReferenceError: colName is not defined` が 出た
 *    ⇒ ★押さないと 分からない★のでは 遅い。★機械で 数えて 赤に する★。
 *
 *  ★数え方★
 *    book.html の 中の 台本から
 *      ・作っている 名前 … `function 名(` / `var 名 =` / `let` / `const` / 引数
 *      ・呼んでいる 名前 … `名(` の 形
 *    を 取り、★どこにも 無い 名前★を 探す。
 *    ★日本語も 英語も 両方 見る★（colName は 英語だったので 英語も 見ないと 意味が 無い）。
 *    「作っている」に 数える物＝book.html の 中／lib・js の *.js／exally-formula.js／
 *      ★ブラウザが はじめから 持っている 名前の 一覧★（下に 書いてある）。
 *
 *  走らせ方: node tests/no-missing-call.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const 壊す = process.argv.includes('--self-test');
let 緑 = 0, 赤 = 0;
const ok = (名, 条件, 添え) => {
  if (条件) { 緑++; console.log('  ok   ' + 名); }
  else { 赤++; console.log('  ★NG★ ' + 名 + (添え !== undefined ? '\n         ' + 添え : '')); }
};

const 日本語 = /[぀-ヿ一-鿿]/;
const 名前の形 = '[A-Za-z_$\\u3040-\\u30FF\\u4E00-\\u9FFF][\\w$\\u3040-\\u30FF\\u4E00-\\u9FFF]*';

/** 台本（<script> の 中で src の 無い 物）だけ 取り出す */
function 台本たち(html) {
  const 出 = [];
  const re = /<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/gi;
  let m;
  while ((m = re.exec(html)) !== null) if (m[1].trim()) 出.push(m[1]);
  return 出;
}

/** 字の中（'…' "…" `…`）と 説明（// と /* *\/）を 消す＝そこの 名前は 呼んでいない */
function 中身だけ(コード) {
  let 出 = '';
  let i = 0;
  const n = コード.length;
  while (i < n) {
    const c = コード[i];
    if (c === '/' && コード[i + 1] === '/') { while (i < n && コード[i] !== '\n') i++; continue; }
    if (c === '/' && コード[i + 1] === '*') {
      i += 2;
      while (i < n && !(コード[i] === '*' && コード[i + 1] === '/')) i++;
      i += 2; continue;
    }
    /* ★正規表現（/…/）も 飛ばす★＝中に ' や " が 入っていると
       字の 始まりだと 思い込んで ★そこから 先を 丸ごと 食べてしまう★。
       （08-30 実測＝これで 6個の 在る 働きを「無い」と 言っていた） */
    if (c === '/') {
      let j = 出.length - 1;
      while (j >= 0 && /\s/.test(出[j])) j--;
      const 前 = j >= 0 ? 出[j] : '(';
      if ('(,=:[!&|?{};+-*%~^<>'.indexOf(前) >= 0 || j < 0) {
        i++;
        let 角 = false;
        while (i < n) {
          const d = コード[i];
          if (d === '\\\\') { i += 2; continue; }
          if (d === '[') 角 = true;
          else if (d === ']') 角 = false;
          else if (d === '/' && !角) break;
          else if (d === String.fromCharCode(10)) break;   /* 割り算だった＝行を またがない */
          i++;
        }
        i++;
        while (i < n && /[gimsuyd]/.test(コード[i])) i++;
        出 += '/x/'; continue;
      }
    }
    if (c === '"' || c === "'" || c === '`') {
      const 閉 = c; i++;
      while (i < n && コード[i] !== 閉) { if (コード[i] === '\\') i++; i++; }
      i++; 出 += '""'; continue;
    }
    出 += c; i++;
  }
  return 出;
}

function 作っている名前(コード) {
  const 出 = new Set();
  let m;
  const 足す = (re, 番) => { const r = new RegExp(re, 'g'); while ((m = r.exec(コード)) !== null) 出.add(m[番]); };
  足す('function\\s+(' + 名前の形 + ')\\s*\\(', 1);
  足す('(?:var|let|const)\\s+(' + 名前の形 + ')\\s*=', 1);
  足す('(' + 名前の形 + ')\\s*[:=]\\s*function', 1);
  足す('(' + 名前の形 + ')\\s*=\\s*\\([^)]*\\)\\s*=>', 1);
  /* 引数（function 名(あ, い) と (あ, い) => の 中） */
  const 引 = new RegExp('function[^(]*\\(([^)]*)\\)', 'g');
  while ((m = 引.exec(コード)) !== null) {
    m[1].split(',').forEach(s => { const t = s.trim().split('=')[0].trim(); if (t) 出.add(t); });
  }
  /* for (var 名 of …) / catch (名) */
  足す('for\\s*\\(\\s*(?:var|let|const)\\s+(' + 名前の形 + ')', 1);
  足す('catch\\s*\\(\\s*(' + 名前の形 + ')', 1);
  /* window.名 = … （showToast は notify の 別名として こう 作っている） */
  足す('window\\.(' + 名前の形 + ')\\s*=', 1);
  /* 束の 中の 名前（{ 名: … }）＝物の 持ち物なので 呼ぶ 相手には ならないが 数えておく */
  足す('\\b(' + 名前の形 + ')\\s*:', 1);
  return 出;
}

function 呼んでいる名前(コード) {
  const 出 = new Map();     /* 名 → 何回 */
  const re = new RegExp('(^|[^\\w$.\\u3040-\\u30FF\\u4E00-\\u9FFF])(' + 名前の形 + ')\\s*\\(', 'g');
  let m;
  while ((m = re.exec(コード)) !== null) {
    const 名 = m[2];
    if (['if', 'for', 'while', 'switch', 'catch', 'function', 'return', 'typeof',
      'new', 'delete', 'void', 'in', 'of', 'do', 'else', 'try', 'case'].indexOf(名) >= 0) continue;
    出.set(名, (出.get(名) || 0) + 1);
  }
  return 出;
}

/** ★別の ファイルが 作っている 名前★
 *  book.html は lib/*.js ・ js/*.js ・ exally-formula.js を 読み込んでいる。
 *  そこで 作られた 名前は「無い」では ない ⇒ ★読み込む 物は 全部 見る★。 */
function 部品が出している名前() {
  const 出 = new Set();
  const 相手 = [];
  for (const d of ['lib', 'js']) {
    const dir = path.join(ROOT, d);
    if (!fs.existsSync(dir)) continue;
    for (const f of fs.readdirSync(dir)) if (f.endsWith('.js')) 相手.push(path.join(dir, f));
  }
  for (const f of ['exally-formula.js']) {
    const p = path.join(ROOT, f);
    if (fs.existsSync(p)) 相手.push(p);
  }
  for (const p of 相手) {
    const t = fs.readFileSync(p, 'utf8');
    let m;
    /* ★外から 呼べる 物だけ★＝root./window./self. に 付けた 物と
       ★行の 先頭（字下げなし）の function★。
       lib/*.js は かたまり（IIFE）の 中で 字下げして 作るので
       ★中の function は 外から 呼べない★＝数に 入れない。
       （入れてしまうと colName の ような 事故を ★見逃す★） */
    for (const 形 of ['root\\.', 'window\\.', 'self\\.']) {
      const re = new RegExp(形 + '(' + 名前の形 + ')\\s*=', 'g');
      while ((m = re.exec(t)) !== null) 出.add(m[1]);
    }
    const re3 = new RegExp('^function\\s+(' + 名前の形 + ')\\s*\\(', 'gm');
    while ((m = re3.exec(t)) !== null) 出.add(m[1]);
  }
  return 出;
}

/** ★ブラウザと JavaScript が はじめから 持っている 名前★
 *  （ここに 無い 英語の 名前を 呼んでいたら ★どこにも 無い★＝赤に する）
 *  ※ 増やす時は ★本当に ブラウザが 持っている物だけ★ 足す。
 *    「赤を 消す為に 足す」のは ★見張りを 殺す★ 事に なる。 */
const もとから在る名前 = new Set([
  'Array', 'Object', 'String', 'Number', 'Boolean', 'Symbol', 'BigInt',
  'Date', 'RegExp', 'Error', 'TypeError', 'RangeError', 'Function', 'Promise',
  'Map', 'Set', 'WeakMap', 'WeakSet', 'Proxy', 'Reflect', 'JSON', 'Math',
  'parseInt', 'parseFloat', 'isNaN', 'isFinite', 'encodeURIComponent',
  'decodeURIComponent', 'encodeURI', 'decodeURI', 'eval', 'structuredClone',
  'setTimeout', 'clearTimeout', 'setInterval', 'clearInterval',
  'requestAnimationFrame', 'cancelAnimationFrame', 'queueMicrotask',
  'fetch', 'Blob', 'File', 'FileReader', 'FormData', 'Headers', 'Request',
  'Response', 'URL', 'URLSearchParams', 'AbortController', 'TextEncoder',
  'TextDecoder', 'BroadcastChannel', 'MessageChannel', 'Worker',
  'Uint8Array', 'Uint16Array', 'Uint32Array', 'Int8Array', 'Int16Array',
  'Int32Array', 'Float32Array', 'Float64Array', 'ArrayBuffer', 'DataView',
  'Image', 'Audio', 'Event', 'CustomEvent', 'MouseEvent', 'KeyboardEvent',
  'MutationObserver', 'ResizeObserver', 'IntersectionObserver',
  'DOMParser', 'XMLSerializer', 'XMLHttpRequest', 'Intl', 'Notification',
  'alert', 'confirm', 'prompt', 'atob', 'btoa', 'print', 'open', 'close',
  'getComputedStyle', 'matchMedia', 'scrollTo', 'scrollBy', 'focus', 'blur',
  'require', 'importScripts', 'HyperFormula', 'supabase',
]);

function 調べる(html名) {
  const html = fs.readFileSync(path.join(ROOT, html名), 'utf8');
  const 作 = 部品が出している名前();
  const 呼 = new Map();
  for (const t of 台本たち(html)) {
    const c = 中身だけ(t);
    作っている名前(c).forEach(v => 作.add(v));
    呼んでいる名前(c).forEach((n, k) => 呼.set(k, (呼.get(k) || 0) + n));
  }
  const 無い = [];
  呼.forEach((回, 名) => {
    if (作.has(名)) return;
    if (もとから在る名前.has(名)) return;   /* ブラウザが 持っている物 */
    無い.push(名 + '（' + 回 + '回）');
  });
  return { 無い: 無い, 呼んだ数: 呼.size, 作った数: 作.size };
}

console.log('\n[① 数え方が 空振りしていない]');
{
  const r = 調べる('book.html');
  ok('★呼んでいる 名前を 100個 以上 見つけている★', r.呼んだ数 > 100, String(r.呼んだ数));
  ok('★作っている 名前を 300個 以上 見つけている★', r.作った数 > 300, String(r.作った数));
}

console.log('\n[② ★呼んでいるのに 無い 働きが 0個★]');
/* ★repo の 画面ぜんぶ★（kyuyo/ は Rakually の 持ち物なので 触らない＝見るだけ） */
const 画面たち = fs.readdirSync(ROOT).filter(f => f.endsWith('.html'));
for (const f of 画面たち) {
  if (!fs.existsSync(path.join(ROOT, f))) continue;
  const r = 調べる(f);
  ok(f + ' … 無い 働き 0個', r.無い.length === 0, r.無い.join(' / '));
}

console.log('\n[③ 数え方じたいが 正しいか（わざと 作った 例で 試す）]');
{
  const にせ = '<script>function あ(){ return 無い働き(1); }</script>';
  const p = path.join(ROOT, 'tests', '_tmp-missing-call.html');
  fs.writeFileSync(p, にせ, 'utf8');
  const r = 調べる(path.join('tests', '_tmp-missing-call.html'));
  fs.unlinkSync(p);
  ok('★無い 働きを ちゃんと 見つける★', r.無い.length === 1 && /無い働き/.test(r.無い[0]),
    JSON.stringify(r.無い));
}
{
  const よい = '<script>function 足す(a){ return a; } function あ(){ return 足す(1); }</script>';
  const p = path.join(ROOT, 'tests', '_tmp-ok-call.html');
  fs.writeFileSync(p, よい, 'utf8');
  const r = 調べる(path.join('tests', '_tmp-ok-call.html'));
  fs.unlinkSync(p);
  ok('★在る 働きを 無いと 言わない★', r.無い.length === 0, JSON.stringify(r.無い));
}
{
  const 字 = '<script>function あ(){ return "無い働き(1)"; }</script>';
  const p = path.join(ROOT, 'tests', '_tmp-str-call.html');
  fs.writeFileSync(p, 字, 'utf8');
  const r = 調べる(path.join('tests', '_tmp-str-call.html'));
  fs.unlinkSync(p);
  ok('★字の 中は 呼んでいると 数えない★', r.無い.length === 0, JSON.stringify(r.無い));
}
{
  const 説 = '<script>function あ(){ /* 無い働き(1) */ return 1; }</script>';
  const p = path.join(ROOT, 'tests', '_tmp-cmt-call.html');
  fs.writeFileSync(p, 説, 'utf8');
  const r = 調べる(path.join('tests', '_tmp-cmt-call.html'));
  fs.unlinkSync(p);
  ok('★説明の 中も 数えない★', r.無い.length === 0, JSON.stringify(r.無い));
}

console.log('\nno-missing-call: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝わざと 壊して 赤に なるか★');
  let 素通り = 0;
  const にせ = '<script>function あ(){ return 存在しない働き(1); }</script>';
  const p = path.join(ROOT, 'tests', '_tmp-self.html');
  fs.writeFileSync(p, にせ, 'utf8');
  const r = 調べる(path.join('tests', '_tmp-self.html'));
  fs.unlinkSync(p);
  if (!r.無い.length) { 素通り++; console.log('  ★素通り★ 無い 働きを 見逃した'); }
  else console.log('  ok   無い 働きを 見つけた … ' + r.無い[0]);
  /* ★08-30 の 本物の 事故を そのまま 再現する★ */
  const 事故 = '<script>function colLetter(c){return c;}\n'
    + 'function 改ページを足す(){ return colName(1) + colLetter(1); }</script>';
  const p2 = path.join(ROOT, 'tests', '_tmp-self2.html');
  fs.writeFileSync(p2, 事故, 'utf8');
  const r2 = 調べる(path.join('tests', '_tmp-self2.html'));
  fs.unlinkSync(p2);
  if (!r2.無い.some(v => /colName/.test(v))) {
    素通り++; console.log('  ★素通り★ 08-30 の colName 事故を 見逃した');
  } else console.log('  ok   08-30 の colName 事故を 見つけた … ' + r2.無い.join(' / '));
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
process.exit(赤 ? 1 : 0);
