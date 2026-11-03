/* hairanai.test.mjs — ★書き出す ファイルに 入らない 物を「言う」★（2026-09-06）
 *
 *  ★★何を 守るか★★
 *    2026-09-06 実測 … ★うちで 足した シートは 書き出すと 消える★
 *      場所 … `js/book-open.js` `saveXlsxLike()`
 *             `if (opened.sheetNames.indexOf(sh.name) < 0) return;`
 *    ★預かった ファイルは 1バイトも 壊していない★（08-30 実測ずみ）
 *    ⇒ 悪いのは ★黙って 消える★事＝画面は「値だけ 書き換えました」としか 言わない
 *    ⇒★数を 出して「入りません」と はっきり 言う★（指示役の 裁定・2026-09-06）
 *
 *  ★消す・止める 事は しない★＝★言うだけ★（zip へ 足すのは 別の 回）
 *
 *  使い方: node tests/hairanai.test.mjs
 *          node tests/hairanai.test.mjs --self-test
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(import.meta.url);
const H = require_(path.join(ROOT, 'lib/hairanai.js'));
let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };

if (process.argv.includes('--self-test')) {
  console.log('\n[hairanai --self-test] わざと壊して赤になるか');
  T('★何も 足していなければ ★何も 言わない★（うるさく しない）', () => {
    const r = H.入らない物を数える(['明細'], [{ name: '明細' }], {});
    if (r.件 !== 0) throw new Error('0件で ないと 出た: ' + JSON.stringify(r));
    if (H.言葉(r) !== '') throw new Error('★0件なのに 言葉が 出た★');
  });
  T('★★足した シートを 数える★★', () => {
    const r = H.入らない物を数える(['明細'], [{ name: '明細' }, { name: '集計' }], {});
    if (r.件 !== 1) throw new Error('数え違い: ' + r.件);
    if (H.言葉(r).indexOf('集計') < 0) throw new Error('★どのシートか 言っていない★');
  });
  T('★浮かぶ物も 数える', () => {
    const r = H.入らない物を数える(['明細'], [{ name: '明細' }], { '図形・グラフ・画像': 3, '付箋': 2 });
    if (r.件 !== 5) throw new Error('数え違い: ' + r.件);
  });
  T('★★「入りません」と はっきり 書く（「触っていません」では 通じない）★★', () => {
    const s = H.言葉(H.入らない物を数える(['明細'], [{ name: '明細' }, { name: '集計' }], {}));
    if (s.indexOf('入りません') < 0) throw new Error('★はっきり 言っていない★');
  });
  T('★★出来ない事の 横に 行き先が 在る（今／これから を 分ける）★★', () => {
    const s = H.言葉(H.入らない物を数える(['明細'], [{ name: '明細' }, { name: '集計' }], {}));
    if (s.indexOf('画面の 中だけに 残ります') < 0) throw new Error('★今 どうなるかを 書いていない★');
    if (s.indexOf('これから') < 0) throw new Error('★これからを 書いていない★');
    /* ★★「今 出来る 事」と 見出しを 付けて 中身が「まだ 有りません」は 嘘★★
       （2026-09-06 指示役＝★見出しが 約束して 中身が 裏切る★） */
    if (/今 出来る 事/.test(s)) throw new Error('★『今 出来る 事』と 言って 出来ない事を 書いている★');
  });
  T('★★お客さんに 出す 字に ★ を 使わない★★', () => {
    /* ★★ は ★私たちの 便りの 印★＝お客さんには ★壊れた 字★に 見える
       （2026-09-06 指示役／同じ日に Rakunally も 23件 直している） */
    const s = H.言葉(H.入らない物を数える(['明細'], [{ name: '明細' }, { name: '集計' }],
      { '図形・グラフ・画像': 2 }));
    const 数 = (s.match(/★/g) || []).length;
    if (数) throw new Error('★ が ' + 数 + '個 出ています＝お客さんには 壊れた字に 見えます');
    const t = H.一行(H.入らない物を数える(['明細'], [{ name: '明細' }, { name: '集計' }], {}));
    if ((t.match(/★/g) || []).length) throw new Error('一行の 方に ★ が 出ています');
  });
  T('★★短い（3行まで）★★', () => {
    /* ★お客さんは 書き出す 直前＝急いでいる★／★1行目だけで 分かる 形★ */
    const s = H.言葉(H.入らない物を数える(['明細'], [{ name: '明細' }, { name: '集計' }], {}));
    const 行 = s.split('\n').length;
    if (行 > 3) throw new Error('★' + 行 + '行 在ります（3行まで）★');
    if (s.split('\n')[0].indexOf('入りません') !== 0) throw new Error('★1行目が「入りません」で 始まっていない★');
  });
  T('★★無い 逃げ道を 書かない★★', () => {
    /* ★「新しいブックとして 書き出す」は ★口が 無い★（2026-09-06 実測）
       ⇒ 書いたら ★出来ない物を 案内する★事に なる */
    const s = H.言葉(H.入らない物を数える(['明細'], [{ name: '明細' }, { name: '集計' }], {}));
    if (/新しいブックとして 書き出す/.test(s)) throw new Error('★まだ 無い 口を 案内している★');
    if (s.indexOf('これから') < 0) throw new Error('★まだ 無い事を 書いていない★');
  });
  console.log('\n' + pass + ' passed, ' + fail + ' failed');
  process.exit(fail ? 1 : 0);
}

/* ══ 本番＝本物の book.html を 動かして 窓に 出るか 見る ══════ */
let JSDOM;
try { ({ JSDOM } = await import('jsdom')); }
catch { console.log('★jsdomが入っていません。この検証は飛ばせません（SKIPを緑と呼ばない）'); process.exit(1); }

console.log('\n[hairanai] 本物の 画面で「入りません」が 出るか');

const html = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
const CANVAS_STUB = `
(function(){
  var noop=function(){};
  var ctx=new Proxy({}, { get:function(t,k){
    if(k==='measureText') return function(){ return {width:40}; };
    if(k==='canvas') return {width:800,height:600};
    if(k==='getImageData') return function(){ return {data:[]}; };
    if(k==='createLinearGradient'||k==='createPattern') return function(){ return {addColorStop:noop}; };
    return noop;
  }});
  HTMLCanvasElement.prototype.getContext=function(){ return ctx; };
})();`;
const dom = new JSDOM(html.replace(/<script[\s\S]*?<\/script>/g, ''), {
  runScripts: 'dangerously', url: 'http://localhost/', pretendToBeVisual: true,
  beforeParse(w) {
    w.fetch = () => Promise.reject(new Error('no net'));
    w.scrollTo = () => {}; w.alert = () => {};
    w.matchMedia = w.matchMedia || (() => ({ matches: false, addListener() {}, removeListener() {}, addEventListener() {}, removeEventListener() {} }));
    w.requestAnimationFrame = (cb) => setTimeout(() => cb(Date.now()), 0);
    w.cancelAnimationFrame = () => {};
    w.eval(CANVAS_STUB);
  },
});
const win = dom.window, doc = win.document;
const inject = (code) => { const el = doc.createElement('script'); el.textContent = code; doc.body.appendChild(el); };
for (const src of [...html.matchAll(/<script src="([^"]+)"><\/script>/g)].map((m) => m[1].split('?')[0]).filter((s) => !/^https?:/.test(s))) {
  const p = path.join(ROOT, src);
  if (fs.existsSync(p)) { try { inject(fs.readFileSync(p, 'utf8')); } catch (e) { /* 続ける */ } }
}
for (const m of html.matchAll(/<script(?![^>]*\ssrc=)[^>]*>([\s\S]*?)<\/script>/g)) { try { inject(m[1]); } catch (e) { /* 続ける */ } }
try { doc.dispatchEvent(new win.Event('DOMContentLoaded', { bubbles: true })); } catch (e) { /* 続ける */ }
try { win.dispatchEvent(new win.Event('load')); } catch (e) { /* 続ける */ }

T('★本物の 画面が 立ち上がっている（立たないと 何を 測っても 嘘）', () => {
  if (typeof win._askBeforeWrite !== 'function') throw new Error('_askBeforeWrite が 無い');
  if (typeof win._浮かぶ物を数える !== 'function') throw new Error('数える 口が 無い');
  if (!win.Hairanai) throw new Error('lib/hairanai.js が 読まれていない');
});

/* ★後から 読み込む 部品を 先に 積む★（本番は _loadScript で 読む＝jsdom には 来ない）
   ⇒★積み忘れると _askBeforeWrite が 1行目で 止まり、窓が 空のまま★
     （2026-09-06 実際に 踏んだ＝「出ていない」と 出たが 私の 測り台が 足りなかった） */
for (const 後から of ['lib/diff-preview.js']) {
  inject(fs.readFileSync(path.join(ROOT, 後から), 'utf8'));
}
if (typeof win.DiffPreview === 'undefined') throw new Error('★DiffPreview を 積めていない★');

/* ★開いた ファイルの ふり★＝元は「明細」だけ／画面には「集計」も 在る */
/* ★BookOpen は 画面では 後から 読み込む★＝jsdom では 無い事が 在る
   ⇒★偽物で 代える（★開いた ファイルの ふり★）★／無いのに 素通りさせない */
const 元の = { name: 'sample.xlsx', sheetNames: ['明細'], base: {}, bytes: new win.Uint8Array(1) };
if (!win.BookOpen) win.BookOpen = {};
win.BookOpen.current = () => 元の;
win.BookOpen.isOpened = () => true;
win.sheets = [
  { name: '明細', data: {}, objects: [{ 名: '四角 1' }, { 名: '円 2' }], comments: { '0,0': { 字: 'めも' } } },
  { name: '集計', data: {} },
];

T('★★数える口が 実物の 入れ物を 数えている★★', () => {
  const n = win._浮かぶ物を数える();
  if (n['図形・グラフ・画像'] !== 2) throw new Error('図形の 数が 違う: ' + JSON.stringify(n));
  if (n['付箋'] !== 1) throw new Error('付箋の 数が 違う: ' + JSON.stringify(n));
});

const 出た = (() => {
  const plan = { total: 1, sheets: [{ name: '明細', count: 1, rows: [{ addr: 'A1', before: 'あ', after: 'い', byUser: true }], more: 0 }] };
  try { win._askBeforeWrite(plan, 'sample.xlsx'); }
  catch (e) { console.log('  ★窓が 開けなかった★: ' + (e && e.message)); }
  const el = doc.getElementById('diffBody');
  return el ? el.innerHTML : '★diffBody が 無い★';
})();

T('★★窓に「入りません」が 出ている★★', () => {
  if (出た.indexOf('入りません') < 0) {
    throw new Error('★出ていない★＝黙って 消えます\n     出た字: ' + 出た.slice(0, 200));
  }
});

T('★★数が 出ている（新しいシート 1枚・図形 2個・付箋 1個）★★', () => {
  for (const 要る of ['新しいシート 1枚', '集計', '図形・グラフ・画像 2個', '付箋 1個']) {
    if (出た.indexOf(要る) < 0) throw new Error('出ていない: ' + 要る);
  }
});

T('★字が 背景に 埋もれない（箱に 色と 字の 色が 別で 在る）', () => {
  if (出た.indexOf('class="hairanai"') < 0) throw new Error('箱に 入っていない');
  if (html.indexOf('.hairanai{') < 0) throw new Error('見た目の 決まりが 無い');
  const css = html.slice(html.indexOf('.hairanai{'), html.indexOf('.hairanai{') + 260);
  if (css.indexOf('background') < 0 || css.indexOf('color') < 0) throw new Error('背景か 字の 色が 無い');
});

/* ★★逆向き＝足していない時は 出さない★★（うるさい 知らせは 読まれなくなる） */
T('★★何も 足していなければ 出さない★★', () => {
  win.sheets = [{ name: '明細', data: {} }];
  const plan = { total: 1, sheets: [{ name: '明細', count: 1, rows: [], more: 0 }] };
  win._askBeforeWrite(plan, 'sample.xlsx');
  const s = doc.getElementById('diffBody').innerHTML;
  if (s.indexOf('入りません') >= 0) throw new Error('★足していないのに 出た★');
});

/* ══ ★★形で 言い分ける★★ ══（2026-09-21）
     ★★なぜ 要るか★★
       2026-09-21 に `.xlsx` / `.xlsm` で ★足した 板が 入る ように なりました★。
       ⇒この 台が そのままだと ★入る 物を 「入りません」と 言います★
       ⇒★★逆向きの 嘘です★★（黙って 消えるのと 同じくらい 悪い）
     ★★`.xlsb` は まだ です★★＝★司さんの 実物は この 形★
       ＝`saveXlsb()` は 別の 道（`workbook.bin` が 2進）
     ★形が 分からない 時は 「入らない」側★
       ＝★黙って 消えるより 出過ぎる 方が まし★ */
T('★★`.xlsx` / `.xlsm` では 板を 「入りません」と 言わない★★', () => {
  for (const k of ['xlsx', 'xlsm']) {
    const r = H.入らない物を数える(['Sheet1'], [{ name: 'Sheet1' }, { name: 'Sheet2' }], {}, k);
    if (r.件 !== 0) {
      throw new Error('★' + k + ' で ' + r.件 + '件 出ました★ ' + JSON.stringify(r.中身)
        + '／★2026-09-21 に 入る ように なって います★');
    }
  }
});

/* ══ ★★`.xlsb` は 「板は 入る／字の マスは 入らない」★★ ══（2026-09-22）
     ★★この 門は 1日で 2回 書き換わりました★★
       09-21 ... 「`.xlsb` では 今も 『入りません』と 言う」（★まだ 直して いない★）
       09-22 ... ★板が 入る ように なりました★ ⇒ ★上の 門が 正しく 赤に しました★
     ⇒★門が 古く なった のを 門 自身が 教えました★
     ★★但し 字の マスは 本当に 入りません★★
       ＝`.xlsb` の 字は `sharedStrings.bin` の 番号を 指します
       ＝足すと ★元の 板が 指す 番号が ずれる 恐れ★
       ⇒★まとめて 「入りません」と 言うと 逆向きの 嘘★ に なります */
T('★★`.xlsb` で 数だけなら 「入りません」と 言わない★★', () => {
  const r = H.入らない物を数える(['Sheet1'],
    [{ name: 'Sheet1', data: {} }, { name: 'Sheet2', data: { '0,0': { v: 123 } } }], {}, 'xlsb');
  if (r.件 !== 0) {
    throw new Error('★xlsb（数だけ）で ' + r.件 + '件★ ' + JSON.stringify(r.中身)
      + '／★2026-09-22 に 板が 入る ように なって います★');
  }
});

T('★★`.xlsb` で 字の マスが 在れば その 数を 言う★★', () => {
  const r = H.入らない物を数える(['Sheet1'], [{ name: 'Sheet1', data: {} },
    { name: 'Sheet2', data: { '0,0': { v: 123 }, '0,1': { v: 'abc' }, '1,0': { v: 'xyz' } } }],
  {}, 'xlsb');
  if (r.件 !== 2) throw new Error('★' + r.件 + '件★（2件 の はず）' + JSON.stringify(r.中身));
  if (String(r.中身[0].何).indexOf('字') < 0) {
    throw new Error('★中身が 違う★ ' + JSON.stringify(r.中身));
  }
});

T('★★`.xlsx` は 字の マスが 在っても 言わない★★（★そちらは 入る★）', () => {
  const r = H.入らない物を数える(['Sheet1'], [{ name: 'Sheet1', data: {} },
    { name: 'Sheet2', data: { '0,1': { v: 'abc' } } }], {}, 'xlsx');
  if (r.件 !== 0) throw new Error('★xlsx で ' + r.件 + '件★ ' + JSON.stringify(r.中身));
});

T('★`.xls` は 今も 板ごと 「入りません」★（★直して いません★）', () => {
  const r = H.入らない物を数える(['Sheet1'], [{ name: 'Sheet1' }, { name: 'Sheet2' }], {}, 'xls');
  if (r.件 !== 1) throw new Error('★xls で ' + r.件 + '件★（1件 の はず）');
  if (r.中身[0].何 !== '新しいシート') throw new Error('★中身が 違う★ ' + JSON.stringify(r.中身));
});

T('★★形を 渡さなければ 「入りません」と 言う★★（★黙って 消さない★）', () => {
  const r = H.入らない物を数える(['Sheet1'], [{ name: 'Sheet1' }, { name: 'Sheet2' }], {});
  if (r.件 !== 1) {
    throw new Error('★形 無しで ' + r.件 + '件★（1件 の はず）'
      + '／★渡し忘れた 時に 黙る のが 一番 悪い★');
  }
});

T('★浮かぶ物は 形に よらず 数える（板だけ 直りました）★', () => {
  const r = H.入らない物を数える(['Sheet1'], [{ name: 'Sheet1' }],
    { '図形・グラフ・画像': 2 }, 'xlsx');
  if (r.件 !== 2) throw new Error('★' + r.件 + '件★（2件 の はず）');
});

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
