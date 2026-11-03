/* ctx-menu.test.mjs — ★右クリックのメニューが 画面の中に収まるか★
 *
 *  真値は tests/fixtures/ctx-menu-golden.json。
 *  ★指示役が 実配信を押して測った事故★（2026-08-21）：
 *    メニューの高さ 743px 、画面 619px で ★上端が −470px★。
 *    ★作った7つのうち 5つ（並べ替え・絞り込み・固定・印刷・入力の決まり）に
 *      客の手が届いていなかった★。★置いた≠届く★。
 *
 *  jsdom には 幅も高さも無いので、★数える所を lib/menu-place.js に出して★
 *  どの画面の高さでも 総当たりで 数え切れるようにしてある。
 *  実ブラウザで 619/400/337px を 押して測った値は golden に入っている。
 *
 *  --self-test … わざと壊して「何通りで赤くなるか」を数える（★repo は読むだけ★）
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const SELF = process.argv.includes('--self-test');
const OVERRIDE = process.env.EXALLY_CTX_OVERRIDE ? JSON.parse(process.env.EXALLY_CTX_OVERRIDE) : {};
const srcPath = (rel) => OVERRIDE[rel] || path.join(ROOT, rel);

let JSDOM;
try { ({ JSDOM } = await import('jsdom')); }
catch { console.log('★jsdomが入っていません。この検証は飛ばせません（SKIPを緑と呼ばない）。npm install してください。'); process.exit(1); }

const GOLD = JSON.parse(fs.readFileSync(path.join(ROOT, 'tests/fixtures/ctx-menu-golden.json'), 'utf8'));

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ok   ' + n); } catch (e) { fail++; console.log('  NG   ' + n + '\n       ' + (e && e.message)); } };
const ok = (c, m) => { if (!c) throw new Error(m || 'expected truthy'); };
const eq = (a, b, m) => { if (a !== b) throw new Error((m ? m + ': ' : '') + 'expected ' + JSON.stringify(b) + ' got ' + JSON.stringify(a)); };

const CANVAS_STUB = [
  '(function(){var noop=function(){};var ctx=new Proxy({},{get:function(t,k){',
  ' if(k==="measureText")return function(){return{width:40};};',
  ' if(k==="canvas")return{width:900,height:600};',
  ' if(k==="getImageData")return function(){return{data:[]};};',
  ' if(k==="createLinearGradient"||k==="createPattern")return function(){return{addColorStop:noop};};',
  ' return noop;}});HTMLCanvasElement.prototype.getContext=function(){return ctx;};})();',
].join('\n');

/* ★開いた窓と 刷った回数を数える偽物★（本物のブラウザの窓は出さない） */
const WINDOW_SPY = [
  'window.__開いた窓 = [];',
  'window.open = function(){',
  '  var 書いた = "";',
  '  var w = { __刷った: 0, focus: function(){}, print: function(){ this.__刷った++; },',
  '            document: { open:function(){}, close:function(){}, write:function(s){ 書いた += s; } },',
  '            get 中身(){ return 書いた; } };',
  '  window.__開いた窓.push(w); return w;',
  '};',
].join('\n');

const html = fs.readFileSync(srcPath('book.html'), 'utf8');
const dom = new JSDOM(html.replace(/<script[\s\S]*?<\/script>/g, ''), {
  runScripts: 'dangerously', url: 'http://localhost/', pretendToBeVisual: true,
  beforeParse(w) {
    w.fetch = () => Promise.reject(new Error('no net')); w.scrollTo = () => {}; w.alert = () => {};
    w.matchMedia = w.matchMedia || (() => ({ matches: false, addListener() {}, removeListener() {}, addEventListener() {}, removeEventListener() {} }));
    w.requestAnimationFrame = (cb) => setTimeout(() => cb(1), 0); w.cancelAnimationFrame = () => {};
    w.eval(CANVAS_STUB);
  },
});
const win = dom.window, doc = win.document;
const inject = (c) => { const el = doc.createElement('script'); el.textContent = c; doc.body.appendChild(el); };
const srcs = [...html.matchAll(/<script src="([^"]+)"><\/script>/g)].map((m) => m[1].split('?')[0]).filter((s) => !/^https?:/.test(s));
for (const s of srcs) { const p = srcPath(s); if (fs.existsSync(p)) inject(fs.readFileSync(p, 'utf8')); }
for (const m of html.matchAll(/<script(?![^>]*\ssrc=)[^>]*>([\s\S]*?)<\/script>/g)) inject(m[1]);
try { doc.dispatchEvent(new win.Event('DOMContentLoaded', { bubbles: true })); } catch (e) { /* 続ける */ }
try { win.dispatchEvent(new win.Event('load')); } catch (e) { /* 続ける */ }
inject(WINDOW_SPY);

console.log('\n[ctx-menu] ★右クリックのメニューが 画面の中に収まるか★');
console.log('  真値 = ' + GOLD._measured_with);

const 前 = GOLD['★直す前（指示役が実配信を押して測った）★'];
const 後 = GOLD['★直した後（実ブラウザで押して測った）★'];
const 上に居る物 = GOLD['★上に居なければならない物（増やしたら赤くする）★'];

/* ── 土台 ── */
T('★画面が立ち上がっていて MenuPlace が在る', () => {
  ok(win.MenuPlace && typeof win.MenuPlace.place === 'function', 'MenuPlace が無い');
  ok(typeof win.showCtxMenu === 'function', 'showCtxMenu が無い');
  ok(doc.getElementById('ctx-menu'), 'メニューが無い');
});

/* ── ① ★指示役が踏んだ形★ を そのまま数える ── */
T('★指示役の実測（高さ743px／画面619px）で 上端が画面の外に出ない★', () => {
  const p = win.MenuPlace.place({ x: 200, y: 300, w: 260, h: 前['メニューの高さ'], winW: 1440, winH: 前['画面の高さ'], 余白: 8 });
  ok(p.top >= 0, '★上端が画面の外（' + p.top + 'px）＝踏んだ事故がそのまま★');
  eq(p.top, 8, '端に貼れていない');
  ok(p.中で動かす, '★入り切らないのに 中で動かせない★');
  eq(p.maxHeight, 603, '高さの上限が違う');
});
T('★どこを押しても 上端も左端も 画面の外に出ない（総当たり）★', () => {
  const 高さたち = [619, 400, 337, 200, 100];
  let 外に出た = 0, 数えた = 0;
  for (const winH of 高さたち) {
    for (let y = 0; y <= winH; y += 17) {
      for (const x of [0, 200, 700, 1439]) {
        数えた++;
        const r = win.MenuPlace.届くか({ x, y, w: 260, h: 856, winW: 1440, winH, 余白: 8 });
        if (!r.ok) 外に出た++;
      }
    }
  }
  ok(数えた > 100, '検査が空振りしている（' + 数えた + '通りしか見ていない）');
  eq(外に出た, 0, '★' + 数えた + '通り中 ' + 外に出た + '通りで 手が届かない★');
});
T('★「手が届くか」の判定その物が 空振りしていない（悪い位置を渡して NG になるか）★', () => {
  const o = { x: 200, y: 300, w: 260, h: 743, winW: 1440, winH: 619, 余白: 8 };
  const 悪い = [
    [{ top: -470, left: 200, maxHeight: null, 中で動かす: false, 見える高さ: 743 }, '上端が画面の外'],
    [{ top: 8, left: -50, maxHeight: 603, 中で動かす: true, 見える高さ: 603 }, '左端が画面の外'],
    [{ top: 8, left: 1400, maxHeight: 603, 中で動かす: true, 見える高さ: 603 }, '右端が画面の外'],
    [{ top: 500, left: 200, maxHeight: null, 中で動かす: false, 見える高さ: 743 }, '下端が画面の外'],
    [{ top: 8, left: 200, maxHeight: null, 中で動かす: false, 見える高さ: 603 }, '入り切らないのに 中で動かせない'],
  ];
  for (const [位置, なぜ] of 悪い) {
    const r = win.MenuPlace.届くか(o, 位置);
    ok(!r.ok, '★' + なぜ + ' なのに 「手が届く」と言っている★');
  }
  ok(win.MenuPlace.届くか(o).ok, '今の出し方で 届かないと言っている');
});
T('★押した所が 画面の外でも 端に貼る（余白を守る）★', () => {
  for (const [x, y] of [[-50, -50], [-1, 300], [2000, 300], [200, 2000]]) {
    const p = win.MenuPlace.place({ x, y, w: 260, h: 300, winW: 1440, winH: 619, 余白: 8 });
    ok(p.left >= 8, '★左端が ' + p.left + 'px（余白を守っていない）★');
    ok(p.top >= 8, '★上端が ' + p.top + 'px★');
    ok(p.left + 260 <= 1440 - 8, '★右がはみ出している★');
    ok(p.top + 300 <= 619 - 8, '★下がはみ出している★');
  }
});
T('★入る時は 押した所に そのまま出す（余計に動かさない）★', () => {
  const p = win.MenuPlace.place({ x: 200, y: 100, w: 260, h: 300, winW: 1440, winH: 619, 余白: 8 });
  eq(p.top, 100, '押した所に出ていない');
  eq(p.left, 200, '押した所に出ていない');
  eq(p.maxHeight, null, '入るのに 高さの上限を付けている');
  eq(p.中で動かす, false, '入るのに 中で動かす形にしている');
});
T('★下に入らない時は 上向きに出し直す★', () => {
  const p = win.MenuPlace.place({ x: 200, y: 500, w: 260, h: 300, winW: 1440, winH: 619, 余白: 8 });
  eq(p.top, 200, '上向きに出し直していない');
  ok(p.top + 300 <= 619, '下がはみ出している');
});
T('★右端で押したら 左向きに出す★', () => {
  const p = win.MenuPlace.place({ x: 1400, y: 100, w: 260, h: 300, winW: 1440, winH: 619, 余白: 8 });
  ok(p.left + 260 <= 1440, '★右がはみ出している★');
});

/* ── ② 実ブラウザで測った値と 同じ答えになるか ── */
for (const [名, g] of [['画面619px', 後['画面619px']], ['画面400px', 後['画面400px']], ['画面337px', 後['画面337px']]]) {
  T('★' + 名 + ' … 実ブラウザで測った 上=' + g['メニューの上'] + ' 見える高さ=' + g['見える高さ'] + ' と同じ★', () => {
    const winH = Number(名.replace(/[^0-9]/g, ''));
    const p = win.MenuPlace.place({ x: 200, y: 100, w: 260, h: g['中身の高さ'], winW: 1440, winH, 余白: 8 });
    eq(p.top, g['メニューの上'], '上');
    eq(p.見える高さ, g['見える高さ'], '見える高さ');
    ok(p.中で動かす, '中で動かせない');
  });
}

/* ── ③ 画面の書き方（中で動かせるか）── */
T('★メニューは 中で動かせる書き方になっている（overflow-y:auto）★', () => {
  const i = html.indexOf('#ctx-menu {');
  ok(i > 0, '#ctx-menu の書き方が無い');
  const blk = html.slice(i, html.indexOf('}', i));
  ok(blk.indexOf('overflow-y:auto') >= 0, '★入り切らない時に 中で動かせない★');
});
T('★出す時に 高さの上限を 付け直している（前の値が残らない）★', () => {
  ok(html.indexOf("m.style.maxHeight=''; m.style.left='0px'; m.style.top='0px';") >= 0,
    '★測る前に 上限を外していない＝中身の高さが 正しく取れない★');
  ok(html.indexOf('m.style.maxHeight = 置く.maxHeight') >= 0, '上限を付けていない');
});
T('★出すたびに 一番上から見せる（前に動かした所が残らない）★', () => {
  ok(html.indexOf('m.scrollTop = 0;') >= 0, '★前に動かした位置が残る＝上の項目が見えない★');
});

/* ── ④ ★大事な物が 上に居る（増やしたら赤）★ ──
 *
 *  ★2026-08-31 に 数え方を 変えた（理由を 残す）★
 *    右クリックを ★実Excel の 中身★に 合わせて 25個 → ★37個★に した。
 *    平らに 並べたら 高さ1,193px で 画面(884px)に 入らず スクロールに なった
 *    ので、司さんの 決まり「★ごちゃごちゃに ならないよう ドロップダウンでも
 *    いいから 綺麗に★」に 従って ★組（子メニュー）★に した。
 *    ⇒ 押す物の 印が `onclick="..."` から ★data-ctx="働きの名前"★ に 変わった。
 *    ★2026-08-21 の 事故の 守りは そのまま★＝
 *      大事な5つ（並べ替え・絞り込み・固定・印刷・入力の決まり）は
 *      ★上の段の 8番目までに 居る★（組の 中でも よい＝1回 当てれば 出る）。
 */
/* ★中身は「開いた時」に 作る★＝数える前に 1回 開く（2026-08-31）
   ＝開く前に 数えると ★0個★に なり「落とした」と 誤報する。 */
T('★右クリックを 開くと 中身が 作られる★', () => {
  ok(typeof win.showCtxMenu === 'function', 'showCtxMenu が 無い');
  win.sel(0, 0, 0, 0);
  win.showCtxMenu(10, 10);
  const n = doc.querySelectorAll('#ctx-menu .ctx-item').length;
  ok(n > 0, '★開いても 中身が 0個★');
});

/* ★前から 在った 物の 正本★（lib/ctx-menu.js が 持つ） */
const 前からの物 = (await import(
  pathToFileURL(path.join(ROOT, 'lib/ctx-menu.js')).href)).default.前からの物;

const 上の段 = () => {
  const m = doc.getElementById('ctx-menu');
  return [...m.children].filter((e) => e.classList && e.classList.contains('ctx-item'));
};
const どこに居る = (act) => {
  const 段 = 上の段();
  for (let i = 0; i < 段.length; i++) {
    if (段[i].getAttribute('data-ctx') === act) return i;
    if (段[i].querySelector('[data-ctx="' + act + '"]')) return i;   /* 組の中 */
  }
  return -1;
};
/* ★★2026-08-31：見張りを ★Excel と 同じか★ で 見る 形に 直した★★
 *
 *  司さん「おれは 最初から ★Excelと 同じように 見せろ★って 言わんかったか？」
 *
 *  ★私の 間違い（3回）★
 *    ① 37個 全部を 組（▸）に まとめ、前からの 17個を 2段に 落とした
 *    ② 差し戻す時に ★うちの 並び★（2026-08-21 の 事故直し）で 平らに 並べた
 *    ③ 見張りまで ★うちの 決め事★（大事な5つを 上から9番目まで）を 見ていた
 *
 *  ⇒ ★★見るのは「実Excel の 並びと 同じか」★★
 *    2026-08-21 の 事故（大事な物が 届かない）は
 *    ★MenuPlace の 総当たり★（この 上の ①②）が すでに 見ている。
 *    ★並びまで うちの 都合で 決めない★。
 */

/* ★実Excel の セルメニューの 並び★（docs/excel-commandbars-2026-08-30.tsv の Cell・実測） */
const Excelの並び = [
  'ctxCut',        /* 切り取り(T) */
  'ctxCopy',       /* コピー(C) */
  'ctxPaste',      /* 貼り付け(P) */
  'ctxPasteValue', /* 形式を選択して貼り付け(S)... */
  /* ★2026-09-03★＝★どちらへ 詰めるかを 聞く 窓★に した（司さん「Excelと 同じ」）
     ⇒ 呼ぶ先は 2つとも 同じ 手（引数で 分ける）＝★挿入／削除の 順は そのまま★ */
  '挿入の窓を開く',   /* 挿入(I)... */
  '削除の窓を開く',   /* 削除(D)... */
  'ctxDelete',     /* 数式と値のクリア(N) */
  'フィルター',     /* フィルター(E)   ★Excel も ▸★ */
  '並べ替え',       /* 並べ替え(O)     ★Excel も ▸★ */
  '新しいコメント',      /* コメントの挿入(M) */
  'コメントを消す',      /* コメントの削除(M) */
  'コメントの表示',      /* コメントの表示/非表示(O) */
  'ctxFormat',     /* セルの書式設定(F)... */
  'リンク',              /* ハイパーリンク(H)... */
  'ハイパーリンクを削除',
];

T('★★実Excel の 並び そのままに 出ている★★', () => {
  const 段 = 上の段();
  /* 上の段の 名前か 働きで 引く */
  /* ★組の 見出しは data-ctx を 持たない★（押しても 何も 起きない 行だから）
     ⇒ 字で 引く。★絵文字は 落とす★（字だけで 見る） */
  const 印 = 段.map((e) => {
    const a = e.getAttribute('data-ctx');
    if (a) return a;
    return e.childNodes[0].textContent.trim().replace(/^\S+\s*/, '');
  });
  let 前 = -1;
  for (const 物 of Excelの並び) {
    const i = 印.indexOf(物);
    ok(i >= 0, '★' + 物 + ' がメニューに無い★：' + 印.join(' / '));
    ok(i > 前, '★' + 物 + ' が Excel の 順より 前に 出ている（' + i + ' ≦ ' + 前 + '）★');
    前 = i;
  }
});

T('★Excel が ▸ に している 2つ（フィルター／並べ替え）だけ 組★', () => {
  const 組 = [...doc.querySelectorAll('#ctx-menu .ctx-sub')]
    .map((e) => e.childNodes[0].textContent.trim());
  eq(組.length, 2, '★組の 数★：' + 組.join(' / '));
  ok(組.indexOf('🔽 フィルター') >= 0 || 組.some((x) => x.indexOf('フィルター') >= 0), 'フィルターが 組でない');
  ok(組.some((x) => x.indexOf('並べ替え') >= 0), '並べ替えが 組でない');
});

T('★うちにしか 無い 物は 一番 下（Excel の 並びを 崩さない）★', () => {
  const 段 = 上の段();
  const 印 = 段.map((e) => e.getAttribute('data-ctx') || '');
  const うちだけ = ['freezePanes', 'printSheet', 'openCondFormat', 'openValid', 'explainCell'];
  const Excel最後 = 印.indexOf('ハイパーリンクを削除') >= 0
    ? 印.indexOf('ハイパーリンクを削除') : 印.indexOf('ctxFormat');
  for (const a of うちだけ) {
    const i = 印.indexOf(a);
    ok(i >= 0, '★' + a + ' が無い★');
    ok(i > Excel最後, '★' + a + ' が Excel の 物より 上に 居る★');
  }
});

T('★名前は Excel の 字（長い 説明を 付けない）★', () => {
  const 長い = 上の段()
    .map((e) => e.childNodes[0].textContent.trim())
    .filter((t) => t.replace(/^\S+\s*/, '').length > 12);
  eq(長い.length, 0, '★長すぎる 名前★：' + 長い.join(' / '));
});

/* ── ⑤ 前からある物を 押し直す ── */
T('★押す物を 1つも 落としていない（★37個★・2026-08-21は 25個）★', () => {
  /* ★組の 見出し（▸）は 数えない★＝押しても 何も 起きない 行だから。
     ★これを 数えると 中身を 3個 消しても 緑のままだった★（2026-08-31 実測）。
     押せる 物には data-ctx（働きの名前）が 付いている。 */
  const n = [...doc.querySelectorAll('#ctx-menu .ctx-item[data-ctx]')]
    .filter((e) => e.getAttribute('data-ctx')).length;
  ok(n >= 37, '★' + n + '個しか 無い＝実Excel の 中身を 落とした★');
});
T('★元の 画面が 探している id が 全部 在る★', () => {
  for (const id of ['ctx-valid-list', 'ctx-bad-cells', 'ctx-clear-filter',
    'ctx-unfreeze', 'ctx-unhide-row', 'ctx-unhide-col']) {
    ok(doc.getElementById(id), '★' + id + ' が 無い＝右クリックを 開いた 途端に 落ちる★');
  }
});
T('★前からある物：右クリックを出して 閉じられる★', () => {
  win.sel(0, 0, 0, 0);
  win.showCtxMenu(10, 10);
  ok(doc.getElementById('ctx-menu').classList.contains('show'), '出ていない');
  win.hideCtxMenu();
  ok(!doc.getElementById('ctx-menu').classList.contains('show'), '閉じられない');
});

console.log('\n  ' + pass + ' 緑 / ' + fail + ' 赤');

if (SELF) {
  const { spawnSync } = await import('node:child_process');
  const os = await import('node:os');
  const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'exally-ctx-'));
  console.log('\n[self-test] わざと壊して 赤くなるかを数える（★repo は読むだけ★）');
  const BREAKS = [
    ['lib/menu-place.js', '★前の書き方に戻す（y−高さ だけ＝上端が画面の外へ）★',
      (s) => s.replace('    if (top < 余白) top = 余白;', '')],
    ['lib/menu-place.js', '★入り切らなくても 中で動かせるようにしない★',
      (s) => s.replace('    var maxHeight = 中で動かす ? 使える高さ : null;', '    var maxHeight = null;')],
    ['lib/menu-place.js', '★下がはみ出しても 出し直さない★',
      (s) => s.replace('    if (top + 見える高さ > winH - 余白) top = o.y - 見える高さ;', '')],
    ['lib/menu-place.js', '★右がはみ出しても 出し直さない（最後の守りごと外す）★',
      (s) => s.replace('    if (left + 見える幅 > winW - 余白) left = o.x - 見える幅;', '')
              .replace('    if (left + 見える幅 > winW - 余白) left = Math.max(余白, winW - 余白 - 見える幅);', '')],
    ['lib/menu-place.js', '★左端が 画面の外に出るのを 直さない★',
      (s) => s.replace('    if (left < 余白) left = 余白;', '')],
    ['lib/menu-place.js', '★下がはみ出す時の 最後の守りを外す★',
      (s) => s.replace('    if (top + 見える高さ > winH - 余白) top = Math.max(余白, winH - 余白 - 見える高さ);', '')],
    ['lib/menu-place.js', '★入るのに 高さの上限を付けてしまう（無駄にスクロールが出る）★',
      (s) => s.replace('    var 中で動かす = o.h > 使える高さ;', '    var 中で動かす = true;')],
    ['lib/menu-place.js', '★手が届くかの判定を 素通りさせる（上端）★',
      (s) => s.replace("    if (p.top < 0) return { ok: false, why: '上端が画面の外（' + p.top + 'px）' };", '')],
    ['lib/menu-place.js', '★手が届くかの判定を 素通りさせる（左端）★',
      (s) => s.replace("    if (p.left < 0) return { ok: false, why: '左端が画面の外（' + p.left + 'px）' };", '')],
    ['lib/menu-place.js', '★手が届くかの判定を 素通りさせる（右端）★',
      (s) => s.replace("    if (p.left + Math.min(o.w, o.winW) > o.winW) return { ok: false, why: '右端が画面の外' };", '')],
    /* ★#ctx-menu の 中の 1行だけ 消す★
       （ただ 'overflow-y:auto;' と 書くと ★別の 場所の 1つ目★が 消えて
         この 検査は 素通りする。2026-08-31 実際に 踏んだ） */
    ['book.html', '★メニューを 中で動かせなくする（overflow-y を消す）★',
      (s) => {
        const i = s.indexOf('#ctx-menu {');
        const j = s.indexOf('}', i);
        return s.slice(0, i) + s.slice(i, j).replace('overflow-y:auto;', '') + s.slice(j);
      }],
    ['book.html', '★測る前に 高さの上限を外さない（中身の高さが取れない）★',
      (s) => s.replace("  m.style.maxHeight=''; m.style.left='0px'; m.style.top='0px';", "  m.style.left='0px'; m.style.top='0px';")],
    /* ★2026-09-03 に 直した★＝`m.scrollTop = 0;` が ★2か所★に なった
       （一覧が 長い時に 1行へ 切り替えた 後の 置き直し）。
       ★1か所だけ 消しても もう1か所が 残って 素通りに なった★ ⇒ ★全部 消す★ */
    ['book.html', '★出すたびに 一番上へ戻さない（前の位置が残る）★',
      (s) => s.split('m.scrollTop = 0;').join('/*消した*/')],
    ['book.html', '★上限を付けない（画面より高いまま出す）★',
      (s) => s.replace("  m.style.maxHeight = 置く.maxHeight ? (置く.maxHeight+'px') : '';", '')],
    /* ★2026-08-31 に 印を 直した★＝中身の 正本が 画面から lib/ctx-menu.js へ 移った。
       ★壊す 先も 一緒に 移す★（印が 古いままだと ★素通り★＝見張りが 眠る）。 */
    /* ★★2026-08-31：★Excel の 並びを 崩したら 赤★ に 直した★★
       前は ★うちの 決め事★（大事な5つを 上に）を 壊していた。
       今 見るのは ★実Excel と 同じ 並びか★。 */
    ['lib/ctx-menu.js', '★Excel の 順を 入れ替える（コピーを 一番 下へ）★',
      (s) => s.replace(
        "    { 印: '📋', 名: 'コピー',   画面: 'ctxCopy',  鍵: 'Ctrl+C', Excel: 'コピー(C)' },",
        "")
        .replace("    { 印: '🤖', 名: 'AIに解説させる'",
          "    { 印: '📋', 名: 'コピー', 画面: 'ctxCopy', 鍵: 'Ctrl+C', Excel: 'コピー(C)' },"
          + String.fromCharCode(10) + "    { 印: '🤖', 名: 'AIに解説させる'")],
    ['lib/ctx-menu.js', '★Excel が 平らに している 物を 組（▸）に 落とす★',
      (s) => s.replace(
        "    { 印: '🎨', 名: 'セルの書式設定', 画面: 'ctxFormat', Excel: 'セルの書式設定(F)...' },",
        "    { 印: '🎨', 名: '書式', 子: ["
        + "{ 印: '🎨', 名: 'セルの書式設定', 画面: 'ctxFormat' }] },")],
    ['lib/ctx-menu.js', '★Excel の 物を メニューから 外す（切り取り）★',
      (s) => s.replace("画面: 'ctxCut'", "画面: 'ctxCopy'")],
    ['lib/ctx-menu.js', '★名前に 長い 説明を 足す（Excel と 別物に なる）★',
      (s) => s.replace("名: 'セルの書式設定'",
        "名: 'セルの書式設定（字の 色や 大きさ・罫線・塗りを まとめて 決める）'")],
    ['lib/ctx-menu.js', '★うちだけの 物を Excel の 物より 上に 出す★',
      (s) => s.replace("    { 印: '✂️', 名: '切り取り'",
        "    { 印: '🤖', 名: 'AIに解説させる', 画面: 'explainCell' },"
        + String.fromCharCode(10) + "    { 印: '✂️', 名: '切り取り'")],

    ['lib/ctx-menu.js', '★元の 画面が 探している id を 消す（開いた途端に 落ちる）★',
      (s) => s.replace("id: 'ctx-unhide-row',", '')],
    ['lib/ctx-menu.js', '★実Excel の 中身を 落とす（コメント 3つを 消す）★',
      (s) => s.split("リボン: '新しいコメント'").join("リボン: '在るわけない'")
        .split("リボン: 'コメントを消す'").join("リボン: '在るわけない'")
        .split("リボン: 'コメントの表示'").join("リボン: '在るわけない'")],
    ['book.html', '★中身を 画面に 直に 書き戻す（部品を 使わない）★',
      (s) => s.replace('window.CtxMenu.出す(', 'window.__無い.出す(')],
  ];
  let red = 0;
  for (const [rel, name, brk] of BREAKS) {
    const orig = fs.readFileSync(path.join(ROOT, rel), 'utf8');
    const bad = brk(orig);
    if (bad === orig) { console.log('  ★置換できず★  ' + name); continue; }
    const tmpFile = path.join(TMP, rel.replace(/[\\/]/g, '_'));
    fs.writeFileSync(tmpFile, bad, 'utf8');
    const env = Object.assign({}, process.env, { EXALLY_CTX_OVERRIDE: JSON.stringify({ [rel]: tmpFile }) });
    const isRed = spawnSync(process.execPath, [path.join(__dirname, 'ctx-menu.test.mjs')], { encoding: 'utf8', env }).status !== 0;
    if (isRed) { red++; console.log('  赤くなった  ' + name); }
    else console.log('  ★素通り★  ' + name);
  }
  /* ★repo を書き換えていない事を 押した後に 数える★ */
  for (const rel of ['book.html', 'lib/menu-place.js']) {
    const now = fs.readFileSync(path.join(ROOT, rel), 'utf8');
    if (now.includes('ふえた0') || now.includes('var 中で動かす = true;')) {
      console.log('  ★NG★ ' + rel + ' に わざと壊した物が残っている');
      process.exit(1);
    }
  }
  try { fs.rmSync(TMP, { recursive: true, force: true }); } catch (e) { /* 消せなくても検査は済んでいる */ }
  console.log('\n  ' + red + '/' + BREAKS.length + ' 通りで赤くなった');
  process.exit(red === BREAKS.length ? 0 : 1);
}

process.exit(fail ? 1 : 0);
