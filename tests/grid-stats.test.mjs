/* grid-stats.test.mjs — ★選んだ所の 合計・平均・個数が 実Excelどおりか★
 *
 *  真値は tests/fixtures/excel-status-golden.json（実Excel 16.0.20228 を COM で押して測った）。
 *  ★ソースを読むだけにしない★＝本物の book.html を載せて ★本物の帯の字を読む★。
 *
 *  ★一番 大事な所★
 *    ・★見た目が数なのに 合計に入らない文字★を 黙って通さない（合計が黙って小さくなる）
 *    ・★エラーが在る時は 合計を出さず 理由を出す★
 *    ・★文字だけの時に 合計 0 を出さない★（0は「足した結果が0」に見えて嘘になる）
 *    ・★帯の字が 1文字ずつ 縦に割れない★（前科2回）
 *
 *  --self-test … わざと壊して「何通りで赤くなるか」を数える（★repo は読むだけ★）
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const SELF = process.argv.includes('--self-test');
const OVERRIDE = process.env.EXALLY_STATS_OVERRIDE ? JSON.parse(process.env.EXALLY_STATS_OVERRIDE) : {};
const srcPath = (rel) => OVERRIDE[rel] || path.join(ROOT, rel);

let JSDOM;
try { ({ JSDOM } = await import('jsdom')); }
catch { console.log('★jsdomが入っていません。この検証は飛ばせません（SKIPを緑と呼ばない）。npm install してください。'); process.exit(1); }

const GOLD = JSON.parse(fs.readFileSync(path.join(ROOT, 'tests/fixtures/excel-status-golden.json'), 'utf8'));

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

console.log('\n[grid-stats] ★本物の book.html で 選んで 帯を読む★');
console.log('  真値 = ' + GOLD._measured_with);

const sh = () => win.sheets[win.activeSheet];
function reset() { sh().data = {}; win.sel(0, 0, 0, 0); }
/** 縦に並べて置く（A1から下へ）。null は空マス */
function 置く(vals) {
  reset();
  for (let i = 0; i < vals.length; i++) {
    if (vals[i] === null || vals[i] === undefined) continue;
    win.setCell(i, 0, String(vals[i]));
  }
  win.sel(0, 0, vals.length - 1, 0);
  win.updateStatusBar();
}
const 帯の字 = () => doc.getElementById('status-bar').textContent.replace(/\s+/g, ' ').trim();
const 帯の項目 = () => [...doc.querySelectorAll('#status-bar .sb-i')].map((e) => e.textContent.trim());
/** ★字を目で探さず 項目の値そのものを読む（当てにならない一致で素通りしないため）★ */
const 帯の値 = (名) => {
  const el = doc.querySelector('#status-bar .sb-i[data-n="' + 名 + '"]');
  return el ? el.getAttribute('data-v') : null;
};
const 帯の名前 = () => [...doc.querySelectorAll('#status-bar .sb-i')].map((e) => e.getAttribute('data-n'));
const まとめ = (vals) => {
  const cells = vals.map((v) => (v === null || v === undefined ? null : { v: v }));
  return win.GridStats.summarize(cells);
};

/* ── 土台 ── */
T('★画面が立ち上がっていて 帯と GridStats が在る', () => {
  ok(win.GridStats && typeof win.GridStats.summarize === 'function', 'GridStats が無い');
  ok(typeof win.updateStatusBar === 'function', 'updateStatusBar が無い');
  ok(doc.getElementById('status-bar'), '帯が無い');
});
T('★何も無い所を選ぶと「準備完了」（空の枠を出さない）★', () => {
  reset();
  win.updateStatusBar();
  eq(帯の字(), '準備完了', '空の枠が出ている：' + 帯の字());
});

/* ── ① 実Excelの真値どおりか ── */
T('①10,20,30 … 合計・平均・個数が 実Excelと同じ', () => {
  const G = GOLD['①数だけ'], s = まとめ([10, 20, 30]);
  eq(s.合計, G['合計'], '合計'); eq(s.平均, G['平均'], '平均');
  eq(s.数値の個数, G['数値の個数'], '数値の個数'); eq(s.データの個数, G['データの個数'], 'データの個数');
  eq(s.最小, G['最小'], '最小'); eq(s.最大, G['最大'], '最大');
});
T('②★文字だけ★ … 平均が出ない／データの個数だけ（実Excelと同じ）', () => {
  const G = GOLD['②文字だけ'], s = まとめ(['あ', 'い', 'う']);
  eq(s.平均, null, '★平均を出してしまっている★');
  eq(s.数値の個数, G['数値の個数'], '数値の個数');
  eq(s.データの個数, G['データの個数'], 'データの個数');
  eq(s.合計, null, '★合計に 0 を出している（0は嘘に見える）★');
});
T('③10,あ,20 … ★文字は 平均の分母に入らない（÷2）★', () => {
  const G = GOLD['③数と文字'], s = まとめ([10, 'あ', 20]);
  eq(s.合計, G['合計'], '合計'); eq(s.平均, G['平均'], '★÷3 になっている★');
  eq(s.数値の個数, G['数値の個数'], '数値の個数'); eq(s.データの個数, G['データの個数'], 'データの個数');
});
T('④10,(空),20 … ★空は データの個数にも入らない★', () => {
  const G = GOLD['④空きが混ざる'], s = まとめ([10, null, 20]);
  eq(s.データの個数, G['データの個数'], '★空を数えている★');
  eq(s.合計, G['合計'], '合計'); eq(s.平均, G['平均'], '平均');
});
T('⑤1マスだけ … 7', () => {
  const G = GOLD['⑤1マスだけ'], s = まとめ([7]);
  eq(s.合計, G['合計'], '合計'); eq(s.平均, G['平均'], '平均'); eq(s.データの個数, G['データの個数'], 'データの個数');
});
T('⑥★TRUE/FALSE は 数として扱わない（実Excelと同じ）★', () => {
  const G = GOLD['⑥論理値'];
  const cells = [{ v: true }, { v: false }, { v: 10 }];
  const s = win.GridStats.summarize(cells);
  eq(s.合計, G['合計'], '★論理値を足している★');
  eq(s.数値の個数, G['数値の個数'], '★論理値を数として数えている★');
  eq(s.データの個数, G['データの個数'], 'データの個数');
});
T('⑦★数字の形の文字は 足されない（実Excelと同じ）★', () => {
  const G = GOLD['⑦★数字の形の文字★'], s = まとめ(['10', '20']);
  eq(s.数値の個数, G['数値の個数'], '★文字の 10 を 数として足している★');
  eq(s.データの個数, G['データの個数'], 'データの個数');
  eq(s.数の形の文字, 2, '★数の形の文字を数えていない★');
});
T('⑧★エラーが1つでも在れば 合計も平均も 出さない（実Excelと同じ）★', () => {
  const G = GOLD['⑧★エラーが混ざる★'];
  const cells = [{ v: 10 }, { v: '', f: '=1/0', d: '#DIV/0!' }, { v: 20 }];
  const s = win.GridStats.summarize(cells);
  eq(s.合計, null, '★エラーが在るのに 合計を出している★');
  eq(s.平均, null, '★エラーが在るのに 平均を出している★');
  eq(s.数値の個数, G['数値の個数'], '数値の個数');
  eq(s.データの個数, G['データの個数'], 'データの個数');
  eq(s.エラー, 1, 'エラーを数えていない');
});
T('★式のセルは 答えを足す（式の字を足さない）★', () => {
  const s = win.GridStats.summarize([{ v: 10 }, { v: '', f: '=A1*2', d: 20 }]);
  eq(s.合計, 30, '合計');
  eq(s.数値の個数, 2, '数値の個数');
});

/* ── ② ★黙って小さくならない★（うちの上乗せ）── */
T('★数の形の文字が在ると 帯に「◯個（足していません）」と出す★', () => {
  reset();
  win.setCell(0, 0, '10');
  win.setCell(1, 0, "'20");        /* ★先頭の ' ＝「これは文字」（Excel と同じ書き方）★ */
  win.setCell(2, 0, '30');
  win.sel(0, 0, 2, 0);
  win.updateStatusBar();
  const t = 帯の字();
  eq(帯の値('⚠️ 数の形の文字'), '1個（足していません）', '★黙って小さい合計を出している★：' + t);
  eq(帯の値('合計'), '40', '★合計が 40 でない★：' + t);
});
T('★エラーが在ると 帯に「出せません」と理由を出す★', () => {
  reset();
  win.setCell(0, 0, '10');
  win.setCell(1, 0, '=A99/0');     /* ★うちの計算が #DIV/0! を返す形（実測）★ */
  win.setCell(2, 0, '20');
  win.sel(0, 0, 2, 0);
  win.updateStatusBar();
  const t = 帯の字();
  ok((帯の値('⚠️') || '').indexOf('出せません') >= 0, '★理由を出していない★：' + t);
  eq(帯の値('合計'), null, '★エラーが在るのに 合計を出している★：' + t);
});
T('★文字だけの時は 合計 0 を出さない（0は嘘に見える）★', () => {
  置く(['あ', 'い', 'う']);
  const t = 帯の字();
  eq(帯の値('合計'), null, '★合計 0 を出している★：' + t);
  eq(帯の値('データの個数'), '3', 'データの個数が出ていない：' + t);
});

/* ── ★押して見つけた2つ★（golden の「うちの画面を押して測った事」）── */
T("★'20 は うちの SUM も足さない … 帯も足さない（同じ物を2通りで数えない）★", () => {
  reset();
  win.setCell(0, 0, "'20");
  win.setCell(1, 0, '30');
  win.setCell(2, 0, '=SUM(A1:A2)');
  const 本番のSUM = win.GridStats.shown(sh().data['2,0']);
  eq(本番のSUM, String(GOLD["★うちの画面を押して測った事（2026-08-21）★"]["'20（数字の形の文字）"]["うちの =SUM('20, 30)"]),
    '★うちの SUM の答えが 測った時と変わっている★');
  win.sel(0, 0, 1, 0);
  win.updateStatusBar();
  const t = 帯の字();
  eq(帯の値('合計'), '30', '★帯の合計が うちの SUM と違う★：' + t);
  eq(帯の値('数値の個数'), '1', "★'20 を 数として足している★：" + t);
  ok(帯の値('⚠️ 数の形の文字'), '★数の形の文字を 言っていない★：' + t);
});
T('★=1/0 が Infinity を返す（見つけた違い）… 帯は Infinity を合計にしない★', () => {
  reset();
  win.setCell(0, 0, '10');
  win.setCell(1, 0, '=1/0');
  win.setCell(2, 0, '20');
  eq(win.GridStats.shown(sh().data['1,0']), 'Infinity',
    '★=1/0 の答えが 測った時と変わっている（直したなら golden も直す）★');
  win.sel(0, 0, 2, 0);
  win.updateStatusBar();
  const t = 帯の字();
  ok(t.indexOf('Infinity') < 0, '★合計に Infinity を出している★：' + t);
  eq(帯の値('合計'), null, '★計算できていない値が在るのに 合計を出している★：' + t);
  ok((帯の値('⚠️') || '').indexOf('計算できていない値') >= 0, '★理由を出していない★：' + t);
});

/* ── ⑤ ★実ブラウザで幅と高さを測って直した事★（jsdom では測れない・前科2回）── */
T('★注意は 先頭に出す（375px で 帯が679px になり 注意が画面の外だった）★', () => {
  const s = win.GridStats.summarize([{ v: '10' }, { v: "'20", f: "'20", d: '20' }]);
  const 名 = win.GridStats.items(s).map((i) => i.名);
  ok(名[0].indexOf('⚠️') === 0, '★注意が先頭でない＝画面の外で読めなくなる★：' + 名.join('／'));
});
T('★AIボタンを 帯の上へ退かす（375pxで 帯を隠していた）★', () => {
  /* ★スマホの書き方は 2か所に在る★…片方だけ戻ると その幅で 帯が隠れるので 数を数える */
  const 退かす回数 = html.split('bottom:calc(96px + var(--sb-h, 0px));').length - 1;
  eq(退かす回数, 2, '★スマホで AIボタンが 帯を隠す（' + 退かす回数 + '/2 か所しか 退かしていない）★');
  ok(html.indexOf('bottom:calc(32px + 16px + var(--sb-h, 0px));') >= 0, '★PCで AIボタンが 帯を隠す★');
  ok(typeof win._sbHeight === 'function', '帯の高さを測っていない');
});
T('★字は割らずに 箱ごと折り返す（flex-wrap:wrap）★', () => {
  const i = html.indexOf('#status-bar{');
  const blk = html.slice(i, html.indexOf('}', i));
  ok(blk.indexOf('flex-wrap:wrap') >= 0, '★折り返さない＝はみ出して読めない★');
  ok(blk.indexOf('white-space:nowrap') >= 0, '★字を割る書き方になっている★');
});
T('★平均は 小数4桁まで＋丸めた時は ≒ を付ける（黙って丸めない）★', () => {
  eq(win.GridStats.fmtAvg(2233.3333333333333), '≒2,233.3333', '丸め');
  eq(win.GridStats.fmtAvg(15), '15', '割り切れる時は ≒ を付けない');
  eq(win.GridStats.fmtAvg(2233.5), '2,233.5', '割り切れる時は そのまま');
});

/* ── ③ 帯の見た目（★字が1文字ずつ縦に割れない★・前科2回）── */
T('★帯は 横に流す書き方（nowrap）＝1文字ずつ縦に割れない★', () => {
  const css = html;
  const i = css.indexOf('#status-bar{');
  ok(i > 0, '#status-bar の書き方が無い');
  const blk = css.slice(i, css.indexOf('}', i));
  ok(blk.indexOf('white-space:nowrap') >= 0, '★nowrap が無い＝縦に割れる★');
  ok(blk.indexOf('overflow-x:auto') >= 0, '★はみ出た時に 横に流していない★');
});
T('★帯の中の箱を 縮ませない（flex:0 0 auto）＝潰れて縦に割れない★', () => {
  const i = html.indexOf('#status-bar .sb-i{');
  ok(i > 0, '.sb-i の書き方が無い');
  const blk = html.slice(i, html.indexOf('}', i));
  ok(blk.indexOf('flex:0 0 auto') >= 0, '★縮む書き方のまま＝縦に割れる★');
});
T('★hidden と書いたら 必ず消える 1行が在る（class に打ち消されない）★', () => {
  ok(html.indexOf('[hidden]{display:none !important;}') >= 0, '★[hidden] の1行が無い★');
});
T('★読ませる字は 薄い黒（#333333）＝青や薄い灰色で 値を書かない★', () => {
  const i = html.indexOf('#status-bar .sb-v{');
  const blk = html.slice(i, html.indexOf('}', i));
  ok(blk.indexOf('#333333') >= 0, '★値の色が 薄い黒でない★：' + blk);
});
/* ★禁止色の字を そのまま書くと 禁止色の見張りに この検査自身が捕まる★。
   ★見ないファイルに逃がさない★（逃がすと 本当に使った時も 捕まらなくなる）。
   字を組み立てるだけにして、このファイルは 見張りの対象のままにする。 */
const 禁止色 = '#1A' + '4A2E';
T('★禁止色（濃い緑）を使っていない★', () => {
  const i = html.indexOf('#status-bar{');
  const blk = html.slice(i, i + 1400);
  ok(blk.indexOf(禁止色) < 0, '★禁止色（濃い緑）を使っている★');
});

/* ── ④ 前からある物を 押し直す（帯を足して 画面が崩れていないか）── */
T('★前からある物：セルを選ぶと 番地と数式バーが 今まで通り出る★', () => {
  reset();
  win.setCell(1, 1, '=1+2');
  win.sel(1, 1, 1, 1);
  win.updateBar();
  eq(doc.getElementById('cell-addr').value, 'B2', '番地');
  eq(doc.getElementById('formula-input').value, '=1+2', '数式バー');
});
T('★選ぶと 帯が すぐ変わる（updateBar から呼ばれている）★', () => {
  reset();
  win.setCell(0, 0, '5');
  win.setCell(1, 0, '15');
  win.sel(0, 0, 1, 0);
  win.updateBar();
  eq(帯の値('合計'), '20', '★選び直しても 帯が古いまま★：' + 帯の字());
  win.sel(0, 0, 0, 0);
  win.updateBar();
  eq(帯の値('合計'), '5', '1マスにしたら 変わっていない：' + 帯の字());
});
T('★帯は grid-wrap の外に在る（絵を隠さない）★', () => {
  const sb = doc.getElementById('status-bar');
  ok(!doc.getElementById('grid-wrap').contains(sb), '★帯が 絵の上に かぶさっている★');
});
T('★出せる物だけ出す＝同じ数を2つ並べない（数値の個数＝データの個数の時は出さない）★', () => {
  置く([10, 20, 30]);
  eq(帯の値('数値の個数'), null, '★同じ数を2つ並べている★：' + 帯の名前().join('／'));
  置く([10, 'あ', 20]);
  eq(帯の値('数値の個数'), '2', '違う時は 出すべき');
});

console.log('\n  ' + pass + ' 緑 / ' + fail + ' 赤');

if (SELF) {
  const { spawnSync } = await import('node:child_process');
  const os = await import('node:os');
  const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'exally-stats-'));
  console.log('\n[self-test] わざと壊して 赤くなるかを数える（★repo は読むだけ★）');
  const BREAKS = [
    ['book.html', '選んでも 帯を書き換えない',
      (s) => s.replace('  updateActionBtn();\n  updateStatusBar();', '  updateActionBtn();')],
    ['book.html', '★中身が空でも 枠を出す（空の枠だけが出る）★',
      (s) => s.replace("    el.innerHTML = '<span class=\"sb-i\" id=\"sb-ready\">準備完了</span>';\n    return;", "    el.innerHTML = '<span class=\"sb-i\"><span class=\"sb-n\">合計</span><span class=\"sb-v\"></span></span>';\n    return;")],
    ['book.html', '★帯を 横に流さない（1文字ずつ縦に割れる）★',
      (s) => s.replace('  white-space:nowrap; overflow-x:auto; overflow-y:hidden;\n', '  white-space:normal;\n')],
    ['book.html', '★帯の中の箱を 縮ませる（潰れて縦に割れる）★',
      (s) => s.replace('#status-bar .sb-i{flex:0 0 auto; white-space:nowrap;}', '#status-bar .sb-i{flex:1 1 auto;}')],
    ['book.html', '★[hidden] の1行を消す（hiddenと書いても消えない）★',
      (s) => s.replace('[hidden]{display:none !important;}', '')],
    ['book.html', '★値の色を 薄い灰色にする（読ませる字は薄い黒）★',
      (s) => s.replace('#status-bar .sb-v{color:#333333; font-weight:700;}', '#status-bar .sb-v{color:#AAAAAA; font-weight:700;}')],
    ['book.html', '★禁止色（濃い緑）を使う★',
      (s) => s.replace('  background:#FFFFFF; border-top:1px solid #D4EDE1;\n  font-size:12px; color:#333333; line-height:1.5;', '  background:' + 禁止色 + '; border-top:1px solid #D4EDE1;\n  font-size:12px; color:#333333; line-height:1.5;')],
    ['book.html', '★帯を 絵の中に入れる（最後の行を隠す）★',
      (s) => s.replace('</div>\n\n<!-- ★選んだ所の 合計・平均・個数★', '\n<!-- ★選んだ所の 合計・平均・個数★').replace('<div id="status-bar"><span class="sb-i" id="sb-ready">準備完了</span></div>', '<div id="status-bar"><span class="sb-i" id="sb-ready">準備完了</span></div>\n</div>')],
    ['lib/grid-stats.js', '★注意を 後ろに置く（375pxで画面の外＝読めない）★',
      (s) => s.replace('return 注意.concat(普通);', 'return 普通.concat(注意);')],
    ['lib/grid-stats.js', '★平均を 黙って丸める（≒ を付けない）★',
      (s) => s.replace("    return (r === n ? '' : '≒') + fmt(r);", '    return fmt(r);')],
    ['book.html', '★AIボタンを 帯の上へ退かさない（帯が読めなくなる）★',
      (s) => s.replace('bottom:calc(96px + var(--sb-h, 0px));', 'bottom:96px;')],
    ['book.html', '★箱ごとの折り返しをやめる（はみ出して 注意が読めない）★',
      (s) => s.replace('  flex-wrap:wrap; row-gap:2px;\n', '')],
    ['lib/grid-stats.js', "★'20 を 数として足す（うちの SUM と食い違う）★",
      (s) => s.replace("    if (isForcedText(cell)) return null;", '')],
    ['lib/grid-stats.js', '★Infinity を 黙って合計に入れる★',
      (s) => s.replace('      if (isNotComputed(t)) { 計算できていない++; continue; }', '')],
    ['lib/grid-stats.js', '★空を データの個数に入れる（実測は入らない）★',
      (s) => s.replace('      if (空) continue;                              /* ★空は数えない（実測④）★ */', '')],
    ['lib/grid-stats.js', '★文字を 平均の分母に入れる（実測は入らない）★',
      (s) => s.replace('      平均: 数を出せる ? (合計 / 数値の個数) : null,', '      平均: 数を出せる ? (合計 / データの個数) : null,')],
    ['lib/grid-stats.js', '★数字の形の文字を 足してしまう（実測は足さない）★',
      (s) => s.replace('    if (typeof raw !== \'string\') return false;\n    var s = raw.trim();', '    if (typeof raw !== \'string\') return false;\n    var s = raw.trim();\n    if (true) return false;')],
    ['lib/grid-stats.js', '★数の形の文字が在る事を 言わない（黙って小さい合計を出す）★',
      (s) => s.replace("      注意.push({ 名: '⚠️ 数の形の文字', 値: s.数の形の文字 + '個（足していません）', 注意: true });", '')],
    ['lib/grid-stats.js', '★エラーが在っても 合計を出す（実測は出ない）★',
      (s) => s.replace('(エラー === 0) && (計算できていない === 0) && ', '')],
    ['lib/grid-stats.js', '★出せない理由を 言わない★',
      (s) => s.replace("      出せない理由: エラー > 0 ? ('エラーが ' + エラー + '個 在るので 合計は出せません')", "      出せない理由: エラー > 0 ? ('')")],
    ['lib/grid-stats.js', '★論理値を 数として足す（実測は足さない）★',
      (s) => s.replace('    if (typeof raw === \'number\' && isFinite(raw)) return raw;', '    if (typeof raw === \'number\' && isFinite(raw)) return raw;\n    if (typeof raw === \'boolean\') return raw ? 1 : 0;')],
    ['lib/grid-stats.js', '★文字だけの時に 合計 0 を出す（0は嘘に見える）★',
      (s) => s.replace('      合計: 数を出せる ? 合計 : null,', '      合計: 合計,')],
    ['lib/grid-stats.js', '式のセルの 答えではなく 何も足さない',
      (s) => s.replace("    if (cell.f !== undefined && cell.f !== null && cell.f !== '') {", '    if (false) {')],
    ['lib/grid-stats.js', '数値の個数とデータの個数が同じでも 2つ並べる',
      (s) => s.replace("    if (s.数値の個数 !== s.データの個数) 普通.push(", "    if (true) 普通.push(")],
  ];
  let red = 0;
  for (const [rel, name, brk] of BREAKS) {
    const orig = fs.readFileSync(path.join(ROOT, rel), 'utf8');
    const bad = brk(orig);
    if (bad === orig) { console.log('  ★置換できず★  ' + name); continue; }
    const tmpFile = path.join(TMP, rel.replace(/[\\/]/g, '_'));
    fs.writeFileSync(tmpFile, bad, 'utf8');
    const env = Object.assign({}, process.env, { EXALLY_STATS_OVERRIDE: JSON.stringify({ [rel]: tmpFile }) });
    const isRed = spawnSync(process.execPath, [path.join(__dirname, 'grid-stats.test.mjs')], { encoding: 'utf8', env }).status !== 0;
    if (isRed) { red++; console.log('  赤くなった  ' + name); }
    else console.log('  ★素通り★  ' + name);
  }
  /* ★repo を書き換えていない事を 押した後に 数える★ */
  for (const rel of ['book.html', 'lib/grid-stats.js']) {
    const now = fs.readFileSync(path.join(ROOT, rel), 'utf8');
    if (now.includes(禁止色) || now.includes('if (true) return false;') || now.includes('white-space:normal;')) {
      console.log('  ★NG★ ' + rel + ' に わざと壊した物が残っている');
      process.exit(1);
    }
  }
  try { fs.rmSync(TMP, { recursive: true, force: true }); } catch (e) { /* 消せなくても検査は済んでいる */ }
  console.log('\n  ' + red + '/' + BREAKS.length + ' 通りで赤くなった');
  process.exit(red === BREAKS.length ? 0 : 1);
}

process.exit(fail ? 1 : 0);
