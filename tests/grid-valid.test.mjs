/* grid-valid.test.mjs — ★入力の決まり（データの入力規則）が 実Excelどおりか★
 *
 *  真値は tests/fixtures/excel-valid-golden.json（実Excel 16.0.20228 を COM で押して測った）。
 *  ★ソースを読むだけにしない★＝本物の book.html を載せて ★本物の入力欄に打つ★。
 *
 *  ★一番 大事な所★
 *    ・★止めるのは「打った時」だけ★（貼り付け・読み込みでは 止めない＝実Excelと同じ）
 *    ・★先に入っていた 合わない値は 消さない★（実Excelと同じ）
 *    ・★だが 黙らない＝合っていない値が何個 在るかを 数えて出す★（うちの上乗せ）
 *    ・★決まりが1つも無い時に「0個」と言わない★（未検査を 安全と言わない）
 *
 *  --self-test … わざと壊して「何通りで赤くなるか」を数える（★repo は読むだけ★）
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const SELF = process.argv.includes('--self-test');
const OVERRIDE = process.env.EXALLY_VALID_OVERRIDE ? JSON.parse(process.env.EXALLY_VALID_OVERRIDE) : {};
const srcPath = (rel) => OVERRIDE[rel] || path.join(ROOT, rel);

let JSDOM;
try { ({ JSDOM } = await import('jsdom')); }
catch { console.log('★jsdomが入っていません。この検証は飛ばせません（SKIPを緑と呼ばない）。npm install してください。'); process.exit(1); }

const GOLD = JSON.parse(fs.readFileSync(path.join(ROOT, 'tests/fixtures/excel-valid-golden.json'), 'utf8'));

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

console.log('\n[grid-valid] ★本物の book.html で 入力の決まりを 実際に押す★');
console.log('  真値 = ' + GOLD._measured_with);

const sh = () => win.sheets[win.activeSheet];
function reset() { sh().data = {}; sh().validations = {}; win.sel(0, 0, 0, 0); }
function toast(fn) {
  const 出た = []; const 元 = win.showToast;
  win.showToast = function (m) { 出た.push(String(m)); };
  try { fn(); } finally { win.showToast = 元; }
  return 出た;
}
/** ★実際に打つ★＝セルの入力欄に字を入れて commitEdit を呼ぶ（人がやる道と同じ） */
function 打つ(r, c, text) {
  win.sel(r, c, r, c);
  win.startEdit(r, c);
  doc.getElementById('cell-input').value = text;
  return toast(() => win.commitEdit());
}
/** 窓を開いて 決まりを付ける */
function 決まりを付ける(r1, c1, r2, c2, kind, a, b) {
  win.sel(r1, c1, r2, c2);
  win.openValid();
  doc.getElementById('validKind').value = kind;
  win.validKindChange();
  if (kind === 'list') doc.getElementById('validItems').value = a;
  else {
    doc.getElementById('validMin').value = a === undefined ? '' : String(a);
    doc.getElementById('validMax').value = b === undefined ? '' : String(b);
  }
  return toast(() => win.applyValid());
}
const 見えている字 = (r, c) => win.GridPrint.shown(sh().data[r + ',' + c]);

/* ── 土台 ── */
T('★画面が立ち上がっていて 入力の決まりの部品が 全部 在る', () => {
  ok(win.GridValid && typeof win.GridValid.check === 'function', 'GridValid が無い');
  ok(typeof win.openValid === 'function', 'openValid が無い');
  ok(typeof win.applyValid === 'function', 'applyValid が無い');
  ok(typeof win.clearValid === 'function', 'clearValid が無い');
  ok(typeof win.showBadCells === 'function', 'showBadCells が無い');
  ok(typeof win.validGate === 'function', 'validGate が無い');
  ok(doc.getElementById('validOverlay'), '窓が無い');
});
T('★右クリックに 入力の決まりが在る（出来ている物だけ出す）', () => {
  ok(doc.body.innerHTML.indexOf('openValid()') >= 0, '右クリックに入口が無い');
});

/* ── ① 実Excelの真値どおりか（純関数）── */
const G1 = GOLD['①一覧から選ぶ'], G2 = GOLD['②整数 1〜10'];
T('①一覧「りんご,みかん,ぶどう」… 合う値/合わない値が 実Excelと同じ', () => {
  const rule = { kind: 'list', items: win.GridValid.parseList(G1['付けた物']) };
  eq(win.GridValid.check(rule, 'りんご').ok, G1['合う値『りんご』'], 'りんご');
  eq(win.GridValid.check(rule, 'すいか').ok, G1['合わない値『すいか』'], 'すいか');
});
T('②整数1〜10 … 5/99/★5.5★ が 実Excelと同じ', () => {
  const rule = { kind: 'int', min: G2['下'], max: G2['上'] };
  eq(win.GridValid.check(rule, '5').ok, G2['5'], '5');
  eq(win.GridValid.check(rule, '99').ok, G2['99'], '99');
  eq(win.GridValid.check(rule, '5.5').ok, G2['5.5'], '★小数は整数の決まりに合わない★');
});
T('③空は「合っている」扱い（実測 IgnoreBlank=True）', () => {
  const rule = { kind: 'list', items: ['あ', 'い'] };
  eq(win.GridValid.check(rule, '').ok, GOLD['③空にすると']['結果'], '空');
  eq(win.GridValid.check(rule, '   ').ok, true, '空白だけ');
  eq(win.GridValid.check({ kind: 'int', min: 1, max: 10 }, '').ok, true, '数の決まりでも 空はOK');
});

/* ── ② 実際に打って 止まるか ── */
T('★決まりに合う字は 打てば 入る★', () => {
  reset();
  決まりを付ける(0, 0, 0, 0, 'list', 'りんご,みかん,ぶどう');
  打つ(0, 0, 'みかん');
  eq(見えている字(0, 0), 'みかん', 'A1');
});
T('★決まりに合わない字は 打っても 入らない＋理由が出る★', () => {
  reset();
  決まりを付ける(0, 0, 0, 0, 'list', 'りんご,みかん,ぶどう');
  打つ(0, 0, 'りんご');
  const 出た = 打つ(0, 0, 'すいか');
  eq(見えている字(0, 0), 'りんご', '★前の値のまま★');
  ok(出た.join('／').indexOf('入りません') >= 0, '理由が出ていない：' + JSON.stringify(出た));
  ok(出た.join('／').indexOf('りんご') >= 0, '選べる物を 出していない：' + JSON.stringify(出た));
});
T('★整数の決まり … 5.5 は 打っても 入らない★', () => {
  reset();
  決まりを付ける(1, 1, 1, 1, 'int', 1, 10);
  const 出た = 打つ(1, 1, '5.5');
  eq(見えている字(1, 1), '', '入ってしまった');
  ok(出た.join('／').indexOf('整数') >= 0, '理由が出ていない：' + JSON.stringify(出た));
  打つ(1, 1, '7');
  eq(見えている字(1, 1), '7', '合う数は入る');
});
T('★空は 打てる（決まりが在っても消せる）★', () => {
  reset();
  決まりを付ける(0, 0, 0, 0, 'list', 'あ,い');
  打つ(0, 0, 'あ');
  打つ(0, 0, '');
  eq(見えている字(0, 0), '', '空にできない');
});
T('★決まりの無いセルは 何でも打てる★', () => {
  reset();
  決まりを付ける(0, 0, 0, 0, 'list', 'あ,い');
  打つ(0, 1, 'なんでも');
  eq(見えている字(0, 1), 'なんでも', 'B1');
});

/* ── ③ ★止めるのは「打った時」だけ★（実Excelで測った通り）── */
T('★貼り付け・読み込みでは 止めない（実Excel と同じ）★', () => {
  reset();
  決まりを付ける(0, 0, 0, 0, 'list', 'りんご,みかん');
  win.setCell(0, 0, 'すいか');            /* 貼り付け・読み込みが通る道 */
  eq(見えている字(0, 0), 'すいか', '★Excel も入る。ここで止めたら 別物になる★');
});
T('★先に入っていた 合わない値は 決まりを足しても 消えない（実Excel と同じ）★', () => {
  reset();
  win.setCell(2, 2, 'はみ出た値');
  決まりを付ける(2, 2, 2, 2, 'list', 'あ,い,う');
  eq(見えている字(2, 2), GOLD['★一番 大事な実測★']['⑤先に値・後から決まり']['値'], '★値が消えている★');
});

/* ── ④ ★うちの上乗せ＝黙って通さない★ ── */
T('★決まりを足した時に「合っていない値が◯個」と数えて出す★', () => {
  reset();
  win.setCell(0, 0, 'すいか');
  win.setCell(1, 0, 'ばなな');
  const 出た = 決まりを付ける(0, 0, 1, 0, 'list', 'りんご,みかん');
  ok(出た.join('／').indexOf('2個') >= 0, '数を出していない：' + JSON.stringify(出た));
  ok(出た.join('／').indexOf('消していません') >= 0, '消していない事を言っていない：' + JSON.stringify(出た));
});
T('★合っていない値を数えて 最初の1つへ飛ぶ（何マス中 何個 も出す）★', () => {
  reset();
  決まりを付ける(0, 0, 2, 0, 'int', 1, 10);
  win.setCell(1, 0, '99');
  const s = toast(() => win.showBadCells()).join('／');
  ok(s.indexOf('1個') >= 0, '数が出ていない：' + s);
  ok(s.indexOf('3マス中') >= 0, '★何マス中 かを出していない★：' + s);
  ok(s.indexOf('A2') >= 0, '場所が出ていない：' + s);
  eq(win.selR1, 1, '飛んでいない');
});
T('★合っていない値が0個なら そう言う（黙らない）★', () => {
  reset();
  決まりを付ける(0, 0, 2, 0, 'int', 1, 10);
  win.setCell(1, 0, '5');
  const s = toast(() => win.showBadCells()).join('／');
  ok(s.indexOf('0個') >= 0, '0個と言っていない：' + s);
});
T('★決まりが1つも無い時は「まだ在りません」と言う（0個 と言わない）★', () => {
  reset();
  const s = toast(() => win.showBadCells()).join('／');
  ok(s.indexOf('まだ') >= 0, '未検査を 0個 と言っている：' + s);
});

/* ── ⑤ 窓の中身 ── */
T('★出来ていない物のボタンを出さない＝窓に在るのは 測った2つだけ★', () => {
  const 形 = [...doc.getElementById('validKind').options].map((o) => o.value);
  eq(形.join(','), 'list,int', '測っていない形を出している：' + 形.join(','));
});
T('★今 何の決まりが付いているかを 人の言葉で出す（確かめられる）★', () => {
  reset();
  決まりを付ける(0, 0, 0, 0, 'list', 'りんご,みかん');
  win.sel(0, 0, 0, 0); win.openValid();
  const t = doc.getElementById('validNow').textContent;
  ok(t.indexOf('りんご') >= 0 && t.indexOf('みかん') >= 0, '今の決まりが出ていない：' + t);
  win.closeValid();
});
T('★決まりが無いセルでは「外す」を出さない（押せない物を見せない）★', () => {
  reset();
  win.sel(5, 5, 5, 5); win.openValid();
  eq(doc.getElementById('validClearBtn').style.display, 'none', '外すボタンが出ている');
  win.closeValid();
  決まりを付ける(5, 5, 5, 5, 'int', 1, 10);
  win.sel(5, 5, 5, 5); win.openValid();
  ok(doc.getElementById('validClearBtn').style.display !== 'none', '外すボタンが出ていない');
  win.closeValid();
});
T('★中身が空のまま押すと 理由を出して 決まりを付けない★', () => {
  reset();
  win.sel(0, 0, 0, 0); win.openValid();
  doc.getElementById('validKind').value = 'list'; win.validKindChange();
  doc.getElementById('validItems').value = '';
  toast(() => win.applyValid());
  ok(doc.getElementById('validMsg').textContent.length > 0, '理由が出ていない');
  eq(Object.keys(sh().validations || {}).length, 0, '空なのに 決まりが付いた');
  win.closeValid();
});
T('★下が上より大きい時は 止めて 理由を出す★', () => {
  reset();
  win.sel(0, 0, 0, 0); win.openValid();
  doc.getElementById('validKind').value = 'int'; win.validKindChange();
  doc.getElementById('validMin').value = '10'; doc.getElementById('validMax').value = '1';
  toast(() => win.applyValid());
  ok(doc.getElementById('validMsg').textContent.indexOf('大きく') >= 0, '理由が出ていない');
  eq(Object.keys(sh().validations || {}).length, 0, '変な決まりが付いた');
  win.closeValid();
});
T('★決まりを外せる★', () => {
  reset();
  決まりを付ける(0, 0, 0, 0, 'list', 'あ,い');
  win.sel(0, 0, 0, 0); win.openValid();
  toast(() => win.clearValid());
  eq(win.validRuleAt(0, 0), null, '外れていない');
  打つ(0, 0, 'なんでも');
  eq(見えている字(0, 0), 'なんでも', '外したのに 止まる');
});
T('★範囲にまとめて付く（1マスずつではない）★', () => {
  reset();
  決まりを付ける(0, 0, 4, 1, 'int', 1, 10);
  eq(Object.keys(sh().validations).length, 10, '10マスに付いていない');
  ok(win.validRuleAt(4, 1), 'B5 に付いていない');
  ok(!win.validRuleAt(5, 0), 'A6 まで付いている');
});

/* ── ⑥ 一覧から選べる（Excel の セル内▼）── */
T('★一覧の決まりが在るセルでは 右クリックに 選ぶ物が出る★', () => {
  reset();
  決まりを付ける(0, 0, 0, 0, 'list', 'りんご,みかん');
  win.sel(0, 0, 0, 0);
  win.showCtxMenu(10, 10);
  const 出た = [...doc.querySelectorAll('#ctx-valid-list [data-valid-pick]')].map((e) => e.getAttribute('data-valid-pick'));
  eq(出た.join(','), 'りんご,みかん', '選ぶ物が出ていない：' + 出た.join(','));
  win.hideCtxMenu();
});
T('★決まりが無いセルでは 選ぶ物を出さない★', () => {
  reset();
  win.sel(3, 3, 3, 3);
  win.showCtxMenu(10, 10);
  eq(doc.querySelectorAll('#ctx-valid-list [data-valid-pick]').length, 0, '出ている');
  win.hideCtxMenu();
});
T('★選んだ物が そのまま入る★', () => {
  reset();
  決まりを付ける(0, 0, 0, 0, 'list', 'りんご,みかん');
  win.sel(0, 0, 0, 0);
  win.pickValidValue('みかん');
  eq(見えている字(0, 0), 'みかん', '入っていない');
});
T('★決まりが1つも無い時は「合っていない値を数える」を出さない★', () => {
  reset();
  win.sel(0, 0, 0, 0); win.showCtxMenu(10, 10);
  eq(doc.getElementById('ctx-bad-cells').style.display, 'none', '押せない物が出ている');
  決まりを付ける(0, 0, 0, 0, 'int', 1, 10);
  win.showCtxMenu(10, 10);
  ok(doc.getElementById('ctx-bad-cells').style.display !== 'none', '出ていない');
  win.hideCtxMenu();
});

/* ── ⑦ 前からある物を 押し直す（触った所の周り）── */
T('★前からある物：普通のセルは 今まで通り 打てる★', () => {
  reset();
  打つ(0, 0, '100');
  打つ(1, 0, '=A1*2');
  eq(見えている字(1, 0), '200', '式が計算されていない');
});
T('★前からある物：数式バーからも 今まで通り 入る★', () => {
  reset();
  win.sel(0, 0, 0, 0);
  doc.getElementById('formula-input').value = 'かきくけこ';
  toast(() => win.fkbCommit());
  eq(見えている字(0, 0), 'かきくけこ', '数式バーが効かない');
});
T('★数式バーでも 決まりで止まる★', () => {
  reset();
  決まりを付ける(0, 0, 0, 0, 'list', 'りんご,みかん');
  win.sel(0, 0, 0, 0);
  doc.getElementById('formula-input').value = 'すいか';
  const 出た = toast(() => win.fkbCommit());
  eq(見えている字(0, 0), '', '入ってしまった');
  ok(出た.join('／').indexOf('入りません') >= 0, '理由が出ていない');
});
/* ── ⑧ ★行や列を動かしたら 決まりも一緒に動く★（動かないと 別の行に付いたまま残る）── */
T('★行を入れると 決まりも 下へ動く★', () => {
  reset();
  決まりを付ける(2, 0, 2, 0, 'list', 'あ,い');
  win.sel(0, 0, 0, 0);
  win.ctxInsertRow();
  ok(!win.validRuleAt(2, 0), '★A3 に 決まりが残っている（動いていない）★');
  ok(win.validRuleAt(3, 0), 'A4 に 動いていない');
  打つ(3, 0, 'ううう');
  eq(見えている字(3, 0), '', '動いた先で 止まらない');
});
T('★行を消すと 決まりも 上へ動く★', () => {
  reset();
  決まりを付ける(3, 0, 3, 0, 'int', 1, 10);
  win.sel(0, 0, 0, 0);
  win.ctxDeleteRow();
  ok(!win.validRuleAt(3, 0), '★A4 に 決まりが残っている★');
  ok(win.validRuleAt(2, 0), 'A3 に 動いていない');
});
T('★消した行に付いていた決まりは 一緒に消える★', () => {
  reset();
  決まりを付ける(1, 0, 1, 0, 'int', 1, 10);
  win.sel(1, 0, 1, 0);
  win.ctxDeleteRow();
  eq(Object.keys(sh().validations).length, 0, '消えていない');
});
T('★列を入れると 決まりも 右へ動く★', () => {
  reset();
  決まりを付ける(0, 2, 0, 2, 'list', 'あ,い');
  win.sel(0, 0, 0, 0);
  win.ctxInsertCol();
  ok(!win.validRuleAt(0, 2), '★C1 に 決まりが残っている★');
  ok(win.validRuleAt(0, 3), 'D1 に 動いていない');
});
T('★列を消すと 決まりも 左へ動く★', () => {
  reset();
  決まりを付ける(0, 3, 0, 3, 'int', 1, 10);
  win.sel(0, 0, 0, 0);
  win.ctxDeleteCol();
  ok(!win.validRuleAt(0, 3), '★D1 に 決まりが残っている★');
  ok(win.validRuleAt(0, 2), 'C1 に 動いていない');
});
T('★取り消し（元に戻す）で 決まりも 戻る★', () => {
  reset();
  決まりを付ける(2, 0, 2, 0, 'list', 'あ,い');
  win.sel(0, 0, 0, 0);
  win.ctxInsertRow();
  win.doUndo();
  ok(win.validRuleAt(2, 0), '★戻っていない★');
  ok(!win.validRuleAt(3, 0), 'A4 に残っている');
});



console.log('\n  ' + pass + ' 緑 / ' + fail + ' 赤');

if (SELF) {
  const { spawnSync } = await import('node:child_process');
  const os = await import('node:os');
  const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'exally-valid-'));
  console.log('\n[self-test] わざと壊して 赤くなるかを数える（★repo は読むだけ★）');
  const BREAKS = [
    ['book.html', '打った時に 決まりを見ない（何でも入る）',
      (s) => s.replace('  if(!validGate(editingCell.r,editingCell.c,el.value)) return;\n', '')],
    ['book.html', '数式バーで 決まりを見ない',
      (s) => s.replace("  if(!validGate(selR1,selC1,document.getElementById('formula-input').value)) return;\n", '')],
    ['book.html', '★合っていない値の数を 出さない（黙る）★',
      (s) => s.replace("    + (bad.count ? '／★合っていない値が ' + bad.count + '個 在ります（消していません）★' : ''));", '  );')],
    ['book.html', '★決まりが無いのに「0個」と言う（未検査を安全と言う）★',
      (s) => s.replace("  if(!決まり数){ showToast('このシートには まだ 決まりが在りません'); return; }",
        "  if(!決まり数){ showToast('合っていない値は 0個でした'); return; }")],
    ['book.html', '★何マス中 かを出さない★',
      (s) => s.replace("  showToast('★合っていない値が ' + bad.count + '個★（' + 決まり数 + 'マス中）／'",
        "  showToast('★合っていない値が ' + bad.count + '個★／'")],
    ['book.html', '決まりを足す時に 合わない値を 消してしまう（Excelは消さない）',
      (s) => s.replace('  var bad = GridValid.countBad(sh.data, sh.validations, GridPrint.shown);',
        '  var bad = GridValid.countBad(sh.data, sh.validations, GridPrint.shown);\n  for(var bi=0;bi<bad.cells.length;bi++) setCell(bad.cells[bi].r,bad.cells[bi].c,"");')],
    ['book.html', '一覧の決まりでも 右クリックに 選ぶ物を出さない',
      (s) => s.replace("  if(vr && vr.kind==='list'){", '  if(false && vr){')],
    ['book.html', '決まりが無くても「合っていない値を数える」を出す',
      (s) => s.replace("  document.getElementById('ctx-bad-cells').style.display = 決まり数 ? 'block' : 'none';",
        "  document.getElementById('ctx-bad-cells').style.display = 'block';")],
    ['book.html', '外すボタンを いつも出す（押せない物を見せる）',
      (s) => s.replace("  document.getElementById('validClearBtn').style.display = rule ? '' : 'none';",
        "  document.getElementById('validClearBtn').style.display = '';")],
    ['book.html', '★測っていない形（日付）のボタンを足す★',
      (s) => s.replace('<option value="int">整数の範囲</option>', '<option value="int">整数の範囲</option><option value="date">日付</option>')],
    ['book.html', '空の一覧でも 決まりを付けてしまう',
      (s) => s.replace("    if(!items.length){ document.getElementById('validMsg').textContent = '選ぶ物を カンマで区切って書いてね（例：りんご,みかん）'; return; }", '')],
    ['book.html', '下が上より大きくても 通してしまう',
      (s) => s.replace("    if(mn!=='' && mx!=='' && Number(mn) > Number(mx)){\n      document.getElementById('validMsg').textContent = '下が 上より大きくなっています'; return; }", '')],
    ['book.html', '範囲ではなく 1マスにしか付けない',
      (s) => s.replace('sh.validations = GridValid.setRange(_validRules(), selR1, selC1, selR2, selC2, rule);',
        'sh.validations = GridValid.setRange(_validRules(), selR1, selC1, selR1, selC1, rule);')],
    ['book.html', '一覧から選んだ物を 入れない',
      (s) => s.replace('function pickValidValue(t){\n  hideCtxMenu();\n  setCell(selR1,selC1,t);', 'function pickValidValue(t){\n  hideCtxMenu();')],
    ['book.html', '★行を入れても 決まりを動かさない（別の行に付いたまま）★',
      (s) => s.replace('  _moveValidations(s, true, true);\n', '')],
    ['book.html', '★行を消しても 決まりを動かさない★',
      (s) => s.replace('  _moveValidations(s, true, false);\n', '')],
    ['book.html', '★列を入れても 決まりを動かさない★',
      (s) => s.replace('  _moveValidations(s, false, true);\n', '')],
    ['book.html', '★列を消しても 決まりを動かさない★',
      (s) => s.replace('  _moveValidations(s, false, false);\n', '')],
    ['book.html', '消した行の決まりを 残してしまう',
      (s) => s.replace('    else { if(n >= 上 && n <= 下) return; if(n > 下) n -= 幅; }', '    else { if(n > 下) n -= 幅; }')],
    ['lib/grid-valid.js', '★空を 合わない扱いにする（実測は 合っている）★',
      (s) => s.replace("    if (isBlank(text)) return { ok: true, why: '' };", '')],
    ['lib/grid-valid.js', '★小数を 整数の決まりで 通してしまう（実測は 通らない）★',
      (s) => s.replace("      if (rule.kind === 'int' && Math.floor(n) !== n) {", '      if (rule.kind === "___" && Math.floor(n) !== n) {')],
    ['lib/grid-valid.js', '★上の数を 見ない★',
      (s) => s.replace("      if (上 !== null && n > 上) return { ok: false, why: 上 + ' 以下を入れてください' };", '')],
    ['lib/grid-valid.js', '★理由に 選べる物を出さない★',
      (s) => s.replace("      return { ok: false, why: 'この列は ' + items.join('／') + ' から選びます' };", "      return { ok: false, why: 'だめです' };")],
    ['lib/grid-valid.js', '合っていない値を 数え落とす',
      (s) => s.replace('      if (!res.ok) cells.push({ r: r, c: c, text: t, why: res.why });', '')],
    ['lib/grid-valid.js', '決まりの中身を 人の言葉で出さない',
      (s) => s.replace("    if (rule.kind === 'list') return '一覧から選ぶ：' + (rule.items || []).join('／');", "    if (rule.kind === 'list') return '';")],
  ];
  let red = 0;
  for (const [rel, name, brk] of BREAKS) {
    const orig = fs.readFileSync(path.join(ROOT, rel), 'utf8');
    const bad = brk(orig);
    if (bad === orig) { console.log('  ★置換できず★  ' + name); continue; }
    const tmpFile = path.join(TMP, rel.replace(/[\\/]/g, '_'));
    fs.writeFileSync(tmpFile, bad, 'utf8');
    const env = Object.assign({}, process.env, { EXALLY_VALID_OVERRIDE: JSON.stringify({ [rel]: tmpFile }) });
    const isRed = spawnSync(process.execPath, [path.join(__dirname, 'grid-valid.test.mjs')], { encoding: 'utf8', env }).status !== 0;
    if (isRed) { red++; console.log('  赤くなった  ' + name); }
    else console.log('  ★素通り★  ' + name);
  }
  /* ★repo を書き換えていない事を 押した後に 数える★ */
  for (const rel of ['book.html', 'lib/grid-valid.js']) {
    const now = fs.readFileSync(path.join(ROOT, rel), 'utf8');
    if (now.includes('<option value="date">') || now.includes('if(false && vr)') || now.includes('rule.kind === "___"')) {
      console.log('  ★NG★ ' + rel + ' に わざと壊した物が残っている');
      process.exit(1);
    }
  }
  try { fs.rmSync(TMP, { recursive: true, force: true }); } catch (e) { /* 消せなくても検査は済んでいる */ }
  console.log('\n  ' + red + '/' + BREAKS.length + ' 通りで赤くなった');
  process.exit(red === BREAKS.length ? 0 : 1);
}

process.exit(fail ? 1 : 0);
