/* shindan-ui.test.mjs — ★本物の book.html で 診断の知らせと一覧を 実際に押す★
 *
 *  ★部品が緑＝画面で使える ではない★（条件付き書式で 一度 踏んだ）。
 *  ここでは ★本物の画面を読み込んで、本物の 診断を始める() を動かす★。
 *
 *  ★同じ回で もう1つ★ … 直した所の控え（書き出した瞬間に 捨てていた物）
 *
 *  使い方: node tests/shindan-ui.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const SELF = process.argv.includes('--self-test');
const OVERRIDE = process.env.EXALLY_SHINDANUI_OVERRIDE ? JSON.parse(process.env.EXALLY_SHINDANUI_OVERRIDE) : {};
const 読む = (rel) => fs.readFileSync(OVERRIDE[rel] || path.join(ROOT, rel), 'utf8');

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ok   ' + n); } catch (e) { fail++; console.log('  NG   ' + n + '\n       ' + (e && e.message)); } };
const ok = (c, m) => { if (!c) throw new Error(m || 'expected truthy'); };
const eq = (a, b, m) => { if (a !== b) throw new Error((m ? m + ': ' : '') + '期待=' + JSON.stringify(b) + ' 実際=' + JSON.stringify(a)); };

const book = 読む('book.html');

console.log('');
console.log('[shindan-ui] ★診断の知らせ・一覧・控え（本物の画面で押す）★');

/* ══ 出来ていない物のボタンを見せない ══ */
T('★ボタンは 最初は 出ていない（見つかった時だけ 出す）★', () => {
  ok(/<button id="shindanBtn" hidden/.test(book), '★最初から 出ている（何も無いのに押せる）★');
});
T('★[hidden] を class の display に負けさせない1行★', () => {
  ok(/#shindanBtn\[hidden\]\s*\{\s*display\s*:\s*none\s*!important/.test(book),
    '★この1行が無いと「中身が空の枠だけ」が残る（他アプリで2回 事故）★');
});
T('★見つからなかった時は 何も出さない（0件で騒がない）★', () => {
  ok(/if\(_shindanResult\.式の本数 > 0\) 診断の知らせを出す/.test(book), '★0件でも 知らせている★');
  ok(/if\(!r \|\| !r\.式の本数\) return;/.test(book), '★中身が無くても 窓が開く★');
});
T('★調べられなくても 表は そのまま使える（落ちない）★', () => {
  const i = book.indexOf('function 診断を始める');
  const 所 = book.slice(i, i + 900);
  ok(/catch\(e\)\{ _shindanBusy = false; return; \}/.test(所), '★失敗したら 画面ごと止まる★');
});
T('★小分けで調べる（画面を固めない）★', () => {
  ok(/Shindan\.調べる途中\(sheets, \{ 一度に: 3000 \}\)/.test(book), '★一度に全部 やっている★');
  ok(/requestAnimationFrame \|\| window\.setTimeout/.test(book), '★次の小分けへ 渡していない★');
});
T('★AIを使っていない事を 客に言う（0円の道だと伝える）★', () => {
  ok(book.indexOf('AIは使っていません') > 0, '★0円だと 言っていない★');
});
T('★客に見せる字に ★ を書かない（診断の知らせ）★', () => {
  const i = book.indexOf('function 診断の知らせを出す');
  const 所 = book.slice(i, book.indexOf('function openShindan'));
  const 字 = (所.replace(/\/\*[\s\S]*?\*\//g, '').match(/'[^']*'/g) || []).join('');
  ok(字.indexOf('★') < 0, '★客の字に ★ が出ている★：' + 字.slice(0, 80));
});
T('★多い時は 何件 出していないかを 書く（黙って切らない）★', () => {
  ok(book.indexOf('ほか ') > 0 && /出していません/.test(book), '★黙って 先頭だけ出している★');
});

/* ══ ★直した所の控え（捨てる前に写す）★ ══ */
T('★書き出した後に 捨てる前に 控えへ写す★', () => {
  const i = book.indexOf('控えへ写す(outName, _editedCells, plan.total);');
  ok(i > 0, '★写していない（捨てているだけ）★');
  const j = book.indexOf('_editedCells = {};', i);
  ok(j > i && j - i < 300, '★写す前に 捨てている（順番が逆）★');
});
T('★別のファイルを開く時も 捨てる前に写す★', () => {
  ok(/控えへ写す\(\(BookOpen\.current\(\) \|\| \{\}\)\.name/.test(book), '★開き直しで 消えている★');
});
T('★控えは 際限なく溜めない（古い方から捨てる）★', () => {
  ok(/控えの上限 = 200/.test(book), '★上限が無い★');
  ok(/while\(直した控え\.length > 控えの上限\) 直した控え\.shift\(\);/.test(book), '★新しい方を捨てている／捨てていない★');
});
T('★控えで 本体を落とさない★', () => {
  const i = book.indexOf('function 控えへ写す');
  const 所 = book.slice(i, i + 900);
  ok(/catch\(e\)\{ return null; \}/.test(所), '★控えの失敗で 書き出しごと止まる★');
});

/* ══ 実際に動かす（jsdom で 本物の book.html を読む）══ */
let JSDOM;
try { ({ JSDOM } = await import('jsdom')); }
catch { console.log('★jsdomが入っていません。この検証は飛ばせません'); process.exit(1); }

const 画面を作る = () => {
  const dom = new JSDOM('<!doctype html><html><head></head><body>'
    + '<button id="shindanBtn" hidden></button>'
    + '<div id="dgOverlay" style="display:none"><div id="dgTitle"></div><div id="dgBody"></div>'
    + '<div id="dgNext"></div><div id="dgList"></div><div id="dgMore"></div></div>'
    + '</body></html>', { runScripts: 'dangerously', pretendToBeVisual: true });
  const w = dom.window;
  w.eval(読む('lib/shindan.js'));
  /* 本物の book.html から ★診断の所だけ★ 切り出して そのまま動かす */
  const src = book.slice(book.indexOf('var _shindanResult = null;'), book.indexOf('function 参照の網を作り始める'));
  w.eval('var sheets = [], activeSheet = 0, selR1=0,selR2=0,selC1=0,selC2=0;'
    + 'var 出た知らせ = [];'
    + 'function showToast(h){ 出た知らせ.push(h); }'
    + 'function switchSheet(i){ activeSheet = i; }'
    + 'var 行った先 = null, 動かした回数 = 0;'
    + 'function scrollToCell(r,c){ 行った先 = r + "," + c; 動かした回数++; }'
    + 'function updateBar(){} function render(){}'
    + src);
  return w;
};

T('★見つかった時だけ ボタンが出て、押すと 一覧が開く（本物の関数を動かす）★', () => {
  const w = 画面を作る();
  w.sheets = [{ name: '計算', data: {
    '0,0': { f: '=IFERROR(#REF!,"")', v: '' },
    '1,0': { f: '=A1+1', v: 2 },
    '2,3': { f: '=IFERROR(INDEX(#REF!,MATCH(B3,月別!A1:A9,0)),"")', v: '' },
  } }];
  eq(w.document.getElementById('shindanBtn').hidden, true, '最初は出ていない');
  w.診断を始める();
  /* 小分けは rAF で回るので 1回ぶん進める */
  for (let i = 0; i < 5; i++) { if (w._shindanResult) break; }
  ok(w._shindanResult, '★調べ終わっていない★');
  eq(w._shindanResult.式の本数, 2);
  eq(w.document.getElementById('shindanBtn').hidden, false, '★見つかったのに ボタンが出ない★');
  eq(w.document.getElementById('shindanBtn').textContent, '危ない所 2か所');
  ok(w.出た知らせ.length === 1, '★知らせが 出ていない／2回 出ている★');
  ok(w.出た知らせ[0].indexOf('2か所') > 0, '★何か所かを 言っていない★');
  ok(w.出た知らせ[0].indexOf('★') < 0, '★客の字に ★ が出ている★');

  w.openShindan();
  eq(w.document.getElementById('dgOverlay').style.display, 'flex', '★窓が 開かない★');
  const 行 = w.document.getElementById('dgList').children;
  eq(行.length, 2, '★一覧の行数が 合わない★');
  ok(行[0].textContent.indexOf('計算 の A1') >= 0, '★場所を 出していない★：' + 行[0].textContent);
  ok(行[0].textContent.indexOf('（空）') >= 0, '★今 何が出ているかを 出していない★');
  ok(行[1].textContent.indexOf('D3') >= 0, '★2件目の場所が 違う★：' + 行[1].textContent);
});
T('★一覧を押すと その場所へ行く（見に行けないと 意味がない）★', () => {
  const w = 画面を作る();
  w.sheets = [{ name: 'あ', data: {} }, { name: '計算', data: { '4,2': { f: '=IFERROR(#REF!,"")', v: '' } } }];
  w.診断を始める();
  w.openShindan();
  w.document.getElementById('dgList').children[0].onclick();
  eq(w.activeSheet, 1, '★別のシートへ 移っていない★');
  eq(w.selR1, 4, '★行が 合っていない★');
  eq(w.selC1, 2, '★列が 合っていない★');
  /* ★選んだだけでは 画面の外のままで 見えない★＝実際に そこへ動かしたかを見る
     （わざと壊しても 素通りしていた・2026-08-25 検査の側の穴） */
  /* ★自分で押して見つけた★＝端に貼り付くと 浮いているボタンに隠れる。
     ⇒ ★先に 6行 下へ動かしてから 戻す★（最後は 行き先そのもの） */
  eq(w.行った先, '4,2', '★その場所まで 画面を動かしていない（選んだだけ）★');
  eq(w.動かした回数, 2, '★端に貼り付く（下のボタンに隠れる）★');
  eq(w.document.getElementById('dgOverlay').style.display, 'none', '★窓が 開いたまま★');
});
T('★見つからなければ ボタンも 知らせも 出ない★', () => {
  const w = 画面を作る();
  w.sheets = [{ name: '計算', data: { '0,0': { f: '=A1+1', v: 2 }, '1,0': { f: '=IFERROR(A1,"")', v: 2 } } }];
  w.診断を始める();
  eq(w.document.getElementById('shindanBtn').hidden, true, '★何も無いのに ボタンが出た★');
  eq(w.出た知らせ.length, 0, '★何も無いのに 知らせが出た★');
  w.openShindan();
  eq(w.document.getElementById('dgOverlay').style.display, 'none', '★何も無いのに 窓が開いた★');
});

console.log('');
console.log('  ' + pass + ' 緑 / ' + fail + ' 赤');

if (SELF) {
  const { spawnSync } = await import('node:child_process');
  const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'exally-dgui-'));
  console.log('');
  console.log('[self-test] わざと壊して 赤くなるかを数える（★repo は読むだけ★）');
  const BREAKS = [
    ['book.html', '★最初から ボタンを出す★', (s) => s.replace('<button id="shindanBtn" hidden', '<button id="shindanBtn"')],
    ['book.html', '★[hidden] の1行を 消す★', (s) => s.replace('#shindanBtn[hidden]{display:none!important;}', '')],
    ['book.html', '★0件でも 知らせる★', (s) => s.replace('if(_shindanResult.式の本数 > 0) 診断の知らせを出す(_shindanResult);', '診断の知らせを出す(_shindanResult);')],
    ['book.html', '★中身が無くても 窓を開く★', (s) => s.replace('if(!r || !r.式の本数) return;', 'r = r || { 見つけた: [], 式の本数: 0 };')],
    ['book.html', '★調べられない時に 画面ごと止まる★', (s) => s.replace('catch(e){ _shindanBusy = false; return; }', 'catch(e){ throw e; }')],
    ['book.html', '★一度に全部 やる（画面が固まる）★', (s) => s.replace('Shindan.調べる途中(sheets, { 一度に: 3000 })', 'Shindan.調べる途中(sheets, { 一度に: 1e9 })')],
    ['book.html', '★0円だと 言わない★', (s) => s.replace('AIは使っていません', '')],
    ['book.html', '★場所を 出さない★', (s) => s.replace("頭.textContent = x.シート + ' の ' + x.セル + '　いま出ている物：' + x.いま出ている物;", "頭.textContent = 'ここが危ない';")],
    ['book.html', '★押しても 場所へ行かない★', (s) => s.replace('scrollToCell(x.r, x.c); updateBar(); render();', '')],
    ['book.html', '★端に貼り付く（浮いているボタンに隠れる）★',
      (s) => s.replace('  scrollToCell(Math.min(x.r + 6, 最終行 || (x.r + 6)), x.c);', '')],
    ['book.html', '★窓を閉じない（一覧の下に隠れて セルが見えない）★', (s) => s.replace('function 診断の場所へ行く(x){\n  closeShindan();', 'function 診断の場所へ行く(x){')],
    ['book.html', '★捨てる前に 写さない（今までの形に戻す）★', (s) => s.replace('      控えへ写す(outName, _editedCells, plan.total);', '')],
    ['book.html', '★開き直しで 消す★', (s) => s.replace("    控えへ写す((BookOpen.current() || {}).name || '（名前なし）', _editedCells, 0);", '')],
    ['book.html', '★控えに 上限を付けない★', (s) => s.replace('while(直した控え.length > 控えの上限) 直した控え.shift();', '')],
    ['book.html', '★控えの失敗で 書き出しごと止まる★', (s) => s.replace('  }catch(e){ return null; }        /* ★控えで 本体を落とさない★ */', '  }catch(e){ throw e; }')],
    ['book.html', '★黙って 先頭だけ出す★', (s) => s.replace("('ほか ' + (r.見つけた.length - 出す) + 'か所は 出していません（多いので 先頭' + 出す + 'か所）。')", "''")],
  ];
  let red = 0;
  for (const [rel, name, brk] of BREAKS) {
    const orig = fs.readFileSync(path.join(ROOT, rel), 'utf8');
    let bad = brk(orig);
    if (bad === orig) bad = brk(orig.replace(/\r\n/g, '\n'));   /* ★改行の違いで 素通りしない★ */
    if (bad === orig || bad === orig.replace(/\r\n/g, '\n')) { console.log('  ★置換できず★  ' + name); continue; }
    const f = path.join(TMP, 'book.html');
    fs.writeFileSync(f, bad, 'utf8');
    const env = Object.assign({}, process.env, { EXALLY_SHINDANUI_OVERRIDE: JSON.stringify({ 'book.html': f }) });
    const r = spawnSync(process.execPath, ['--max-old-space-size=4096', path.join(__dirname, 'shindan-ui.test.mjs')], { encoding: 'utf8', env });
    if (r.status !== 0) { red++; console.log('  赤くなった  ' + name); }
    else console.log('  ★素通り★  ' + name);
  }
  try { fs.rmSync(TMP, { recursive: true, force: true }); } catch (e) { /* 消せなくても検査は済んでいる */ }
  console.log('');
  console.log('  ' + red + '/' + BREAKS.length + ' 通りで赤くなった');
  process.exit(red === BREAKS.length ? 0 : 1);
}

process.exit(fail ? 1 : 0);
