/* book-scan-ui.test.mjs — ★調べている間の知らせ（ぐるぐる＋何枚目＋％）★
 *
 *  ★司さん 2026-08-25「読み込み中は ぐるぐる／何% を出せ」★
 *    ・人は 画面が変わらないと「固まった」と思う
 *    ・★300ミリ秒を超えたら出す★（それ未満で出すと チカチカして かえって不安になる）
 *    ・★終わったら 必ず消す★（消し忘れが一番 怖い）
 *    ・★失敗した時も 必ず消して 何が起きたか1行 出す★
 *    ・★[hidden]{display:none!important} を1行★（class の display に負けて
 *      「中身が空の枠だけ」が残る事故が 他アプリで2回）
 *    ・★注意書きを flex/grid の箱に入れると 1文字ずつ縦に割れる★＝折り返さない指定を確かめる
 *
 *  使い方: node tests/book-scan-ui.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { 注記を外す } from '../scripts/lib/chuki.mjs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const SELF = process.argv.includes('--self-test');
const OVERRIDE = process.env.EXALLY_SCAN_OVERRIDE ? JSON.parse(process.env.EXALLY_SCAN_OVERRIDE) : {};
const 読む = (rel) => fs.readFileSync(OVERRIDE[rel] || path.join(ROOT, rel), 'utf8');

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ok   ' + n); } catch (e) { fail++; console.log('  NG   ' + n + '\n       ' + (e && e.message)); } };
const ok = (c, m) => { if (!c) throw new Error(m || 'expected truthy'); };
const eq = (a, b, m) => { if (a !== b) throw new Error((m ? m + ': ' : '') + '期待=' + JSON.stringify(b) + ' 実際=' + JSON.stringify(a)); };

console.log('');
console.log('[book-scan-ui] ★調べている間の知らせ★');

const book = 読む('book.html');

T('★知らせの箱が 画面に在る（最初は hidden）★', () => {
  ok(/<div id="bookScan" hidden>/.test(book), '★箱が無い、または 最初から出ている★');
  ok(book.indexOf('id="bookScanTxt"') > 0, '字を入れる所が無い');
  ok(/class="bs-s"/.test(book), 'ぐるぐるが無い');
});
T('★[hidden] を class の display に負けさせない1行が 在る★', () => {
  ok(/#bookScan\[hidden\]\s*\{\s*display\s*:\s*none\s*!important/.test(book),
    '★この1行が無いと「中身が空の枠だけ」が残る（他アプリで2回 事故）★');
});
T('★字が 1文字ずつ縦に割れない指定（折り返さない）★', () => {
  ok(/#bookScan \.bs-t\{[^}]*white-space\s*:\s*nowrap/.test(book),
    '★flex の箱の中で 字が縦に割れる★');
});
T('★300ミリ秒を超えてから 出す（それ未満では出さない）★', () => {
  ok(/setTimeout\(function\(\)\{[\s\S]{0,120}bookScan[\s\S]{0,120}\}, 300\)/.test(book)
    || /\}, 300\);/.test(book), '★300ミリ秒の待ちが無い＝すぐ出てチカチカする★');
  ok(book.indexOf('出した = true') > 0, '出したかどうかを 持っていない');
});
T('★終わったら 必ず消す／失敗しても 消す★', () => {
  ok(/function 消す\(\)\{[\s\S]{0,200}hidden = true/.test(book), '消す物が無い');
  const 網の所0 = book.slice(book.indexOf('function 参照の網を作り始める'), book.indexOf('function _sheetsReferencedBy'));
  const 出来た所 = 網の所0.slice(網の所0.indexOf('if(r.done){'), 網の所0.indexOf('if(r.done){') + 200);
  ok(出来た所.indexOf('消す();') >= 0, '★終わった時に 消していない★');
  /* ★7,700行の中で最初の catch を見ていた（検査の側の間違い・2026-08-25）★
     ⇒ ★網を作る所★の中だけを見る */
  const 網の所 = book.slice(book.indexOf('function 参照の網を作り始める'), book.indexOf('function _sheetsReferencedBy'));
  /* ★400字も見ると 次の「出来た時」の 消す() まで拾ってしまう（自己確認で素通りした）★
     ⇒ ★catch の中だけ★を切る（次の行 if(r.done) の手前まで） */
  const c0 = 網の所.indexOf('catch(e){');
  const 失敗所 = 網の所.slice(c0, 網の所.indexOf('if(r.done){', c0));
  ok(失敗所.indexOf('消す();') >= 0, '★失敗した時に 消していない（枠が残る）★');
});
T('★失敗した時は 何が起きたか1行 出す（黙って消えない）★', () => {
  const i = book.indexOf('ブック全体は調べられませんでした');
  ok(i > 0, '★失敗を 客に伝えていない★');
  const 前後 = book.slice(i - 200, i + 300);
  ok(前後.indexOf('見ているシートの計算は そのまま使えます') >= 0, '★次にどうなるかを 言っていない★');
});
T('★何枚目／％を 出す（何が起きているかを 言葉でも）★', () => {
  ok(book.indexOf('枚目（') > 0, '★何枚目を 出していない★');
  ok(/割 = Math\.min\(99/.test(book), '★100%のまま止まって見える（99で止める）★');
});
T('★客に見せる字に ★ を書かない★', () => {
  const 出す字 = [];
  const re = /知らせ字\.textContent = ([^;]+);/g;
  let m;
  while ((m = re.exec(book))) 出す字.push(m[1]);
  ok(出す字.length >= 1, '出す字が見つからない');
  for (const t of 出す字) ok(t.indexOf('★') < 0, '★客の字に ★ が出ている★：' + t);
});

/* ★2026-08-25 実際に撮ってみたら、★終わった時の知らせ★に ★ が出ていた★
   ＝この検査は「調べている間」の字しか見ていなかった。⇒ ★画面に出る知らせ 全部★を見る。 */
T('★客に見せる知らせ 全部に ★ を書かない（showToast をひとつ残らず）★', () => {
  const 出す = [];
  let i = 0;
  while ((i = book.indexOf('showToast(', i)) >= 0) {
    let d = 0, j = i + 'showToast'.length;
    for (; j < book.length; j++) {
      const c = book[j];
      if (c === '(') d++;
      else if (c === ')') { d--; if (d === 0) break; }
    }
    出す.push(book.slice(i, j));
    i = j;
  }
  ok(出す.length >= 10, '知らせが見つからない（' + 出す.length + '件）');
  const 悪い = [];
  for (const t of 出す) {
    /* 文字列の中だけを見る（変数名や注記は 客に出ない） */
    const 字 = (t.match(/'[^']*'/g) || []).join('');
    if (字.indexOf('★') >= 0) 悪い.push(字.slice(0, 60));
  }
  eq(悪い.length, 0, '★客の知らせに ★ が出ている：' + 悪い.join(' ／ '));
  console.log('       … 知らせ ' + 出す.length + '件を 数えた');
});

/* ★同じ回で もう1つ踏んだ：知らせの出口は2つ在る（showToast=HTML／_findMsg=字だけ）★
   ＝_findMsg に <b> を書くと 画面に「<b>」と そのまま出る。 */
T('★字だけの知らせ(_findMsg)に タグを書かない★', () => {
  const 悪い = [];
  let i = 0;
  while ((i = book.indexOf('_findMsg(', i)) >= 0) {
    let d = 0, j = i + '_findMsg'.length;
    for (; j < book.length; j++) {
      const c = book[j];
      if (c === '(') d++;
      else if (c === ')') { d--; if (d === 0) break; }
    }
    const 字 = (book.slice(i, j).match(/'[^']*'/g) || []).join('');
    if (/<[a-zA-Z/]/.test(字)) 悪い.push(字.slice(0, 60));
    i = j;
  }
  eq(悪い.length, 0, '★字だけの所に タグが出ている：' + 悪い.join(' ／ '));
});

/* ── 実際に動かす（jsdom）── */
let JSDOM;
try { ({ JSDOM } = await import('jsdom')); }
catch { console.log('★jsdomが入っていません。この検証は飛ばせません'); process.exit(1); }

const 画面 = () => {
  const dom = new JSDOM('<!doctype html><html><head><style>' +
    '#bookScan[hidden]{display:none!important;}</style></head><body>' +
    '<div id="bookScan" hidden><span class="bs-s"></span><span class="bs-t" id="bookScanTxt">ブック全体を調べています…</span></div>' +
    '</body></html>', { runScripts: 'dangerously', pretendToBeVisual: true });
  return dom.window;
};
T('★hidden を外すと 見える／戻すと 消える（class に負けていない）★', () => {
  const w = 画面();
  const el = w.document.getElementById('bookScan');
  eq(w.getComputedStyle(el).display, 'none', '最初から出ている');
  el.hidden = false;
  ok(w.getComputedStyle(el).display !== 'none', '★hidden を外しても 見えない★');
  el.hidden = true;
  eq(w.getComputedStyle(el).display, 'none', '★消えていない（枠が残る）★');
});

console.log('');
console.log('  ' + pass + ' 緑 / ' + fail + ' 赤');

if (SELF) {
  const { spawnSync } = await import('node:child_process');
  const os = await import('node:os');
  const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'exally-scan-'));
  console.log('');
  console.log('[self-test] わざと壊して 赤くなるかを数える（★repo は読むだけ★）');
  const BREAKS = [
    ['★知らせの箱を 消す★', (s) => s.replace('<div id="bookScan" hidden>', '<div id="bookScanX" hidden>')],
    ['★最初から 出しておく★', (s) => s.replace('<div id="bookScan" hidden>', '<div id="bookScan">')],
    ['★[hidden] の1行を 消す★', (s) => s.replace('#bookScan[hidden]{display:none!important;}', '')],
    ['★字を 折り返させる（縦に割れる）★', (s) => s.replace('#bookScan .bs-t{font-size:13px;color:#333333;line-height:1.6;white-space:nowrap;}', '#bookScan .bs-t{font-size:13px;color:#333333;line-height:1.6;}')],
    ['★終わっても 消さない★', (s) => s.replace('    if(r.done){\n      消す();', '    if(r.done){')],
    ['★失敗した時に 消さない★', (s) => s.replace('      消す(); _refGraphBusy = false;', '      _refGraphBusy = false;')],
    ['★失敗を 黙る★', (s) => s.replace('ブック全体は調べられませんでした', '')],
    ['★何枚目を 出さない★', (s) => s.replace('枚目（', '（')],
    ['★100%で止まって見える（99で止めない）★', (s) => s.replace('Math.min(99,', 'Math.min(100,')],
    ['★客の字に ★ を書く★', (s) => s.replace("'ブック全体を調べています… '", "'★ブック全体を調べています…★ '")],
    ['★終わった時の知らせに ★ を書く★', (s) => s.replace("'本／別のシートを見ている式 '", "'本／★別のシートを見ている式 '")],
    ['★字だけの知らせに タグを書く★', (s) => s.replace("'／うち 式が '", "'／<b>うち 式が '")],
  ];
  let red = 0;
  for (const [name, brk] of BREAKS) {
    const orig = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
    const bad = brk(orig);
    if (bad === orig) { console.log('  ★置換できず★  ' + name); continue; }
    const f = path.join(TMP, 'book.html');
    fs.writeFileSync(f, bad, 'utf8');
    const env = Object.assign({}, process.env, { EXALLY_SCAN_OVERRIDE: JSON.stringify({ 'book.html': f }) });
    const r = spawnSync(process.execPath, ['--max-old-space-size=4096', path.join(__dirname, 'book-scan-ui.test.mjs')], { encoding: 'utf8', env });
    if (r.status !== 0) { red++; console.log('  赤くなった  ' + name); }
    else console.log('  ★素通り★  ' + name);
  }
  try { fs.rmSync(TMP, { recursive: true, force: true }); } catch (e) { /* 消せなくても検査は済んでいる */ }
  console.log('');
  console.log('  ' + red + '/' + BREAKS.length + ' 通りで赤くなった');
  process.exit(red === BREAKS.length ? 0 : 1);
}

process.exit(fail ? 1 : 0);
