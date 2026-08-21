/* cond-format-ui.test.mjs — ★本物の book.html を載せて 条件付き書式を 実際に押す★
 *
 *  ★「部品が緑」は「画面で使える」ではない★（配線を毎回 押して確かめる）
 *  真値 = tests/fixtures/cond-format-golden.json（実Excelから取った物）
 *
 *  使い方: node tests/cond-format-ui.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const SELF = process.argv.includes('--self-test');
const OVERRIDE = process.env.EXALLY_CFUI_OVERRIDE ? JSON.parse(process.env.EXALLY_CFUI_OVERRIDE) : {};
const srcPath = (rel) => OVERRIDE[rel] || path.join(ROOT, rel);

let JSDOM;
try { ({ JSDOM } = await import('jsdom')); }
catch { console.log('★jsdomが入っていません。この検証は飛ばせません（SKIPを緑と呼ばない）'); process.exit(1); }

const GOLD = JSON.parse(fs.readFileSync(path.join(ROOT, 'tests/fixtures/cond-format-golden.json'), 'utf8'));

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ok   ' + n); } catch (e) { fail++; console.log('  NG   ' + n + '\n       ' + (e && e.message)); } };
const ok = (c, m) => { if (!c) throw new Error(m || 'expected truthy'); };
const eq = (a, b, m) => { if (a !== b) throw new Error((m ? m + ': ' : '') + '期待=' + JSON.stringify(b) + ' 実際=' + JSON.stringify(a)); };

const CANVAS_STUB = [
  '(function(){var noop=function(){};var 描いた=[];window.__描いた=描いた;',
  'var ctx=new Proxy({},{get:function(t,k){',
  ' if(k==="measureText")return function(){return{width:40};};',
  ' if(k==="canvas")return{width:900,height:600};',
  ' if(k==="getImageData")return function(){return{data:[]};};',
  ' if(k==="createLinearGradient"||k==="createPattern")return function(){return{addColorStop:noop};};',
  ' if(k==="fillStyle")return t.__fill;',
  ' if(k==="fillRect")return function(x,y,w,h){ 描いた.push({x:x,y:y,w:w,h:h,fill:t.__fill}); };',
  ' return noop;},set:function(t,k,v){ if(k==="fillStyle"){t.__fill=v;} t[k]=v; return true; }});',
  'HTMLCanvasElement.prototype.getContext=function(){return ctx;};})();',
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

console.log('');
console.log('[cond-format-ui] ★本物の画面で 条件付き書式を 実際に押す★');

const sh = () => win.sheets[win.activeSheet];
function 下地() {
  sh().data = {};
  if (sh().condFormats) sh().condFormats.length = 0;
  const 値 = [10, 20, 30, 40, 50];
  for (let i = 0; i < 値.length; i++) win.setCell(i, 0, String(値[i]));
  win.setCell(5, 0, 'あいう');          /* A6 = 文字 */
  win.sel(0, 0, 4, 0);                  /* A1:A5 を選ぶ */
}
const 窓 = () => doc.getElementById('cfOverlay');
const 入れる = (id, v) => { const el = doc.getElementById(id); el.value = v; };

T('★部品と 窓と 右クリックの項目が 画面に在る★', () => {
  ok(win.CondFormat && typeof win.CondFormat.当たるか === 'function', 'lib/cond-format.js が読まれていない');
  ok(typeof win.openCondFormat === 'function', '窓を開く物が無い');
  ok(typeof win.cfFormatAt === 'function', '描く時に使う物が無い');
  ok(窓(), '窓のHTMLが無い');
  const 項目 = [...doc.querySelectorAll('#ctx-menu .ctx-item')].map((e) => e.textContent);
  ok(項目.some((t) => t.indexOf('条件付き書式') >= 0), '★右クリックに 出ていない（置いた≠届く）★：' + 項目.join(' / '));
});

T('★窓を開くと 選んだ場所が出て、既定の色は 実Excelの値★', () => {
  下地();
  win.openCondFormat();
  eq(窓().style.display, 'flex', '窓が開いていない');
  eq(doc.getElementById('cfWhere').textContent, 'A1〜A5');
  eq(doc.getElementById('cfFill').value.toUpperCase(), '#FFC7CE');
  eq(doc.getElementById('cfColor').value.toUpperCase(), '#9C0006');
});

T('★押す前に「今 何個 当たるか」が出る（25より大きい＝3個）★', () => {
  下地(); win.openCondFormat();
  入れる('cfKind', 'セルの値'); win.cfKindChange();
  入れる('cfOp', 'より大きい'); 入れる('cfV1', '25');
  win.cfCount();
  const t = doc.getElementById('cfCount').textContent;
  ok(t.indexOf('3個') >= 0, '★数が出ていない／違う★：' + t);
  ok(t.indexOf('5マス中') >= 0, '母数が出ていない：' + t);
});

T('★押すと 色が付く（30/40/50 だけ）★', () => {
  下地(); win.openCondFormat();
  入れる('cfKind', 'セルの値'); win.cfKindChange();
  入れる('cfOp', 'より大きい'); 入れる('cfV1', '25');
  入れる('cfFill', '#ffc7ce');
  win.applyCondFormat();
  const 色 = (r) => { const f = win.cfFormatAt(r, 0, win.getCell(r, 0)); return f && f.塗り; };
  ok(!色(0), 'A1(10) に色が付いている');
  ok(!色(1), 'A2(20) に色が付いている');
  eq(String(色(2)).toLowerCase(), '#ffc7ce', 'A3(30) に色が付いていない');
  eq(String(色(3)).toLowerCase(), '#ffc7ce');
  eq(String(色(4)).toLowerCase(), '#ffc7ce');
});

T('★文字は どんな数より大きい（実Excelの真値どおり A6「あいう」も当たる）★', () => {
  下地();
  win.sel(0, 0, 5, 0);                    /* A1:A6（文字を含める） */
  win.openCondFormat();
  入れる('cfKind', 'セルの値'); win.cfKindChange();
  入れる('cfOp', 'より大きい'); 入れる('cfV1', '25');
  win.applyCondFormat();
  const f = win.cfFormatAt(5, 0, win.getCell(5, 0));
  ok(f && f.塗り, '★真値＝' + GOLD['①「25 より大きい」を A1:A9 に当てた']['文字『あいう』'] + ' なのに 当たっていない★');
});

T('★合計は 1円も動かない（見た目だけ）★', () => {
  下地();
  win.setCell(7, 0, '=SUM(A1:A5)');
  const 前 = String(win.getCell(7, 0).d);
  win.sel(0, 0, 4, 0);
  win.openCondFormat();
  入れる('cfKind', 'セルの値'); win.cfKindChange();
  入れる('cfOp', 'より大きい'); 入れる('cfV1', '25');
  win.applyCondFormat();
  eq(String(win.getCell(7, 0).d), 前, '★条件付き書式で 計算が変わった★');
  eq(前, '150', '下地の合計が違う');
});

T('★手で塗った色は 消えない（決まりを消すと 元の色に戻る）★', () => {
  下地();
  const cell = win.getCell(2, 0);
  cell.bgColor = '#FFFF00';               /* A3 を手で黄色 */
  win.sel(0, 0, 4, 0);
  win.openCondFormat();
  入れる('cfKind', 'セルの値'); win.cfKindChange();
  入れる('cfOp', 'より大きい'); 入れる('cfV1', '25'); 入れる('cfFill', '#ffc7ce');
  win.applyCondFormat();
  eq(String(win.cfFormatAt(2, 0, cell).塗り).toLowerCase(), '#ffc7ce', '条件付き書式が勝っていない');
  eq(win.getCell(2, 0).bgColor, '#FFFF00', '★手の塗りを 消している★');
  win.cfDeleteRule(0);
  eq(win.cfFormatAt(2, 0, cell), null, '消したのに 残っている');
  eq(win.getCell(2, 0).bgColor, '#FFFF00', '★元の色に戻っていない★');
});

T('★2つ当たったら 上に在る方が勝つ／上下で入れ替えられる★', () => {
  下地();
  win.sel(0, 0, 4, 0);
  win.openCondFormat();
  入れる('cfKind', 'セルの値'); win.cfKindChange();
  入れる('cfOp', 'より大きい'); 入れる('cfV1', '15'); 入れる('cfFill', '#ff0000');
  win.applyCondFormat();
  win.openCondFormat();
  入れる('cfKind', 'セルの値'); win.cfKindChange();
  入れる('cfOp', 'より大きい'); 入れる('cfV1', '35'); 入れる('cfFill', '#00ff00');
  win.applyCondFormat();
  /* ★新しい方が先頭＝強い★ */
  eq(String(win.cfFormatAt(3, 0, win.getCell(3, 0)).塗り).toLowerCase(), '#00ff00', 'A4(40) は 新しい方(>35)が勝つはず');
  eq(String(win.cfFormatAt(1, 0, win.getCell(1, 0)).塗り).toLowerCase(), '#ff0000', 'A2(20) は >15 だけが当たる');
  win.cfMoveRule(0, 1);                   /* >35 を下げる */
  eq(String(win.cfFormatAt(3, 0, win.getCell(3, 0)).塗り).toLowerCase(), '#ff0000', '★順番を変えても 見え方が変わらない★');
});

T('★一覧に 付けた決まりが 出る／消せる★', () => {
  下地();
  win.sel(0, 0, 4, 0);
  win.openCondFormat();
  入れる('cfKind', 'セルの値'); win.cfKindChange();
  入れる('cfOp', 'より大きい'); 入れる('cfV1', '25');
  win.applyCondFormat();
  const 字 = doc.getElementById('cfList').textContent;
  ok(字.indexOf('A1〜A5') >= 0, '場所が出ていない：' + 字);
  ok(字.indexOf('25') >= 0, '決め方が出ていない：' + 字);
  win.cfDeleteRule(0);
  ok(doc.getElementById('cfList').textContent.indexOf('まだ 1つも') >= 0, '消えていない');
});

T('★上位2項目＝40と50（真値どおり）★', () => {
  下地();
  win.sel(0, 0, 4, 0);
  win.openCondFormat();
  入れる('cfKind', '上位下位'); win.cfKindChange();
  入れる('cfOp', '上位'); 入れる('cfV1', '2');
  win.applyCondFormat();
  const 付いた = [0, 1, 2, 3, 4].filter((r) => win.cfFormatAt(r, 0, win.getCell(r, 0)));
  eq(付いた.join(','), '3,4', '★上位2個が 40と50 になっていない★');
});

T('★式のルールは 本物の計算エンジンで測る（=$A1>25 → 30/40/50 の行）★', () => {
  下地();
  win.sel(0, 0, 4, 0);
  win.openCondFormat();
  入れる('cfKind', '式'); win.cfKindChange();
  入れる('cfV1', '=$A1>25');
  win.applyCondFormat();
  const 付いた = [0, 1, 2, 3, 4].filter((r) => win.cfFormatAt(r, 0, win.getCell(r, 0)));
  eq(付いた.join(','), '2,3,4', '★式のルールが 効いていない（真値＝3行目から下）★');
});

T('★実際に描いた色に 条件付き書式の色が出る（描く所の配線）★', () => {
  /* ★jsdom には 幅も高さも無い★（wrapW=0）＝そのまま render すると 1つも描かない。
     だから ★大きさを入れてから 描く所を直接 回す★（実ブラウザでも 後で押して確かめる）。 */
  下地();
  win.wrapW = 900; win.wrapH = 600;
  win.__描いた.length = 0;
  win._renderPass();
  const 前 = win.__描いた.map((o) => String(o.fill).toLowerCase());
  ok(前.length > 0, '★大きさを入れても 1つも描いていない（この検査が空振りしている）★');
  ok(前.indexOf('#ffc7ce') < 0, '決まりが無いのに その色で描いている');
  win.sel(0, 0, 4, 0);
  win.openCondFormat();
  入れる('cfKind', 'セルの値'); win.cfKindChange();
  入れる('cfOp', 'より大きい'); 入れる('cfV1', '25'); 入れる('cfFill', '#ffc7ce');
  win.applyCondFormat();
  win.wrapW = 900; win.wrapH = 600;
  win.__描いた.length = 0;
  win._renderPass();
  const 後 = win.__描いた.map((o) => String(o.fill).toLowerCase());
  ok(後.indexOf('#ffc7ce') >= 0, '★描く時に 条件付き書式を見ていない（画面に色が出ない）★／描いた数=' + 後.length);
});

T('★色を付けても セルの持ち物(bgColor)を 1つも書き換えない★', () => {
  下地();
  win.sel(0, 0, 4, 0);
  const 前 = [0, 1, 2, 3, 4].map((r) => win.getCell(r, 0).bgColor);
  win.openCondFormat();
  入れる('cfKind', 'セルの値'); win.cfKindChange();
  入れる('cfOp', 'より大きい'); 入れる('cfV1', '25'); 入れる('cfFill', '#ffc7ce');
  win.applyCondFormat();
  const 後 = [0, 1, 2, 3, 4].map((r) => win.getCell(r, 0).bgColor);
  eq(後.join('|'), 前.join('|'), '★セルの塗りを 書き換えている（条件を外しても戻らなくなる）★');
});

T('★描く時に 落ちない（実際に render を1回 回す）★', () => {
  下地();
  win.sel(0, 0, 4, 0);
  win.openCondFormat();
  入れる('cfKind', 'セルの値'); win.cfKindChange();
  入れる('cfOp', 'より大きい'); 入れる('cfV1', '25');
  win.applyCondFormat();
  win.render();                            /* ★ここで落ちたら 客の画面が真っ白★ */
  ok(true);
});

console.log('');
console.log('  ' + pass + ' 緑 / ' + fail + ' 赤');

if (SELF) {
  const { spawnSync } = await import('node:child_process');
  const os = await import('node:os');
  const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'exally-cfui-'));
  console.log('');
  console.log('[self-test] わざと壊して 赤くなるかを数える（★repo は読むだけ★）');
  const BREAKS = [
    ['book.html', '★右クリックから 項目を消す★',
      (s) => s.replace('<div class="ctx-item" onclick="openCondFormat()">🎨 条件付き書式（当てはまったら 色を付ける）</div>', '')],
    ['book.html', '★描く時に 条件付き書式を見ない（塗り）★',
      (s) => s.replace('    var cf=cfFormatAt(r,c,cell);', '    var cf=null;')],
    ['book.html', '★手の塗りを 条件付き書式で 消してしまう★',
      (s) => s.replace('  _cfRules().unshift(ru);', '  _cfRules().unshift(ru); getCell(ru.範囲.r1,ru.範囲.c1).bgColor = ru.書式.塗り;')],
    ['book.html', '★新しい決まりを 一番弱い所に足す（Excelと逆）★',
      (s) => s.replace('  _cfRules().unshift(ru);', '  _cfRules().push(ru);')],
    ['book.html', '★押す前の「何個 当たるか」を 出さない★',
      (s) => s.replace("  document.getElementById('cfCount').textContent =", "  if(false) document.getElementById('cfCount').textContent =")],
    ['book.html', '★式のルールを 決め打ちで 当てる★',
      (s) => s.replace('    return v===true || v===1;', '    return true;')],
    ['book.html', '★消しても 残す★',
      (s) => s.replace('function cfDeleteRule(i){ _cfRules().splice(i,1);', 'function cfDeleteRule(i){')],
    ['lib/cond-format.js', '★文字を 数より小さい事にする★',
      (s) => s.replace('if (左が文字 && !右が文字) return 1;', 'if (左が文字 && !右が文字) return -1;')],
  ];
  let red = 0;
  for (const [rel, name, brk] of BREAKS) {
    const orig = fs.readFileSync(path.join(ROOT, rel), 'utf8');
    const bad = brk(orig);
    if (bad === orig) { console.log('  ★置換できず★  ' + name); continue; }
    const tmpFile = path.join(TMP, rel.replace(/[\\/]/g, '_'));
    fs.writeFileSync(tmpFile, bad, 'utf8');
    const env = Object.assign({}, process.env, { EXALLY_CFUI_OVERRIDE: JSON.stringify({ [rel]: tmpFile }) });
    const r = spawnSync(process.execPath, [path.join(__dirname, 'cond-format-ui.test.mjs')], { encoding: 'utf8', env });
    if (r.status !== 0) { red++; console.log('  赤くなった  ' + name); }
    else console.log('  ★素通り★  ' + name);
  }
  for (const rel of ['book.html', 'lib/cond-format.js']) {
    const now = fs.readFileSync(path.join(ROOT, rel), 'utf8');
    if (now.includes('var cf=null;') || now.includes('if(false) document.getElementById(\'cfCount\')')) {
      console.log('  ★NG★ ' + rel + ' に わざと壊した物が残っている'); process.exit(1);
    }
  }
  try { fs.rmSync(TMP, { recursive: true, force: true }); } catch (e) { /* 消せなくても検査は済んでいる */ }
  console.log('');
  console.log('  ' + red + '/' + BREAKS.length + ' 通りで赤くなった');
  process.exit(red === BREAKS.length ? 0 : 1);
}

process.exit(fail ? 1 : 0);
