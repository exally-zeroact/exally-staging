/* shift-keeps-everything.test.mjs — ★行・列を 足しても 消しても 何も 置いていかれない★（押し0b・2026-11-07）
 *  ★なぜ★ 前は（6e0f6cf・jsdom）：
 *    B2:B4（3行）を 選んで 行の挿入＝★1行しか 入らない★
 *    条件付き書式の 範囲・絞り込みで 隠した 行・切り替え・ふりがな・グループ・許可範囲・固定枠・結合・図形・手書きは
 *    ★元の 場所に 置いていかれた★（ふりがなが 別の セルに 付く・結合が 別の 所に 残る）
 *  ★真値＝実Excel 16.0（scratchpad hakaru0b*.ps1 で 測った）★：
 *    足した 行・列は 上（左）を 継ぐ＝高さ・幅・段・入力の決まり・条件付き書式（真下でも 広がる）／隠す・メモは 継がない
 *    固定枠は 固定の 中に 足すと 伸びる・すぐ下は 変わらない・消しても 1 より 減らない
 *    結合は 中に 足すと 広がる・上を 消すと 縮む・1マスで 外れる／許可範囲は 丸ごと 消えても 残る（読めない範囲）
 *    図形は Placement=1＝上に 足すと 下へ・中に 足すと 伸びる・上の行を 消すと 縮む
 *  ★見張り★ book.html と js/*.js の 板の 鍵を 字で 数え、板の鍵の行き先 に 無い 鍵が 在れば 赤。
 *  使い方: node tests/shift-keeps-everything.test.mjs ／ --self-test＝直す前の 形に 戻すと 赤に なるか
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
let pass = 0, fail = 0;
const T = async (n, fn) => { try { await fn(); pass++; console.log('  ok   ' + n); } catch (e) { fail++; console.log('  NG   ' + n + '\n       ' + (e && e.message)); } };
const ok = (c, m) => { if (!c) throw new Error(m || 'expected truthy'); };
const eq = (a, b, m) => { const A = JSON.stringify(a), B = JSON.stringify(b); if (A !== B) throw new Error((m ? m + ': ' : '') + '期待=' + B + ' 実際=' + A); };

let JSDOM;
try { ({ JSDOM } = await import('jsdom')); }
catch { console.log('★jsdomが入っていません（SKIPを緑と呼ばない）。npm ci してください。'); process.exit(1); }

const SELF = process.argv.includes('--self-test');
let html = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
if (SELF) {
  /* ★直す前の 形に 戻す★（どれか 1つでも 当たらなければ 壊し方が 古い＝止める） */
  const 戻す = [
    ['if(入れる){ if(n >= 上) n += 幅; if(', 'if(入れる){ if(n >= 上) n += 1; if('],
    ['function _選びを揃える(){\n  var t;', 'function _選びを揃える(){ return;\n  var t;'],
    ['  if (出る) notify(', '  出る = false;\n  if (出る) notify('],
    ['if(nr>=ROWS) return;  /* 外へ', 'if(false) return;  /* 外へ'],
    ['    if (r === false) return r;', ''],
    ['var 幅=selR2-selR1+1, 写す=_座標の写しを作る(true, true);', 'var 幅=1, 写す=_座標の写しを作る(true, true);'],
    ['  _番号の棚を動かす(s, 縦, 入れる, 上, 下);\n', ''],
    ['  _四角の棚を動かす(s, 縦, 入れる, 上, 下);\n', ''],
    ['  if (入れる) _決まりを継ぐ(s, 縦, 上, 幅);\n', ''],
    ["'comments', 'ruby'];", "'comments'];"],
    ['  _結合を付け直す(s.data, 結合, true, true, selR1, selR2);\n', ''],
    ['  _浮かぶ物を動かす(s, true, 写す);\n  render(); hideCtxMenu();\n}\nfunction ctxInsertCol', '  render(); hideCtxMenu();\n}\nfunction ctxInsertCol'],
    ["  ruby: '動かす（番地）',\n", ''],
  ];
  for (const [a, b] of 戻す) {
    if (html.indexOf(a) < 0) { console.log('★--self-test の 壊し方が 当たらない★ ' + a.slice(0, 50)); process.exit(1); }
    html = html.replace(a, b);
  }
}

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

console.log('\n[shift-keeps-everything] 行・列を 動かしても 置いていかれない' + (SELF ? '（--self-test＝赤に なるはず）' : ''));
const 板 = () => win.sheets[win.activeSheet];
const 新しい本 = () => { win.eval("sheets.length=1; sheets[0]={name:'Sheet1',data:{},colW:{},rowH:{},hiddenRows:{},hiddenCols:{}}; activeSheet=0; undoStack.length=0; redoStack.length=0; scale=1; scrollTop=0; scrollLeft=0;"); };
const 上から = (r) => win.rowY(r) - win.rowY(0);
const 左から = (c) => win.colX(c) - win.colX(0);
const 鍵 = (o) => Object.keys(o || {}).sort();

/* ── 1 見張り：板の 鍵 全部に 行き先が 書いてあるか（字で 数える） ── */
await T('板の 鍵は 全部 板の鍵の行き先 に 書いてある（book.html と js/*.js を 字で 数える）', () => {
  const 行き先 = win.板の鍵の行き先;
  ok(行き先 && typeof 行き先 === 'object', '板の鍵の行き先 が 無い');
  const 字 = [html, ...fs.readdirSync(path.join(ROOT, 'js')).filter((f) => f.endsWith('.js')).map((f) => fs.readFileSync(path.join(ROOT, 'js', f), 'utf8'))].join('\n');
  const 見た = new Set();
  for (const m of 字.matchAll(/(?:sheets\[[^\]\n]{1,40}\]|(?:^|[^A-Za-z0-9_.$])(?:sh|_sh|s0))\.([A-Za-z_][A-Za-z0-9_]*)/g)) 見た.add(m[1]);
  ok(見た.size >= 15, '数え方が 壊れた（' + 見た.size + ' 個しか 見えない）');
  const 字の関数 = new Set(['length']);
  const 抜け = [...見た].filter((k) => !字の関数.has(k) && !Object.prototype.hasOwnProperty.call(行き先, k));
  eq(抜け.sort(), [], '行き先が 書いていない 鍵');
});

/* ── 2 3行 選んで 行の挿入＝全部の 棚が 3つ 下へ ── */
新しい本();
const 決まり = { kind: 'int', min: '1', max: '9' };
const 決まり2 = { kind: 'list', items: ['a', 'b'] };
win.setCell(0, 0, '見出し', true); win.setCell(4, 0, '5行目', true); win.setCell(5, 1, '結合', true);
{
  const s = 板();
  s.data['5,1'].mergeEnd = { r: 6, c: 2 };
  for (const k of ['5,2', '6,1', '6,2']) s.data[k] = Object.assign(s.data[k] || { v: '', f: '', d: '' }, { merged: { r: 5, c: 1 } });
  s.data['8,3'] = { v: '', f: '', d: 3, _溢れ元: '7,3', _台溢れ元: '7,3' };
  s.rowH = { 0: 30, 4: 40 };
  s.hiddenRows = { 4: true };
  s.filterHidden = { 5: true };
  s.kirikaeRows = { 6: true };
  s.outline = { 0: 2, 5: 2 };
  s.validations = { '0,4': 決まり, '4,4': 決まり2 };
  s.links = { '4,5': 'https://example.com/' };
  s.comments = { '4,6': { 字: 'メモ' } };
  s.sparklines = { '4,7': { 範囲: 'A1:A3' } };
  s.ruby = { '4,0': 'ごぎょう' };
  s.tables = [{ 名: '表1', r1: 4, c1: 0, r2: 6, c2: 2 }];
  s.condFormats = [{ 範囲: { r1: 0, c1: 1, r2: 0, c2: 1 } }, { 範囲: { r1: 4, c1: 1, r2: 5, c2: 1 } }];
  s.allowEdit = [{ 名: 'はんい1', 範囲: 'A5:A6', r1: 4, c1: 0, r2: 5, c2: 0 }];
  s.freezeRow = 2;
}
const 図形の上 = 上から(4), 図形の高さ = 上から(6) - 上から(4);
板().objects = [{ 種類: '四角', 名: 'Rectangle 1', x: 10, y: 図形の上, w: 50, h: 図形の高さ, z: 1 }];
板().ink = [{ 色: '#000000', 太さ: 1.32, 透け: 1, 点: [{ x: 5, y: 図形の上 }, { x: 6, y: 1 }], 名: 'ペン（黒）' }];
const 前の板 = JSON.stringify(板());
win.sel(1, 1, 3, 1); win.ctxInsertRow();
const s = 板();
await T('B2:B4 を 選んで 行の挿入＝3行 入る（A5→A8）', () => { eq(s.data['7,0'] && s.data['7,0'].v, '5行目', 'A8'); eq(s.data['4,0'], undefined, 'A5'); eq(s.data['0,0'].v, '見出し', 'A1'); });
await T('高さ＝足した 行は 上の 高さを 継ぐ・元の 40 は 3つ 下', () => { eq(s.rowH, { 0: 30, 1: 30, 2: 30, 3: 30, 7: 40 }); });
await T('手で 隠した 行は 3つ 下・足した 行は 隠さない', () => { eq(鍵(s.hiddenRows), ['7']); });
await T('絞り込みで 隠した 行・切り替えで 隠した 行も 3つ 下', () => { eq(鍵(s.filterHidden), ['8']); eq(鍵(s.kirikaeRows), ['9']); });
await T('グループの 段＝足した 行は 上の 段を 継ぐ・元の 段は 3つ 下', () => { eq(s.outline, { 0: 2, 1: 2, 2: 2, 3: 2, 8: 2 }); });
await T('入力の 決まり＝足した 行は 上を 継ぐ・元の 決まりは 3つ 下', () => { eq(鍵(s.validations), ['0,4', '1,4', '2,4', '3,4', '7,4']); eq(s.validations['2,4'], 決まり); eq(s.validations['7,4'], 決まり2); });
await T('リンク・付箋・スパークライン・ふりがな も 3つ 下', () => { eq(鍵(s.links), ['7,5']); eq(鍵(s.comments), ['7,6']); eq(鍵(s.sparklines), ['7,7']); eq(s.ruby, { '7,0': 'ごぎょう' }); });
await T('表は 3つ 下', () => { eq([s.tables[0].r1, s.tables[0].r2], [7, 9]); });
await T('条件付き書式＝真下に 足すと 広がる（B1→B1:B4）・下の 物は 3つ 下', () => { eq(s.condFormats.map((x) => [x.範囲.r1, x.範囲.r2]), [[0, 3], [7, 8]]); });
await T('許可範囲は 3つ 下（字も）', () => { eq([s.allowEdit[0].範囲, s.allowEdit[0].r1, s.allowEdit[0].r2], ['A8:A9', 7, 8]); });
await T('固定枠＝固定の 中に 足すと 伸びる（2→5）', () => { eq(s.freezeRow, 5); });
await T('結合 B6:C7 → B9:C10（印も 付け直す・元の 場所に 残さない）', () => {
  eq(s.data['8,1'].mergeEnd, { r: 9, c: 2 }, '元');
  for (const k of ['8,2', '9,1', '9,2']) eq(s.data[k].merged, { r: 8, c: 1 }, k);
  const 残り = Object.keys(s.data).filter((k) => s.data[k].mergeEnd || s.data[k].merged).sort();
  eq(残り, ['8,1', '8,2', '9,1', '9,2']);
});
await T('溢れの 印（元の 場所）も 3つ 下', () => { eq(s.data['11,3']._溢れ元, '10,3'); eq(s.data['11,3']._台溢れ元, '10,3'); });
await T('図形は 行と 一緒に 下へ（大きさは そのまま）', () => { eq(s.objects[0].y, 上から(7)); eq(s.objects[0].h, 図形の高さ); eq(s.objects[0].x, 10); });
await T('手書きの 線も 下へ（上の 点は そのまま）', () => { eq(s.ink[0].点[0].y, 上から(7)); eq(s.ink[0].点[1].y, 1); });
await T('Ctrl+Z で 全部 元へ', () => { win.doUndo(); eq(JSON.stringify(板()), 前の板); });

/* ── 3 消す側（実Excel の 測り） ── */
await T('3行 選んで 行の削除＝3行 消える（全部の 棚が 3つ 上へ）', () => {
  win.doRedo(); win.sel(1, 0, 3, 0); win.ctxDeleteRow();
  eq(JSON.stringify(板().ruby), JSON.stringify({ '4,0': 'ごぎょう' }));
  eq(板().data['4,0'].v, '5行目'); eq(板().freezeRow, 2); eq(板().data['5,1'].mergeEnd, { r: 6, c: 2 });
  eq(板().condFormats.map((x) => [x.範囲.r1, x.範囲.r2]), [[0, 0], [4, 5]]);
  eq(板().objects[0].y, 図形の上); eq(板().objects[0].h, 図形の高さ);
});
await T('結合の 上の 行を 消す＝縮む（B6:C7→B6:C6）・1マスに なったら 外れる（実測 E3:E4→E3）', () => {
  win.sel(5, 0, 5, 0); win.ctxDeleteRow();
  eq(板().data['5,1'].mergeEnd, { r: 5, c: 2 }, '横2マスは 残る');
  eq(Object.keys(板().data).filter((k) => 板().data[k].mergeEnd || 板().data[k].merged).sort(), ['5,1', '5,2']);
  win.sel(0, 1, 0, 1); win.ctxDeleteCol();
  eq(Object.keys(板().data).filter((k) => 板().data[k].mergeEnd || 板().data[k].merged), [], '1マスで 外れる');
});
await T('図形の 上の 行を 消す＝図形が 縮む（Placement=1）', () => {
  新しい本(); const 上 = 上から(4), 高 = 上から(6) - 上から(4);
  板().objects = [{ 種類: '四角', x: 0, y: 上, w: 10, h: 高 }];
  win.sel(4, 0, 4, 0); win.ctxDeleteRow();
  eq(板().objects[0].y, 上); eq(板().objects[0].h, 上から(5) - 上から(4));
});
await T('図形の 中に 行を 足す＝伸びる', () => {
  新しい本(); const 上 = 上から(4), 高 = 上から(6) - 上から(4);
  板().objects = [{ 種類: '四角', x: 0, y: 上, w: 10, h: 高 }];
  win.sel(5, 0, 5, 0); win.ctxInsertRow();
  eq(板().objects[0].y, 上); eq(板().objects[0].h, 上から(7) - 上から(4));
});
await T('固定枠＝すぐ下に 足しても 変わらない・固定の 行を 全部 消しても 1', () => {
  新しい本(); 板().freezeRow = 2;
  win.sel(2, 0, 2, 0); win.ctxInsertRow(); eq(板().freezeRow, 2, 'すぐ下');
  win.sel(0, 0, 4, 0); win.ctxDeleteRow(); eq(板().freezeRow, 1, '全部 消す');
});
await T('条件付き書式を 丸ごと 消す＝ルールも 消える／許可範囲は #REF! で 残る', () => {
  新しい本();
  板().condFormats = [{ 範囲: { r1: 2, c1: 1, r2: 4, c2: 1 } }];
  板().allowEdit = [{ 名: 'r1', 範囲: 'E3:E5', r1: 2, c1: 4, r2: 4, c2: 4 }];
  win.sel(2, 0, 4, 0); win.ctxDeleteRow();
  eq(板().condFormats.length, 0); eq(板().allowEdit.length, 1); eq(板().allowEdit[0].範囲, '#REF!');
});

/* ── 4 列 ── */
await T('2列 選んで 列の挿入＝2列 入り、幅は 左を 継ぎ、切り替え・固定枠・図形も 動く', () => {
  新しい本();
  win.setCell(0, 3, 'D1', true);
  板().colW = { 2: 150 }; 板().kirikaeCols = { 3: true }; 板().freezeCol = 3; 板().hiddenCols = { 2: true };
  板().validations = { '20,2': 決まり };
  const 左 = 左から(3);
  板().objects = [{ 種類: '四角', x: 左, y: 0, w: 30, h: 10 }];
  win.sel(0, 3, 0, 4); win.ctxInsertCol();
  eq(板().data['0,5'].v, 'D1'); eq(板().colW, { 2: 150, 3: 150, 4: 150 }); eq(鍵(板().kirikaeCols), ['5']);
  eq(板().freezeCol, 3, 'すぐ右は 変わらない'); eq(鍵(板().hiddenCols), ['2'], '隠すは 継がない');
  eq(鍵(板().validations), ['20,2', '20,3', '20,4']);
  eq(板().objects[0].x, 左から(5));
});

/* ── 5 逆向きの 選び・シートの 外へ 押し出す（本番前 taiketsu 0b の 止め1・2） ──
 *  上へ 引いて 選ぶと selR1>selR2 の まま 残る（onMM）＝幅が マイナスに なり 足すと 上書きで 消えていた
 *  列を 丸ごと 選んで 行の挿入＝幅 1,048,576＝中身が 全部 シートの 外へ 出て 画面が 空になっていた
 *  実Excel＝空でない セルを 外へ 押し出す 挿入は 断る（中身の 無い 書式だけの マスは 落とす） */
const 逆に選ぶ = (r1, c1, r2, c2) => win.eval('selR1=' + r1 + ';selC1=' + c1 + ';selR2=' + r2 + ';selC2=' + c2 + ';');
const 知らせ = [];
win.notify = (m) => { 知らせ.push(String(m)); };
await T('上へ 引いて 選んだ 3行に 行の挿入＝3行 入り、何も 消えない', () => {
  新しい本();
  for (let r = 0; r < 13; r++) win.setCell(r, 0, 'v' + r, true);
  逆に選ぶ(4, 0, 2, 0); win.ctxInsertRow();
  eq(Object.keys(板().data).length, 13, '13件の まま');
  eq(板().data['1,0'].v, 'v1'); eq(板().data['5,0'].v, 'v2'); eq(板().data['15,0'].v, 'v12');
});
await T('上へ 引いて 選んだ 3行を 消す＝その 3行が 消える', () => {
  新しい本();
  for (let r = 0; r < 8; r++) win.setCell(r, 0, 'v' + r, true);
  逆に選ぶ(4, 0, 2, 0); win.ctxDeleteRow();
  eq(Object.keys(板().data).length, 5); eq(板().data['2,0'].v, 'v5');
});
await T('左へ 引いて 選んだ 2列に 列の挿入・削除＝2列', () => {
  新しい本();
  for (let c = 0; c < 6; c++) win.setCell(0, c, 'c' + c, true);
  逆に選ぶ(0, 3, 0, 2); win.ctxInsertCol();
  eq(Object.keys(板().data).length, 6); eq(板().data['0,4'].v, 'c2');
  逆に選ぶ(0, 3, 0, 2); win.ctxDeleteCol();
  eq(板().data['0,2'].v, 'c2'); eq(Object.keys(板().data).length, 6);
});
await T('上へ 引いて 選んだ 行を 隠す＝その 行が 隠れる', () => {
  新しい本();
  逆に選ぶ(4, 0, 2, 0); win.ctxHideRow();
  eq(鍵(板().hiddenRows), ['2', '3', '4']);
});
await T('列を 丸ごと 選んで 行の挿入＝断って 知らせ、何も 変わらず 控えも 積まない', () => {
  新しい本(); 知らせ.length = 0;
  win.setCell(3, 1, 'x', true);
  win.sel(0, 1, win.ROWS - 1, 1);
  const 前 = JSON.stringify(板()), 控え = win.undoStack.length;
  eq(win.ctxInsertRow(), false);
  eq(JSON.stringify(板()), 前); eq(win.undoStack.length, 控え); ok(知らせ.length === 1 && /外/.test(知らせ[0]), '知らせ=' + 知らせ);
});
await T('一番下の 行に 中身が 在れば 1行 足すのも 断る／一番右の 列も 同じ', () => {
  新しい本(); 知らせ.length = 0;
  win.setCell(win.ROWS - 1, 0, 'end', true); win.setCell(0, win.COLS - 1, 'right', true);
  win.sel(0, 0, 0, 0);
  eq(win.ctxInsertRow(), false); eq(win.ctxInsertCol(), false);
  eq(板().data[(win.ROWS - 1) + ',0'].v, 'end'); eq(板().data['0,' + (win.COLS - 1)].v, 'right'); eq(知らせ.length, 2);
});
await T('一番下が 書式だけの マス（中身なし）なら 足せて、外へ 出る マスは 落ちる', () => {
  新しい本();
  win.setCell(0, 0, 'a', true);
  板().data[(win.ROWS - 1) + ',0'] = { v: '', f: '', d: '', s: { b: 1 } };
  win.sel(0, 0, 0, 0);
  ok(win.ctxInsertRow() !== false, '断られた');
  eq(板().data['1,0'].v, 'a');
  ok(Object.keys(板().data).every((k) => parseInt(k) < win.ROWS), 'シートの 外に 鍵が 残った');
});

/* セルの 挿入・削除の 窓も 同じ 揃えを 通る（叩き直し 止め2）／断ったら 「しました」で 上書きしない（止め1） */
for (const [名, 呼ぶ, 選ぶ, 見る] of [
  ['下方向に シフト', () => win.セルを挿入('down'), [4, 0, 2, 0], (d) => d['5,0'].v === 'v2' && d['1,0'].v === 'v1'],
  ['右方向に シフト', () => win.セルを挿入('right'), [0, 4, 0, 2], (d) => d['0,5'].v === 'v2' && d['0,1'].v === 'v1'],
  ['上方向に シフト（削除）', () => win.セルを削除('up'), [4, 0, 2, 0], (d) => d['2,0'].v === 'v5' && Object.keys(d).length === 5],
  ['左方向に シフト（削除）', () => win.セルを削除('left'), [0, 4, 0, 2], (d) => d['0,2'].v === 'v5' && Object.keys(d).length === 5],
]) {
  await T('逆向きに 選んで セルの ' + 名 + '＝何も 消えず 選んだ 数だけ 動く', () => {
    新しい本();
    const 縦 = 選ぶ[0] !== 0;
    for (let i = 0; i < 8; i++) 縦 ? win.setCell(i, 0, 'v' + i, true) : win.setCell(0, i, 'v' + i, true);
    逆に選ぶ(...選ぶ); 呼ぶ();
    const d = 板().data;
    ok(見る(d), JSON.stringify(Object.fromEntries(Object.entries(d).map(([k, x]) => [k, x.v]))));
    if (!/削除/.test(名)) eq(Object.keys(d).length, 8, '8件の まま');
  });
}
await T('窓で 「行全体」を 選んで 断られた＝知らせは 断りの 1件だけ（「挿入しました」で 消さない）', () => {
  新しい本(); 知らせ.length = 0;
  win.setCell(3, 1, 'x', true);
  win.sel(0, 1, win.ROWS - 1, 1);
  win.挿入の窓を開く();
  doc.querySelector('input[name="shiftPick"][value="row"]').checked = true;
  eq(win.詰める向きを決める(), false);
  eq(知らせ.length, 1, '知らせ=' + JSON.stringify(知らせ)); ok(/外/.test(知らせ[0]));
});

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (SELF) {
  const 効く = fail >= 15;
  console.log(効く ? '★--self-test：直す前の形で ' + fail + ' 本 赤＝見張りが効いている★' : '★--self-test：赤が ' + fail + ' 本しか出ない＝空振り★');
  process.exit(効く ? 0 : 1);
}
process.exit(fail ? 1 : 0);
