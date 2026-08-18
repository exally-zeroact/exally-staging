/* diff-preview.test.mjs — ★直す前に必ず見せる（方針ver.6「絶対に守る3つ」の②）★
 *
 * なぜ必要か（2026-08-18 実物で測った）
 *   司さんの実物を本番と同じ経路で開き、計算!C10 を 5600→9999 に直したら
 *   ★3シート18本★ が書き換わった。★人は「1つ直した」つもり★。
 *   今までは保存を押すと その場で書き込んでいた＝何が変わるか見る場所が無かった。
 *
 * 見る所（指示役 2026-08-18 の裁定）
 *   ①★見せた件数と 実際に書き込んだ件数が同じか★（数える場所は1か所）
 *   ②★「変えない物」を 文言でなく 数で証明する★（触っていない部品は1バイトも変わらない）
 *   ③★[やめる] でファイルが1バイトも作られないか★（実際に押して確かめる）
 *   ④★件数は省略しない★（シートごとの件数は全部・中身だけ先頭3行）
 *   ⑤ 式は式で見せる／前が分からなければ「分かりません」と出す
 *
 * 使い方: node tests/diff-preview.test.mjs [--self-test]
 */
import { createRequire } from 'node:module';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const require = createRequire(import.meta.url);
const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const XLSX = require(path.join(ROOT, 'lib/xlsx.full.min.js'));
const ZipSurgeon = require(path.join(ROOT, 'lib/zip-surgeon.js'));
const XlsxEdit = require(path.join(ROOT, 'lib/xlsx-edit.js'));
const XlsbEdit = require(path.join(ROOT, 'lib/xlsb-edit.js'));
const TableRefs = require(path.join(ROOT, 'lib/table-refs.js'));
const DiffPreview = require(path.join(ROOT, 'lib/diff-preview.js'));

/* book-open.js はブラウザの物なので、必要な物だけ載せた入れ物で読み込む */
const box = { XLSX, ZipSurgeon, XlsxEdit, XlsbEdit, TableRefs };
new Function('self', 'window', fs.readFileSync(path.join(ROOT, 'js/book-open.js'), 'utf8')
  + '\n;self.__BookOpen = self.BookOpen;')(box, box);
const BookOpen = box.__BookOpen;

const FIX = path.join(ROOT, 'tests/fixtures/cross-sheet-sample.xlsb');
const orig = new Uint8Array(fs.readFileSync(FIX));
const fakeFile = (name, b) => ({ name, arrayBuffer: () => Promise.resolve(b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength)) });

let pass = 0, fail = 0;
const seenCounts = {};

/* ★受け取り口(js/book-open.js)を読んでいる画面を数える★
   テスト線の book.html は別系統で ★受け取り口が無い＝窓の置き場所が無い★。
   「無いから飛ばす」ではなく ★数えた物の数を出す★（見ていないだけの緑を作らない）。 */
const PAGES = fs.readdirSync(ROOT).filter(f => /\.html$/i.test(f));
const OPEN_PAGES = PAGES.filter(f => fs.readFileSync(path.join(ROOT, f), 'utf8').indexOf('js/book-open.js') >= 0);
const T = async (n, fn) => { try { await fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };

/* ══ self-test（判定そのものをわざと壊す） ══════════════════════════ */
if (process.argv.includes('--self-test')) {
  console.log('\n[diff-preview --self-test] わざと壊して、赤になる通り数を数える');
  const ways = [];
  const sheets = [{ name: 'S', data: { '0,0': { v: 9 }, '1,0': { v: 8 }, '2,0': { v: 7 }, '3,0': { v: 6 } } }];
  const ch = () => ({ '0,0': 9, '1,0': 8, '2,0': 7, '3,0': 6 });
  const base = { 'S|0,0': 1, 'S|1,0': 2, 'S|2,0': 3, 'S|3,0': 4 };
  const p = DiffPreview.build({ sheets, changedCells: ch, base, edited: { 'S|1,0': { beforeF: null } } });
  ways.push(['① 件数を省略しない（4件を4と数える）', p.total === 4 && p.sheets[0].count === 4]);
  ways.push(['② 中身だけ先頭3行に絞る', p.sheets[0].rows.length === 3 && p.sheets[0].more === 1]);
  ways.push(['③ 「あなたが直した所」を分ける', p.userCount === 1 && p.spreadCount === 3]);
  /* ★2026-08-18 変更: 人が直した行を先頭に出す（実物を撮って気づいた。
     シートの並び順のままだと 14シートの11枚目を直した人が 自分の1つを探す事になる） */
  ways.push(['④ ★人が直した行が先頭に来る★', p.sheets[0].rows[0].byUser === true]);
  ways.push(['⑤ 波及しただけの行に印を付けない', p.sheets[0].rows[1].byUser === false]);
  ways.push(['④b ★人が直したシートが一番上に来る★', (() => {
    const s5 = [{ name: 'A', data: {} }, { name: 'B', data: {} }];
    const q = DiffPreview.build({
      sheets: s5,
      changedCells: (sh) => (sh.name === 'A' ? { '0,0': 1 } : { '0,0': 2 }),
      base: {}, edited: { 'B|0,0': { beforeF: null } },
    });
    return q.sheets[0].name === 'B' && q.sheets[1].name === 'A';
  })()]);
  ways.push(['④c ★先頭3行に入らない所を直した時も そのシートを上に出す★', (() => {
    const many = {}; for (let i = 0; i < 8; i++) many[i + ',0'] = i;
    const s6 = [{ name: 'A', data: {} }, { name: 'B', data: {} }];
    const q = DiffPreview.build({
      sheets: s6,
      changedCells: (sh) => (sh.name === 'A' ? { '0,0': 1 } : many),
      base: {}, edited: { 'B|7,0': { beforeF: null } },
    });
    return q.sheets[0].name === 'B';
  })()]);
  ways.push(['⑥ ★前が控えに無ければ「分かりません」★', (() => {
    const q = DiffPreview.build({ sheets, changedCells: ch, base: {} });
    return q.sheets[0].rows[0].before === DiffPreview.UNKNOWN && q.sheets[0].rows[0].beforeUnknown === true;
  })()]);
  ways.push(['⑦ ★式は式で見せる（前の式 → 後の式）★', (() => {
    const s2 = [{ name: 'S', data: { '0,0': { f: '=B1+1' } } }];
    const q = DiffPreview.build({ sheets: s2, changedCells: () => ({ '0,0': 5 }), base: {}, edited: { 'S|0,0': { beforeF: '=B1' } } });
    const r = q.sheets[0].rows[0];
    return r.kind === 'formula' && r.before === '=B1' && r.after === '=B1+1';
  })()]);
  ways.push(['⑧ ★式で前を控えていなければ「分かりません」★', (() => {
    const s2 = [{ name: 'S', data: { '0,0': { f: '=B1+1' } } }];
    const q = DiffPreview.build({ sheets: s2, changedCells: () => ({ '0,0': 5 }), base: {} });
    return q.sheets[0].rows[0].before === DiffPreview.UNKNOWN;
  })()]);
  ways.push(['⑨ 変わる所が無ければ 0（窓を出さない側の判断材料）', DiffPreview.build({ sheets, changedCells: () => ({}), base }).total === 0]);
  ways.push(['⑩ 番地は行→列の順に並ぶ', (() => {
    const s3 = [{ name: 'S', data: {} }];
    const q = DiffPreview.build({ sheets: s3, changedCells: () => ({ '1,1': 1, '0,2': 2, '0,0': 3 }), base: {}, maxRows: 9 });
    return q.sheets[0].rows.map(r => r.addr).join(',') === 'A1,C1,B2';
  })()]);
  ways.push(['⑪ 書式を当てて読める字にする（961827.27… を出さない）', (() => {
    const s4 = [{ name: 'S', data: { '0,0': { numFmt: '#,##0' } } }];
    const q = DiffPreview.build({
      sheets: s4, changedCells: () => ({ '0,0': 961827.2727272727 }), base: { 'S|0,0': 1 },
      format: (v) => (typeof v === 'number' ? Math.round(v).toLocaleString('en-US') : String(v)),
    });
    return q.sheets[0].rows[0].after === '961,827';
  })()]);
  ways.push(['⑫ 1行目の言葉に 数が入る', /4か所/.test(DiffPreview.headline(p, 'a.xlsb'))]);
  let red = 0;
  for (const [label, okv] of ways) {
    if (okv) { red++; console.log('  ✓ ' + label); pass++; }
    else { console.log('  ✗ ' + label + ' → ★素通りした（見張りの穴）'); fail++; }
  }
  console.log('\n  ── 実測 ── 確かめた ' + ways.length + ' 通り / 通った ' + red + ' 通り');
  console.log('\n' + pass + ' passed, ' + fail + ' failed');
  process.exit(fail ? 1 : 0);
}

/* ══ 本番（実物の形で動かす） ══════════════════════════════════════ */
console.log('\n[diff-preview] 直す前に「何を書き込むか」を見せる');

const opened = await BookOpen.openFile(fakeFile('cross-sheet-sample.xlsb', orig));
const sheets = opened.sheets;
/* 開いた時の控えを「うちのエンジンから見た今のファイル」に取り直す
   （本番も loadSheetIntoEngine の中で rebaseSheet を呼んでいる） */
for (const sh of sheets) BookOpen.rebaseSheet(sh);

await T('★何も触っていなければ 見せる物は0か所（窓を出さない）', () => {
  const p = DiffPreview.build({ sheets, changedCells: BookOpen.changedCells, base: BookOpen.current().base });
  if (p.total !== 0) throw new Error('0でない: ' + p.total + ' / ' + JSON.stringify(p.sheets.map(s => s.name + ':' + s.count)));
});

/* 1セルだけ直す（4月!C4 = 1 → 99。まとめが参照している） */
const target = sheets.find(s => s.name === '4月');
target.data['3,2'] = Object.assign({}, target.data['3,2'], { v: 99, f: '', d: '99' });
const edited = { '4月|3,2': { beforeF: null } };

const plan = DiffPreview.build({
  sheets, changedCells: BookOpen.changedCells, base: BookOpen.current().base, edited,
});

await T('★1セル直すと 見せる物が出る（' + plan.total + 'か所）', () => {
  if (plan.total < 1) throw new Error('0か所のまま＝この検査が空振り');
});
await T('④ ★件数は省略しない（シートごとの件数が全部 出る）', () => {
  const sum = plan.sheets.reduce((a, s) => a + s.count, 0);
  if (sum !== plan.total) throw new Error('合計が合わない ' + sum + ' ≠ ' + plan.total);
  for (const s of plan.sheets) {
    if (s.rows.length > 3) throw new Error(s.name + ' の中身が3行を超えている');
    if (s.count > 3 && s.more !== s.count - 3) throw new Error(s.name + ' の「ほか」の数が合わない');
  }
});
await T('★「あなたが直した所」と「波及した所」を分ける', () => {
  if (plan.userCount !== 1) throw new Error('直した所が1でない: ' + plan.userCount);
  if (plan.userCount + plan.spreadCount !== plan.total) throw new Error('足すと合計にならない');
});

/* ── ① 見せた件数と 実際に書き込んだ件数が同じか ── */
const saved = await BookOpen.saveOpened(sheets);
const savedBytes = saved && saved.bytes ? saved.bytes : saved;

await T('① ★見せた件数と 実際に書き込んだ件数が同じ（数える場所は1か所）', () => {
  const a = XLSX.read(orig, { type: 'array', cellFormula: true });
  const b = XLSX.read(savedBytes, { type: 'array', cellFormula: true });
  let n = 0;
  const diffs = [];
  for (const nm of a.SheetNames) {
    const wa = a.Sheets[nm] || {}, wb2 = b.Sheets[nm] || {};
    const keys = new Set([...Object.keys(wa), ...Object.keys(wb2)].filter(k => k[0] !== '!'));
    for (const k of keys) {
      const va = wa[k] ? wa[k].v : undefined, vb = wb2[k] ? wb2[k].v : undefined;
      if (String(va) !== String(vb)) { n++; if (diffs.length < 6) diffs.push(nm + '!' + k + ' ' + va + '→' + vb); }
    }
  }
  console.log('      ── 実測 ── 見せた ' + plan.total + 'か所 ／ ファイルで実際に変わった ' + n + 'か所');
  if (n !== plan.total) throw new Error('食い違い（見せた ' + plan.total + ' / 実際 ' + n + '）' + diffs.join(' , '));
});

/* ── ② 変えない物を数で証明 ── */
await T('② ★触っていない部品は1バイトも変わらない（文言でなく数で示す）', async () => {
  const za = ZipSurgeon.read(orig), zb = ZipSurgeon.read(savedBytes);
  const na = za.names(), nb = zb.names();
  const same = [], diff = [], gone = [];
  for (const nm of na) {
    if (nb.indexOf(nm) < 0) { gone.push(nm); continue; }
    const a = await za.bytes(nm), b = await zb.bytes(nm);
    let eqv = a.length === b.length;
    if (eqv) for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) { eqv = false; break; }
    (eqv ? same : diff).push(nm);
  }
  console.log('      ── 実測 ── 部品 ' + na.length + '本 ／ 1バイトも変わらない ' + same.length
    + ' ／ 変わった ' + diff.length + '（' + diff.join(',') + '）'
    + ' ／ 外した ' + gone.length + '（' + gone.join(',') + '）');
  /* ★変わってよいのは 直したシートの記録と、索引を外すための2つだけ★
     索引(binaryIndex)は「何バイト目に何がある」の表。長さが変わると嘘になるので外す。 */
  const allowed = /worksheets\/sheet\d+\.bin$|^\[Content_Types\]\.xml$|worksheets\/_rels\/sheet\d+\.bin\.rels$/;
  const bad = diff.filter(n2 => !allowed.test(n2));
  if (bad.length) throw new Error('★触ってはいけない部品が変わった★: ' + bad.join(','));
  const goneBad = gone.filter(n2 => !/binaryIndex/.test(n2));
  if (goneBad.length) throw new Error('★消してはいけない部品が消えた★: ' + goneBad.join(','));
  if (same.length < 5) throw new Error('同じままの部品が少なすぎる＝この検査が空振り（' + same.length + '）');
});

/* ── ⑤ 配線 ── */
console.log('      ── 実測 ── 画面 ' + PAGES.length + '枚のうち 受け取り口(js/book-open.js)を持つのは ' + OPEN_PAGES.length + '枚: ' + (OPEN_PAGES.join(',') || '（無し）'));
if (!OPEN_PAGES.length) {
  console.log('      ＝この repo の画面は受け取り口を持たない系統。★窓ごしの検査は対象が無い★（部品は上で測っている）');
}
if (OPEN_PAGES.length) await T('⑤ 配線（book.html が窓を出してから書き出す）', () => {
  const h = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
  const must = [
    ['lib/diff-preview.js', '窓の部品を読んでいる'],
    ['DiffPreview.build(', '窓の中身を作っている'],
    ['changedCells: BookOpen.changedCells', '★書き込む側と同じ物で数えている（数える口を2つにしない）'],
    ['_askBeforeWrite(plan', '★書き込む前に聞いている'],
    ['id="diffOverlay"', '窓が在る'],
  ];
  for (const [needle, why] of must) if (h.indexOf(needle) < 0) throw new Error(why + ' … 「' + needle + '」が無い');
  if (/[^.\w]alert\s*\(/.test(h)) throw new Error('★alert を使っている（画面が止まる）');
  if (h.indexOf('done(false)') < 0) throw new Error('やめる の道が無い');
});

/* ── ③ [やめる] を★実際に押して★ファイルが1バイトも作られないか ── */
function waitFor(fn, why, ms) {
  const limit = ms || 4000, t0 = Date.now();
  return new Promise((res, rej) => {
    (function tick() {
      if (fn()) return res();
      if (Date.now() - t0 > limit) return rej(new Error(why));
      setTimeout(tick, 10);
    })();
  });
}
if (OPEN_PAGES.length) await T('③ ★[やめる]を実際に押す→1本も書き出さない／[書き出す]→1本', async () => {
  let JSDOM;
  try { ({ JSDOM } = require('jsdom')); }
  catch { throw new Error('jsdom が無い＝押せない。★緑ではない★'); }

  const html = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
  const dom = new JSDOM(html.replace(/<script[\s\S]*?<\/script>/g, ''), {
    runScripts: 'dangerously', pretendToBeVisual: true, url: 'http://localhost/',
  });
  const win = dom.window, doc = win.document;
  win.fetch = () => Promise.reject(new Error('no net'));
  win.scrollTo = () => { }; win.alert = () => { };
  win.DecompressionStream = globalThis.DecompressionStream;
  win.CompressionStream = globalThis.CompressionStream;
  win.Response = globalThis.Response; win.Blob = globalThis.Blob;
  const stub = new Proxy({}, { get: (t, k) => (k === 'measureText' ? (() => ({ width: 10 })) : k === 'canvas' ? { width: 800, height: 600 } : k === 'getImageData' ? (() => ({ data: [] })) : (() => { })) });
  win.HTMLCanvasElement.prototype.getContext = () => stub;
  const inject = (c) => { const el = doc.createElement('script'); el.textContent = c; doc.body.appendChild(el); };
  for (const m of html.matchAll(/<script src="([^"]+)"><\/script>/g)) {
    const src = m[1].split('?')[0];
    if (/^https?:/.test(src)) continue;
    inject(fs.readFileSync(path.join(ROOT, src), 'utf8'));
  }
  for (const m of html.matchAll(/<script(?![^>]*\ssrc=)[^>]*>([\s\S]*?)<\/script>/g)) inject(m[1]);
  await new Promise(r => { if (doc.readyState === 'complete') return r(); win.addEventListener('load', r); setTimeout(r, 3000); });
  for (const f of ['lib/xlsx.full.min.js', 'lib/xlsx-io.js', 'lib/zip-surgeon.js', 'lib/xlsx-edit.js',
    'lib/xlsb-edit.js', 'lib/table-refs.js', 'lib/diff-preview.js', 'js/book-open.js']) {
    inject(fs.readFileSync(path.join(ROOT, f), 'utf8'));
  }
  if (typeof win.saveOpenedBook !== 'function') throw new Error('saveOpenedBook が無い');

  /* ★渡し口を見張る（本当にファイルが作られたかは ここでしか分からない）★ */
  let delivered = 0, lastBytes = null;
  win.FileOut = { deliver: (buf) => { delivered++; lastBytes = buf; return Promise.resolve(); } };
  const toasts = [];
  win.showToast = (msg) => { toasts.push(String(msg)); };

  const res2 = await win.BookOpen.openFile(fakeFile('cross-sheet-sample.xlsb', orig));
  win.sheets = res2.sheets; win.activeSheet = 0; win._engineLoaded = {}; win._editedCells = {};
  win.initFormulaEngine(res2.sheets.map(s => s.name));
  for (let i = 0; i < res2.sheets.length; i++) win.loadSheetIntoEngine(i);
  /* 1セル打つ（本物の入口 setCell を通す＝人が直した所として覚えられる） */
  win.activeSheet = res2.sheets.findIndex(s => s.name === '4月');
  win.setCell(3, 2, '99');

  /* ── まず [やめる] ── */
  const p1 = win.saveOpenedBook();
  const ov = doc.getElementById('diffOverlay');
  await waitFor(() => ov.style.display === 'flex', '窓が出ない');
  const shownText = doc.getElementById('diffHead').textContent;
  const goLabel = doc.getElementById('diffGo').textContent;
  doc.getElementById('diffCancel').click();
  await p1;
  if (delivered !== 0) throw new Error('★やめるを押したのにファイルが作られた（' + delivered + '回）');

  /* ── 次に [書き込む] ── */
  const p2 = win.saveOpenedBook();
  await waitFor(() => ov.style.display === 'flex', '2回目に窓が出ない');
  doc.getElementById('diffGo').click();
  await p2;
  await waitFor(() => delivered === 1, '★書き込むを押したのにファイルが作られていない（' + delivered + '回）');
  if (!lastBytes || !lastBytes.length) throw new Error('作られたファイルが空');

  /* ★窓に出した数と、知らせに出た数が同じか（見せた数と書いた数を2通りに分けない）★ */
  const m1 = /(\d+)か所/.exec(shownText), m2 = /この(\d+)か所を直して 書き出す/.exec(goLabel);
  const m3 = /(\d+)か所 直して 書き出しました/.exec(toasts.join(' '));
  if (!m1 || !m2 || !m3) throw new Error('数が出ていない（窓=' + shownText + ' / ボタン=' + goLabel + ' / 知らせ=' + toasts.join(' | ') + '）');
  if (!(m1[1] === m2[1] && m2[1] === m3[1])) {
    throw new Error('★窓・ボタン・知らせで数が違う★ ' + m1[1] + ' / ' + m2[1] + ' / ' + m3[1]);
  }
  console.log('      ── 実測 ── 窓 ' + m1[1] + 'か所 ／ ボタン ' + m2[1] + 'か所 ／ 書き終わりの知らせ ' + m3[1] + 'か所（全部 同じ）');
  console.log('      ── 実測 ── やめるを押した時に作られたファイル 0本 ／ 書き込むを押した時 1本');
  try { win.close(); } catch (e) { /* 閉じられなくても検査は済んでいる */ }
});

/* ── ★開いていないシートに波及した値も ちゃんと書かれるか★ ──
   2026-08-18 実物で見つけた: 控え(base)は「まだ触っていないファイル」でなければならないのに、
   開いていないシートは ★保存の時に初めて流されて その場で控えを取り直す★ため、
   そこへ波及した値が「変わっていない」ことになり ★ファイルに書かれなかった★。
     実物: 計算!C10 を直して保存 → 全シートを見てから直すと18か所／
           ★計算しか見ずに直すと6か所（売上表!E1 が 1,298,210 のまま）★
   ＝ #ERROR は出ない。★合計が古いまま黙って保存される★（527,000→186,000 と同じ型）。
   直し方＝★最初の1直しの前に 全シートの控えを取る★。ここでそれを固定する。 */
/* ★2通りを両方 回す★（指示役 2026-08-18）
     「見た所だけ」… 直すシートしか開かない＝★普通の使い方★
     「全部見た」  … 先に全シートを開いてから直す
   2026-08-09 に直したはずの「合計が古いまま」が条件で残っていたのは、
   ★条件を変えたテストが無かった★から。だから条件そのものを2通り回す。 */
for (const mode of (OPEN_PAGES.length ? ['見た所だけ（普通の使い方）', '全部見た'] : [])) {
await T('★' + mode + 'でも 開いていないシートに波及した値が書かれる（合計が古いまま保存されない）', async () => {
  let JSDOM;
  try { ({ JSDOM } = require('jsdom')); } catch { throw new Error('jsdom が無い＝★緑ではない★'); }
  const html = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
  const dom = new JSDOM(html.replace(/<script[\s\S]*?<\/script>/g, ''), {
    runScripts: 'dangerously', pretendToBeVisual: true, url: 'http://localhost/',
  });
  const win = dom.window, doc = win.document;
  win.fetch = () => Promise.reject(new Error('no net'));
  win.scrollTo = () => { }; win.alert = () => { };
  win.DecompressionStream = globalThis.DecompressionStream;
  win.CompressionStream = globalThis.CompressionStream;
  win.Response = globalThis.Response; win.Blob = globalThis.Blob;
  const stub = new Proxy({}, { get: (t, k) => (k === 'measureText' ? (() => ({ width: 10 })) : k === 'canvas' ? { width: 800, height: 600 } : k === 'getImageData' ? (() => ({ data: [] })) : (() => { })) });
  win.HTMLCanvasElement.prototype.getContext = () => stub;
  const inject = (c) => { const el = doc.createElement('script'); el.textContent = c; doc.body.appendChild(el); };
  for (const m of html.matchAll(/<script src="([^"]+)"><\/script>/g)) {
    const src = m[1].split('?')[0];
    if (/^https?:/.test(src)) continue;
    inject(fs.readFileSync(path.join(ROOT, src), 'utf8'));
  }
  for (const m of html.matchAll(/<script(?![^>]*\ssrc=)[^>]*>([\s\S]*?)<\/script>/g)) inject(m[1]);
  await new Promise(r => { if (doc.readyState === 'complete') return r(); win.addEventListener('load', r); setTimeout(r, 3000); });
  for (const f of ['lib/xlsx.full.min.js', 'lib/xlsx-io.js', 'lib/zip-surgeon.js', 'lib/xlsx-edit.js',
    'lib/xlsb-edit.js', 'lib/table-refs.js', 'lib/diff-preview.js', 'js/book-open.js']) {
    inject(fs.readFileSync(path.join(ROOT, f), 'utf8'));
  }
  let out = null;
  win.FileOut = { deliver: (b) => { out = b; return Promise.resolve(); } };
  win.showToast = () => { };
  const res3 = await win.BookOpen.openFile(fakeFile('cross-sheet-sample.xlsb', orig));
  win.sheets = res3.sheets; win.activeSheet = 0; win._engineLoaded = {}; win._editedCells = {}; win._baselineTaken = false;
  win.initFormulaEngine(res3.sheets.map(s => s.name));
  const i4 = res3.sheets.findIndex(s => s.name === '4月');
  if (mode === '全部見た') { for (let i = 0; i < res3.sheets.length; i++) win.loadSheetIntoEngine(i); }
  else { win.loadSheetIntoEngine(i4); }   /* ★「4月」しか開かない＝普通の使い方★ */
  win.activeSheet = i4;
  win.setCell(3, 2, '99');
  const p3 = win.saveOpenedBook();
  const ov = doc.getElementById('diffOverlay');
  await waitFor(() => ov.style.display === 'flex', '窓が出ない');
  const head = doc.getElementById('diffHead').textContent;
  doc.getElementById('diffGo').click();
  await p3;
  await waitFor(() => !!out, 'ファイルが作られない');
  const wa = XLSX.read(orig, { type: 'array' }), wb3 = XLSX.read(out, { type: 'array' });
  const at = (wb, s, a) => { const w = wb.Sheets[s]; return w && w[a] ? w[a].v : undefined; };
  const sumName = res3.sheets.map(s => s.name).find(n => /まとめ/.test(n));
  const beforeSum = at(wa, sumName, 'B4'), afterSum = at(wb3, sumName, 'B4');
  console.log('      ── 実測 ── 窓「' + head + '」／ ' + sumName + '!B4 ' + beforeSum + ' → ' + afterSum);
  if (!/2つのシート/.test(head)) throw new Error('★開いていないシートの分が窓に出ていない★（' + head + '）');
  if (String(beforeSum) === String(afterSum)) {
    throw new Error('★開いていないシートの合計が古いまま書き出された★（' + sumName + '!B4 = ' + afterSum + '）');
  }
  /* ★2通りで同じ数になる事★（片方だけ少ないのが今回の穴だった） */
  const mHead = /（(\d+)か所/.exec(head);
  if (!mHead) throw new Error('窓に数が出ていない: ' + head);
  seenCounts[mode] = +mHead[1];
  try { win.close(); } catch (e) { /* 閉じられなくても検査は済んでいる */ }
});
}
if (OPEN_PAGES.length) await T('★「見た所だけ」と「全部見た」で 数が同じ（条件で変わらない）', () => {
  const a = seenCounts['見た所だけ（普通の使い方）'], b = seenCounts['全部見た'];
  console.log('      ── 実測 ── 見た所だけ ' + a + 'か所 ／ 全部見た ' + b + 'か所');
  if (a === undefined || b === undefined) throw new Error('片方が走っていない');
  if (a !== b) throw new Error('★条件で数が変わる★（見た所だけ ' + a + ' / 全部見た ' + b + '）');
});

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
