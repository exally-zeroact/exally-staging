/* engine-diff.mjs — ★計算の道具を 入れ替えた時に「答えが 変わっていないか」を 実物で 数える★
 *
 *  ★なぜ在るか（指示役 2026-08-29）★
 *    計算の道具（HyperFormula）を 2.6.1 → 3.4.0 に 丸ごと 入れ替えた（+67KB）。
 *    ★中身が 別物★なので ★答えが 変わり得る★。
 *    「表の参照 11,669本＝一致率100.0%」は ★前の道具で 出した数字★＝そのまま 使えない。
 *
 *  ★測り方（自分で 決めない・本番と 同じ物を 使う）★
 *    ・読み込みは 本番と 同じ … XLSX.read ＋ lib/table-refs.js（Table[列名]→A1）
 *    ・計算の設定も 本番と 同じ … buildEmpty({useArrayArithmetic:true, smartRounding:false})
 *      ★smartRounding は 切る★（book.html の 決まり。true に すると 1円 ずれる）
 *    ・2つの版で 同じ表を 計算して ★1セルずつ 突き合わせる★
 *    ・★比べる相手は 先に 複製する★（同じ物を 指したまま だと 自分の答えで 上書きされる）
 *
 *  ★★この道具を 作る時に 踏んだ 穴 4つ（★次の人へ★）★★
 *    ①直しの鍵は ★「シート名|行,列」★（A1形式では 1本も 当たらない）… Excel一致 2,043→10,523
 *    ②シートは ★先に 全部 addSheet してから setSheetContent★
 *      （1枚ずつ 作ると、まだ 無いシートを 見る式が その場で #ERROR! になり 後で 直らない）
 *    ③★HyperFormula は 日本語のシート名を そのままでは 読めない★
 *      実測 … `=歩合!B1` → #ERROR! ／ `='歩合'!B2` → 20
 *      ⇒ 本番と 同じ変換（convertFormula ＋ シート名を 引用符で 囲む）を 通す … 10,523→15,446
 *    ④★Exallyの関数プラグインを 積み忘れていた★（★これが 一番 効いた★）
 *      本番は exally-formula.js が TEXT/FILTER/MATCH/INDEX ほか 137本を 自前の物に 差し替える。
 *      積まないと 素の HyperFormula が 答える。実測 … `=TEXT(A5,"aaa")` が
 *      曜日「木」ではなく 字そのまま "aaa" ⇒ ★本番では 直っている 907本を「合わない」と 数えていた★
 *      ★登録の旗（_pluginRegistered）は 1本しか 無い★ので 2つ目の エンジンには 積めない
 *      ⇒ exally-formula.js を エンジンごとに 読み直す … 15,446→16,353
 *    ★数字が おかしい時は まず 自分の 測り方を 疑う★（うちで 7回目・指示役 2026-08-29）
 *
 *  ★★測って 分かった事（2026-08-29 実測・司さんの実物 19,323本）★★
 *    ・道具の入れ替え（2.6.1→3.4.0）で 答えが 変わった式 … ★0本★
 *    ・Excelが 残した答えと 合った … ★16,353本（84.6%）★
 *    ・合わない 2,970本の 中身 …
 *        ★2,918本＝空のセルを 指した時★（Excelは 0・うちは 空。exally-formula.js:44）
 *        ★52本＝.xlsb の 表の参照★＝★うちの計算では ない★。読み取りライブラリの この1行:
 *            case "PtgList": f.push("Table"+w[1].idx+"[#"+w[1].rt+"]")
 *          列の番号（w[1].c / w[1].C）を 持っているのに 使わないので 8列 全部が
 *          `Table1[#Data]` に なる。★.xlsx は <f> の 字を そのまま 使うので 起きない★
 *        ★数どうしで 違う＝0本★／★字で 違う＝0本★（＝お金の数字は 1本も 違わない）
 *
 *  ★出す物★
 *    ・式の本数／答えが 違った本数／どの関数を 使っている式が 違ったか
 *    ・★Excelが ファイルに 残した答え（キャッシュ）とも 比べる★＝どちらが 正しいかの 手掛かり
 *
 *  使い方:
 *    node scripts/engine-diff.mjs "<xlsxのパス>" [<前の道具のjs>] [<今の道具のjs>]
 *      前の道具を 省くと 今の道具どうし（＝差 0 の 空振り確認）に なる。
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const XLSX = require_(path.join(ROOT, 'lib/xlsx.full.min.js'));
const EF = require_(path.join(ROOT, 'exally-formula.js'));
const TableRefs = require_(path.join(ROOT, 'lib/table-refs.js'));
const ZipSurgeon = require_(path.join(ROOT, 'lib/zip-surgeon.js'));

/* ★印（--…）と 場所を 混ぜない★＝混ぜると 印を ファイル名として 読んで 落ちる（実際に 落ちた） */
const 引数 = process.argv.slice(2).filter((x) => !x.startsWith('--'));
const ファイル = 引数[0];
const 前の道具 = 引数[1] || path.join(ROOT, 'hyperformula.full.min.js');
const 今の道具 = 引数[2] || path.join(ROOT, 'hyperformula.full.min.js');
if (!ファイル) { console.log('★どのファイルを 測るか 教えてください★'); process.exit(1); }

/** ★シート名を 引用符で 囲む★（book.html の quoteSheetRefs と 同じ物を 写した）
 *  ★HyperFormula は 日本語のシート名を そのままでは 読めない★
 *    実測 2026-08-29 … `=歩合!B1` → #ERROR! ／ `='歩合'!B2` → 20
 *  ★本番は これを 通してから エンジンに 渡している★。ここを 通さずに 測ると
 *  「別シートを見る式が 8,800本 エラー」という ★嘘の数字★が 出る（実際に 出した）。 */
function シート名を囲む(f, 名前たち) {
  if (typeof f !== 'string' || f.charAt(0) !== '=') return f;
  for (const n of 名前たち) {
    if (/^[A-Za-z_][A-Za-z0-9_]*$/.test(n)) continue;
    const esc = n.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    f = f.replace(new RegExp("(^|[^'\\w])" + esc + "!", 'g'), "$1'" + n + "'!");
  }
  return f;
}

/** 本番と 同じ設定で 表を 載せて 全部 計算する */
function 計算する(道具, sheets) {
  const ns = require_(道具);
  const { HyperFormula } = ns;
  /* ★穴4（2026-08-29 実測）★＝★Exallyの関数（EX.TEXT など）を 積み忘れていた★。
     本番は exally-formula.js が TEXT/FILTER/MATCH/INDEX ほか 137本を 自前の物に 差し替える。
     積まないと 素の HyperFormula が 答えるので、例えば =TEXT(A5,"aaa") が
     曜日「木」ではなく 字そのまま "aaa" を 返し、★本番では 直っている物を
     「合わない」と 数えてしまう★（この積み忘れだけで 3,877本の うち 907本が 嘘だった）。
     ★登録の旗（_pluginRegistered）は 1本しか 無いので、2つ目の エンジンには 積めない★
     ⇒ exally-formula.js を エンジンごとに ★読み直して★ 別の旗を 持たせる。 */
  delete require_.cache[require_.resolve(path.join(ROOT, 'exally-formula.js'))];
  const EFn = require_(path.join(ROOT, 'exally-formula.js'));
  if (typeof EFn.registerExallyFunctions !== 'function' || !EFn.registerExallyFunctions(ns)) {
    throw new Error('★Exallyの関数プラグインを 積めなかった＝この測定は 本番の道では ない★');
  }
  /* ★本番と 同じ設定★（book.html の initFormulaEngine より） */
  const hf = HyperFormula.buildEmpty({ licenseKey: 'gpl-v3', useArrayArithmetic: true, smartRounding: false });
  /* ★シートは 先に 全部 作ってから 中身を 入れる★（本番も addSheetToEngine → loadSheetIntoEngine の順）。
     ★1枚ずつ 作って 入れると、まだ 無いシートを 見ている式が その場で #ERROR! になり
       後から シートが 出来ても 直らない★（2026-08-29 実測：別シート参照が 8,800本 エラーに なっていた）。 */
  for (const s of sheets) hf.addSheet(s.name);
  const 名前たち = sheets.map((x) => x.name);
  for (const s of sheets) {
    const g = s.grid.map((行) => 行.map((v) => (typeof v === 'string' && v.charAt(0) === '=') ? シート名を囲む(v, 名前たち) : v));
    hf.setSheetContent(hf.getSheetId(s.name), g);
  }
  const 出 = {};
  for (const s of sheets) {
    const sid = hf.getSheetId(s.name);
    for (const k of s.式の場所) {
      const [r, c] = k.split(',').map(Number);
      let v;
      try { v = hf.getCellValue({ sheet: sid, row: r, col: c }); }
      catch (e) { v = '★読めない★'; }
      let 答 = (v && typeof v === 'object' && v.value !== undefined) ? v.value : v;
      /* ★画面に出る形に そろえる★＝ここは ★全部 式のセル★（式の場所 しか 見ていない）。
         本番の _hfGetDisplay は 式のセルが 空セルを指したら ★0★ を出す（Excelと同じ・2026-08-29）。
         生の null のまま 比べると ★画面では 合っている物を「合わない」と 数える★
         ＝プラグイン積み忘れ（穴4）と 同じ型の 間違い。 */
      if (答 === null || 答 === undefined) 答 = 0;
      出[s.name + '!' + k] = 答;
    }
  }
  const 版 = HyperFormula.version;
  hf.destroy();
  return { 版, 答え: 出 };
}

/* ── 読み込み（本番と 同じ道）────────────────────────────── */
const bytes = new Uint8Array(fs.readFileSync(ファイル));
const 形 = /\.xlsb$/i.test(ファイル) ? 'xlsb' : (/\.xlsm$/i.test(ファイル) ? 'xlsm' : 'xlsx');
const wb = XLSX.read(bytes, { type: 'array', cellFormula: true, cellNF: true, sheetStubs: false });

let 直し = {}, 直しの数 = 0;
const r = await TableRefs.resolve(bytes, 形, wb, ZipSurgeon);
if (r && r.ok) { 直し = r.fixes || {}; 直しの数 = Object.keys(直し).length; }
else console.log('※ 表の名前での参照を 直せませんでした（' + (r && r.why) + '）');
/* ★直せた数だけ 見て 満足しない★＝table-refs.js は 辻褄が 合わない時 わざと 直さない
   （「壊すより断る」lib/table-refs.js:25）。★断った数を 数えないと 残った #ERROR の
   出どころが 分からない★（2026-08-29 指示役の宿題で 要った）。 */
if (r && r.stats) {
  console.log('  表の参照の 直し … 直した ' + (r.stats.fixed || 0).toLocaleString()
    + ' ／ ★わざと 断った ' + (r.stats.refused || 0).toLocaleString() + '★'
    + ' ／ 見送り ' + (r.stats.skipped || 0).toLocaleString()
    + ' ／ 見たセル ' + (r.stats.cells || 0).toLocaleString());
}

const sheets = [];
let 式の本数 = 0;
const Excelの答え = {};
const 生の式 = {};   /* ★合わない時に 元の式を 見るため（直す前と 直した後の 両方）★ */
for (const nm of wb.SheetNames) {
  const ws = wb.Sheets[nm];
  if (!ws) continue;
  const ref = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : null;
  if (!ref) continue;
  const grid = [];
  const 式の場所 = [];
  for (let R = 0; R <= ref.e.r; R++) {
    const 行 = [];
    for (let C = 0; C <= ref.e.c; C++) {
      const a = XLSX.utils.encode_cell({ r: R, c: C });
      const cell = ws[a];
      if (!cell) { 行.push(null); continue; }
      /* ★鍵の形は 本番と 同じに する★＝「シート名|行,列」（js/book-open.js と 同じ）。
         ★A1 の形で 引いていて 1本も 当たっていなかった★＝Table[列名] が 直らないまま 計算していた
         （2026-08-29 実測で 気づいた。★数字が おかしい時は まず 自分の道具を 疑う★） */
      const 直った = 直し[nm + '|' + R + ',' + C];
      if (cell.f || 直った) {
        const f = 直った || cell.f;
        生の式[nm + '!' + R + ',' + C] = { 元: String(cell.f || ''), 直: 直った ? String(直った) : null };
        /* ★本番と 同じ順で 通す★＝convertFormula（exally-formula.js）→ シート名を囲む */
        let 式 = '=' + String(f).replace(/^=/, '');
        if (EF && typeof EF.convertFormula === 'function') 式 = EF.convertFormula(式);
        行.push(式);
        式の場所.push(R + ',' + C);
        式の本数++;
        if (cell.v !== undefined && cell.v !== null) Excelの答え[nm + '!' + R + ',' + C] = cell.v;
      } else {
        行.push(cell.v === undefined ? null : cell.v);
      }
    }
    grid.push(行);
  }
  sheets.push({ name: nm, grid, 式の場所 });
}

console.log('');
console.log('[計算の道具の 突き合わせ] ★読むだけ★');
console.log('  ファイル … ' + path.basename(ファイル) + '（' + bytes.length.toLocaleString() + 'バイト）');
console.log('  シート ' + sheets.length + '枚 ／ ★式 ' + 式の本数.toLocaleString() + '本★'
  + ' ／ 表の名前での参照を 直した所 ' + 直しの数.toLocaleString() + '本');

/* ★比べる相手は 先に 複製する★（参照のままだと 自分の答えで 上書きされる） */
const 前 = 計算する(前の道具, sheets);
const 前の答え = JSON.parse(JSON.stringify(前.答え));
const 今 = 計算する(今の道具, sheets);

const 鍵 = Object.keys(前の答え);
let 同じ = 0;
const 違い = [];
for (const k of 鍵) {
  const a = 前の答え[k], b = 今.答え[k];
  if (String(a) === String(b)) { 同じ++; continue; }
  違い.push({ k, 前: a, 今: b, excel: Excelの答え[k] });
}
console.log('');
console.log('★前の道具 ' + 前.版 + ' ／ 今の道具 ' + 今.版 + '★');
console.log('  比べた式 … ' + 鍵.length.toLocaleString() + '本');
console.log('  ★答えが 同じ … ' + 同じ.toLocaleString() + '本（' + (鍵.length ? (Math.round(同じ / 鍵.length * 1000) / 10) : 0) + '%）★');
console.log('  ★答えが 違う … ' + 違い.length.toLocaleString() + '本★');

if (違い.length) {
  const 例 = 違い.slice(0, 20);
  console.log('');
  console.log('  違った所（先頭' + 例.length + '件・★どちらが正しいかは Excelの答えで 見る★）');
  for (const d of 例) {
    console.log('   ' + d.k + ' … 前=' + JSON.stringify(d.前) + ' 今=' + JSON.stringify(d.今)
      + (d.excel !== undefined ? ' ／ Excelが 残した答え=' + JSON.stringify(d.excel) : ' ／ Excelの答え=（無し）'));
  }
  if (違い.length > 例.length) console.log('   ほか ' + (違い.length - 例.length) + '本');
}

/* Excelが 残した答えとの 突き合わせ（★どちらの道具が 正しいか★） */
const 突合 = (答え) => {
  let 合 = 0, 差 = 0, 無 = 0;
  for (const k of Object.keys(Excelの答え)) {
    const e = Excelの答え[k], v = 答え[k];
    if (v === undefined) { 無++; continue; }
    const 数 = (typeof e === 'number' && typeof v === 'number');
    if (数 ? Math.abs(e - v) <= Math.max(1e-9, Math.abs(e) * 1e-9) : String(e) === String(v)) 合++;
    else 差++;
  }
  return { 合, 差, 無 };
};
/** ★合わない式を 関数ごとに 数える★（何が 効いているかを 見る） */
function 関数ごと(答え) {
  const 数 = {};
  for (const k of Object.keys(Excelの答え)) {
    const e = Excelの答え[k], v = 答え[k];
    if (v === undefined) continue;
    const 数か = (typeof e === 'number' && typeof v === 'number');
    if (数か ? Math.abs(e - v) <= Math.max(1e-9, Math.abs(e) * 1e-9) : String(e) === String(v)) continue;
    const [sh, rc] = k.split('!');
    const [R, C] = rc.split(',').map(Number);
    const 式 = ((sheets.find((x) => x.name === sh) || {}).grid || [])[R] || [];
    const f = String(式[C] || '');
    const 見た = new Set();
    const re = /([A-Z][A-Z0-9._]*)\s*\(/g;
    let m;
    while ((m = re.exec(f))) 見た.add(m[1]);
    if (!見た.size) 見た.add('（関数なし＝式だけ）');
    for (const n of 見た) 数[n] = (数[n] || 0) + 1;
  }
  return 数;
}

const a = 突合(前の答え), b = 突合(今.答え);
console.log('');
console.log('★Excelが ファイルに 残した答えとの 突き合わせ（' + Object.keys(Excelの答え).length.toLocaleString() + '本）★');
console.log('  前の道具 … 合う ' + a.合.toLocaleString() + ' ／ 違う ' + a.差.toLocaleString() + ' ／ 計算できず ' + a.無.toLocaleString());
if (process.argv.includes('--見本')) {
  let n = 0;
  for (const k of Object.keys(Excelの答え)) {
    const e = Excelの答え[k], v = 今.答え[k];
    const 数 = (typeof e === 'number' && typeof v === 'number');
    const 合う = 数 ? Math.abs(e - v) <= Math.max(1e-9, Math.abs(e) * 1e-9) : String(e) === String(v);
    if (合う || n >= 12) continue;
    /* --だけ=TEXT のように 群を 絞れる（★何が 効いているかを 1群ずつ 見る★） */
    const 絞り = (process.argv.find((x) => x.startsWith('--だけ=')) || '').split('=')[1];
    if (絞り) {
      const [sh2, rc2] = k.split('!');
      const [R2, C2] = rc2.split(',').map(Number);
      const f2 = String((((sheets.find((x) => x.name === sh2) || {}).grid || [])[R2] || [])[C2] || '');
      const 関数あり = /[A-Z][A-Z0-9._]*\s*\(/.test(f2);
      /* ★逃がしは 2つ 書く★＝1つだと `\b` が ★U+0008（バックスペースの 字）★に なり、
         `\(` も 素の `(` に なって ★new RegExp した その場で 落ちる★
         （2026-08-30 監査役が 見つけた＝この道は 1度も 通っていなかった）。 */
      if (絞り === 'なし' ? 関数あり : !new RegExp('\\b' + 絞り + '\\s*\\(').test(f2)) continue;
    }
    n++;
    const [sh, rc] = k.split('!');
    const [R, C] = rc.split(',').map(Number);
    const 式 = (sheets.find((x) => x.name === sh) || {}).grid?.[R]?.[C];
    console.log('   ' + k + ' 式=' + String(式).slice(0, 60) + ' ／ Excel=' + JSON.stringify(e) + ' ／ うち=' + JSON.stringify(v));
  }
}
console.log('  今の道具 … 合う ' + b.合.toLocaleString() + ' ／ 違う ' + b.差.toLocaleString() + ' ／ 計算できず ' + b.無.toLocaleString());
/* ★合わない物を「型」で 分ける★＝1つの決まりの違いで まとめて 説明できるか 見る */
{
  let 空ゼロ = 0, 数の差 = 0, 字の差 = 0, エラー = 0, その他 = 0;
  for (const k of Object.keys(Excelの答え)) {
    const e = Excelの答え[k], v = 今.答え[k];
    if (v === undefined) continue;
    const 数か = (typeof e === 'number' && typeof v === 'number');
    if (数か ? Math.abs(e - v) <= Math.max(1e-9, Math.abs(e) * 1e-9) : String(e) === String(v)) continue;
    if (e === 0 && (v === null || v === '' || v === undefined)) 空ゼロ++;
    else if (typeof v === 'string' && /^#/.test(v)) エラー++;
    else if (typeof e === 'number' && typeof v === 'number') 数の差++;
    else if (typeof e === 'string' || typeof v === 'string') 字の差++;
    else その他++;
  }
  console.log('');
  /* ★本番は 先に 独自層(_jsComputeFormula)を 通る★＝この測り台が 本番の道と 言えるのは
     「独自層が 1本も 答えていない」時だけ。★言い切る前に 数える★（指示役 2026-08-29）。
     独自層は ★式の 一番外側の 関数名★が この26個の時だけ 答える（exally-formula.js:921-929）。 */
  const 独自層の担当 = ('PERCENTILE QUARTILE XIRR DATESTRING OFFSET N CONVERT DSUM DAVERAGE DCOUNT '
    + 'DCOUNTA DMAX DMIN DPRODUCT DGET DSTDEV DSTDEVP DVAR DVARP LINEST BINOM FREQUENCY '
    + 'REDUCE SCAN MAP MAKEARRAY ISOMITTED').split(' ');
  let 独自層が答える = 0;
  const 独自層の内訳 = {};
  for (const s2 of sheets) {
    for (const 行 of s2.grid) {
      for (const c of (行 || [])) {
        if (typeof c !== 'string' || c[0] !== '=') continue;
        const m = c.slice(1).trim().match(/^([A-Z][A-Z0-9.]*)\s*\(/i);
        if (!m) continue;
        const 名 = m[1].toUpperCase().split('.')[0];
        if (独自層の担当.indexOf(名) >= 0) { 独自層が答える++; 独自層の内訳[名] = (独自層の内訳[名] || 0) + 1; }
      }
    }
  }
  console.log('');
  console.log('★この測り台は 本番の道か★');
  console.log('   独自層(_jsComputeFormula)が 答える式 … ' + 独自層が答える.toLocaleString() + '本'
    + (独自層が答える ? '（' + Object.keys(独自層の内訳).map((k) => k + ' ' + 独自層の内訳[k]).join(' / ') + '）'
      : ' ⇒ ★このブックでは 独自層は 出番なし＝HyperFormula だけで 本番と 同じ★'));
  console.log('');
  console.log('★合わない 3つ以上の 型で 分ける★');
  console.log('   ★Excelは 0・うちは 空★ … ' + 空ゼロ.toLocaleString() + '本'
    + '（Excelは 空のセルを 指すと 0 を返す。うち（HyperFormula）は 空のまま）');
  console.log('   うちが エラー … ' + エラー.toLocaleString() + '本');
  console.log('   数どうしで 違う … ' + 数の差.toLocaleString() + '本');
  console.log('   字が 絡んで 違う … ' + 字の差.toLocaleString() + '本');
  console.log('   その他 … ' + その他.toLocaleString() + '本');
  /* ★中身を 見ないと 直せない★＝エラーと 字の差を 少しだけ 出す */
  if (process.argv.includes('--中身')) {
    let e1 = 0, e2 = 0, e3 = 0;
    for (const k of Object.keys(Excelの答え)) {
      const e = Excelの答え[k], v = 今.答え[k];
      if (v === undefined) continue;
      const 数か = (typeof e === 'number' && typeof v === 'number');
      if (数か ? Math.abs(e - v) <= Math.max(1e-9, Math.abs(e) * 1e-9) : String(e) === String(v)) continue;
      const [sh, rc] = k.split('!');
      const [R, C] = rc.split(',').map(Number);
      const f = String((((sheets.find((x) => x.name === sh) || {}).grid || [])[R] || [])[C] || '');
      if (typeof v === 'string' && /^#/.test(v) && e1 < 6) {
        e1++;
        const g = 生の式[k] || {};
        console.log('   ［エラー］' + k + ' ／ Excel=' + JSON.stringify(e) + ' ／ うち=' + v);
        console.log('        元の式 … ' + String(g.元).slice(0, 160));
        if (g.直) console.log('        直した後 … ' + String(g.直).slice(0, 160));
      } else if (数か && e3 < 10) {
        e3++;
        const g3 = 生の式[k] || {};
        console.log('   ［数が違う］' + k + ' ／ Excel=' + e + ' ／ うち=' + v + ' ／ 差=' + (v - e));
        console.log('        元の式 … ' + String(g3.元).slice(0, 150));
        if (g3.直) console.log('        直した後 … ' + String(g3.直).slice(0, 150));
      } else if ((typeof e === 'string' || typeof v === 'string') && !/^#/.test(String(v)) && e2 < 6) {
        e2++; console.log('   ［字］' + k + ' ' + f.slice(0, 70) + ' ／ Excel=' + JSON.stringify(e) + ' ／ うち=' + JSON.stringify(v));
      }
    }
  }
}

const 内訳 = 関数ごと(今.答え);
const 並び = Object.keys(内訳).sort((x, y) => 内訳[y] - 内訳[x]).slice(0, 12);
if (並び.length) {
  console.log('');
  console.log('★合わない式の 関数ごとの 内訳（のべ・多い順）★');
  for (const n of 並び) console.log('   ' + n.padEnd(16) + ' ' + 内訳[n].toLocaleString() + '本');
}
console.log('');
