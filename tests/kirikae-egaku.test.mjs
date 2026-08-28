/* kirikae-egaku.test.mjs — ★切り替えた時に「画面に 描かれる所」に 在るか★
 * =============================================================================
 * ★2026-08-28 差し戻し★（司さん「形が変わっとる」／指示役が同じ物を見た）
 *   ★選ばれた行数（28行）は 合っていたのに 画面には 1行も 描かれていなかった★。
 *   犯人 … ★rowY が 隠れた行も 高さ22で 数えていた★
 *          ⇒ 2/21から始まる28行は y=770 に置かれ、画面（高さ590）の外だった。
 *          （1月分は 表の頭から始まるので たまたま 出ていた
 *            ＝★出る方でしか 数えていなかった★＝この試験が 無かった理由）
 *
 * ★この試験が 見る物★
 *   ・★選んだ行が 画面の中の 描かれる所に 並ぶか★（y が 上から 順に 22ずつ）
 *   ・★出る組（1月分・表の頭から）と 空だった組（3月分21日から・表の途中から）の 両方★
 *   ・★描かれた字を 1行ずつ 足す★（中の値で 閉じない）
 *   ★本物の book.html の 座標の関数を そのまま 切り出して 動かす★（写し取らない）
 *
 * 使い方: node tests/kirikae-egaku.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const SELF = process.argv.includes('--self-test');
const OVERRIDE = process.env.EXALLY_EGAKU_OVERRIDE ? JSON.parse(process.env.EXALLY_EGAKU_OVERRIDE) : {};
const 読む = (rel) => fs.readFileSync(OVERRIDE[rel] || path.join(ROOT, rel), 'utf8');
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const Kirikae = require_(path.join(ROOT, 'lib/kirikae.js'));

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ok   ' + n); } catch (e) { fail++; console.log('  NG   ' + n + '\n       ' + (e && e.message)); } };
const ok = (c, m) => { if (!c) throw new Error(m || 'expected truthy'); };
const eq = (a, b, m) => { if (a !== b) throw new Error((m ? m + ': ' : '') + '期待=' + JSON.stringify(b) + ' 実際=' + JSON.stringify(a)); };

const book = 読む('book.html');

console.log('');
console.log('[kirikae-egaku] ★切り替えた行が 画面に 描かれる所に 在るか★');

/* ── 本物の座標の関数を そのまま 切り出す ── */
const 切る = (頭, 尻, なに) => {
  const i = book.indexOf(頭), j = book.indexOf(尻, i + 1);
  if (i < 0 || j < 0) throw new Error('★' + なに + ' が 画面に 見つかりません★');
  return book.slice(i, j);
};
const 座標の所 = () => 切る('function cW(c){', '// ===== ユーティリティ =====', '座標の所');

/** 本物の関数を 動かす台（画面の大きさと シートを 渡すだけ） */
function 台(sheet, opt) {
  opt = opt || {};
  const 名 = ['sheets', 'activeSheet', 'scale', 'scrollTop', 'scrollLeft',
    'HDR_W', 'HDR_H', 'ROW_H', 'COL_W', 'ROWS', 'COLS', 'wrapW', 'wrapH'];
  const 値 = [[sheet], 0, 1, opt.scrollTop || 0, 0,
    46, 22, 22, 80, 1048576, 16384, opt.wrapW || 1280, opt.wrapH || 590];
  const f = new Function(...名, 座標の所()
    + ';return { rowY:rowY, colX:colX, rH:rH, cW:cW };');
  return f(...値);
}

/* ── 実物と同じ形の表（1行目=題／2行目=見出し／3行目〜=日付と金額）── */
const 表を作る = (人たち, 始まりの通し番号, 日数) => {
  const d = { '0,0': { v: 'ZERO代行　給料表' }, '1,0': { v: '日付' } };
  人たち.forEach((n, i) => { d['1,' + (i + 1)] = { v: n }; });
  for (let r = 0; r < 日数; r++) {
    d[(r + 2) + ',0'] = { v: '', f: '=A', d: String(始まりの通し番号 + r), numFmt: 'm/d;@' };
    人たち.forEach((n, i) => { d[(r + 2) + ',' + (i + 1)] = { v: '', f: '=X', d: String((i + 1) * 1000 + r) }; });
  }
  return { name: '給料表', data: d, colW: {}, rowH: {}, hiddenRows: {}, hiddenCols: {} };
};

/** ★画面の 描かれる所に 在る行★を 拾う（縦に流して 全部 見る） */
function 描かれる行(sheet, opt) {
  const HDR_H = 22, ROW_H = 22;
  const wrapH = (opt && opt.wrapH) || 590;
  const 出 = [];
  for (let 送り = 0; 送り < 60; 送り++) {
    /* ★送りは 行の高さの倍数★＝半端に送ると 境目の行が どちらの回でも 入らない
       （実測 2026-08-28：31行のうち 30行しか 拾えなかった＝★数え方の穴★） */
    const g = 台(sheet, { wrapH: wrapH, scrollTop: 送り * ROW_H * 20 });
    let 見た = 0;
    for (let r = 0; r < 600; r++) {
      if (g.rH(r) <= 0) continue;
      const y = g.rowY(r);
      if (y < HDR_H || y + ROW_H > wrapH) continue;
      見た++;
      if (!出.some((x) => x.行 === r)) 出.push({ 行: r, y: y });
    }
    if (送り > 2 && !見た) break;
  }
  return 出.sort((a, b) => a.行 - b.行);
}

/** ★描かれた字★（見えている列だけ・描く側と 同じ形で 字にする） */
function 描かれた字(sheet, 行たち) {
  const g = 台(sheet, {});
  /* ★表が 使っている列まで★（その先は 空っぽの列が ずっと続く＝Excelと同じ） */
  let 最大列 = 0;
  for (const k in sheet.data) { const c = +k.split(',')[1]; if (c > 最大列) 最大列 = c; }
  return 行たち.map((x) => {
    const 列 = [];
    for (let c = 0; c <= 最大列; c++) {
      if (g.cW(c) <= 0) continue;
      const cell = sheet.data[x.行 + ',' + c];
      const raw = cell ? ((cell.d !== undefined && cell.d !== '') ? cell.d : cell.v) : '';
      列.push(String(raw === undefined ? '' : raw));
    }
    return 列;
  });
}

/** ★日付の行だけ★（表の下の 空っぽの行は 数えない＝Excelでも 空行は ずっと続く） */
const 日付の行だけ = (sheet, 行たち) => 行たち.filter((x) => {
  const cell = sheet.data[x.行 + ',0'];
  if (!cell) return false;
  const v = (cell.v !== undefined && cell.v !== '') ? cell.v : cell.d;
  return Number(v) >= 20000;
});

const 当てる = (sheet, えらび) => {
  const 姿 = Kirikae.見る(sheet);
  const 出 = Kirikae.見る所を決める(sheet, 姿, えらび);
  if (出.なぜ) throw new Error('★切り替えられない：' + 出.なぜ);
  sheet.kirikaeRows = 出.隠す行; sheet.kirikaeCols = 出.隠す列;
  return 出;
};

/* ══ ★表の頭から始まる組（前から 出ていた方）★ ══ */
T('★1月分（表の頭から）＝31行が 画面の中に 並ぶ★', () => {
  const sh = 表を作る(['白石正人', '長野孝'], 46023, 365);
  当てる(sh, { 人: '白石正人', 月: '2026-01', 始まりの日: 1 });
  const 行 = 描かれる行(sh, {});
  const 日付の行 = 日付の行だけ(sh, 行);
  eq(日付の行.length, 31, '★画面に 描かれる所に 31行 無い★');
  eq(日付の行[0].行, 2, '★1/1 が 先頭でない★');
  eq(日付の行[0].y, 22 + 22 * 2, '★見出しの すぐ下に 無い★');
});

/* ══ ★表の途中から始まる組（★空になっていた方★）★ ══ */
T('★3月分・21日から（表の途中から）＝28行が 画面の中に 並ぶ★', () => {
  const sh = 表を作る(['白石正人', '長野孝'], 46023, 365);
  当てる(sh, { 人: '長野孝', 月: '2026-03', 始まりの日: 21 });
  const 行 = 描かれる行(sh, {});
  const 日付の行 = 日付の行だけ(sh, 行);
  /* ★ここが 前は 0行だった★（y=770 で 画面の外） */
  eq(日付の行.length, 28, '★画面に 描かれる所に 28行 無い（画面の外に 置かれている）★');
  eq(日付の行[0].行, 53, '★2/21 が 先頭でない★');
  eq(日付の行[0].y, 22 + 22 * 2, '★見出しの すぐ下に 無い（隠れた行が 場所を取っている）★');
});
T('★隠れた行は 場所を取らない（y が 22ずつ 詰まる）★', () => {
  const sh = 表を作る(['白石正人'], 46023, 365);
  当てる(sh, { 月: '2026-03', 始まりの日: 21 });
  const g = 台(sh, {});
  const 見える = [];
  for (let r = 0; r < 400; r++) if (g.rH(r) > 0) 見える.push(r);
  for (let i = 1; i < Math.min(30, 見える.length); i++) {
    eq(g.rowY(見える[i]) - g.rowY(見える[i - 1]), 22,
      '★' + (見える[i] + 1) + '行目の 前に 隙間が 空いている★');
  }
});
T('★行番号は 元のまま 飛び飛び（Excelと同じ）★', () => {
  const sh = 表を作る(['白石正人'], 46023, 365);
  当てる(sh, { 月: '2026-03', 始まりの日: 21 });
  const 行 = 日付の行だけ(sh, 描かれる行(sh, {})).map((x) => x.行 + 1);
  eq(行[0], 54, '★2/21 の行番号が 元の番号でない★');
  eq(行[行.length - 1], 81, '★3/20 の行番号が 元の番号でない★');
  ok(行[1] - 行[0] === 1, '★続きの日が 飛んでいる★');
});

/* ══ ★描かれた字を 1行ずつ 足す（両方の組で）★ ══ */
T('★描かれた字の合計＝出る組と 空だった組の 両方で 数える★', () => {
  const a = 表を作る(['白石正人'], 46023, 365);
  当てる(a, { 人: '白石正人', 月: '2026-01', 始まりの日: 1 });
  const 字a = 描かれた字(a, 日付の行だけ(a, 描かれる行(a, {})));
  eq(字a.length, 31);
  eq(字a[0].length, 2, '★日付と その人の 2列 でない★');
  const 合a = 字a.reduce((s, ln) => s + (Number(ln[1]) || 0), 0);
  eq(合a, 31000 + (0 + 30) * 31 / 2, '★1月分の 合計が 合わない★');

  const b = 表を作る(['白石正人'], 46023, 365);
  当てる(b, { 人: '白石正人', 月: '2026-03', 始まりの日: 21 });
  const 字b = 描かれた字(b, 日付の行だけ(b, 描かれる行(b, {})));
  eq(字b.length, 28, '★空だった組で 描かれた行が 28行 でない★');
  const 合b = 字b.reduce((s, ln) => s + (Number(ln[1]) || 0), 0);
  /* 2/21 は 51日目（0から数えて 51）＝1051 … 3/20 は 78日目＝1078 */
  eq(合b, (1051 + 1078) * 28 / 2, '★空だった組の 合計が 合わない★');
  console.log('       … 1月分 ' + 字a.length + '行/' + 合a + '　3月分21日から ' + 字b.length + '行/' + 合b);
});

/* ══ ★列も 場所を取らない★ ══ */
T('★隠れた列は 場所を取らない（選んだ人が すぐ隣に来る）★', () => {
  const sh = 表を作る(['白石正人', '長野孝', '長野真道'], 46023, 40);
  当てる(sh, { 人: '長野真道' });
  const g = 台(sh, {});
  eq(g.cW(1), 0, '★白石正人の列が 隠れていない★');
  eq(g.cW(2), 0, '★長野孝の列が 隠れていない★');
  eq(g.colX(3) - g.colX(0), 80, '★隠れた列が 場所を取っている★');
});

T('★選び直したら 座標も 数え直す（前の並びを 使い回さない）★', () => {
  /* ★同じ画面のまま 選び直す★＝ここで 数え直さないと
     ★前の月の隠しのまま 場所を計算する★（画面が ずれる／消える）。 */
  const sh = 表を作る(['白石正人'], 46023, 365);
  const g = 台(sh, {});                       /* ★台は 1つのまま★ */
  当てる(sh, { 月: '2026-01', 始まりの日: 1 });
  eq(g.rowY(2), 22 + 44, '★1月分の 先頭が 見出しの下に 無い★');
  当てる(sh, { 月: '2026-03', 始まりの日: 21 });
  eq(g.rH(2), 0, '★1月の行が まだ 見えている★');
  eq(g.rowY(53), 22 + 44, '★選び直したのに 前の並びで 場所を計算している★');
});

/* ══ ★戻すと 元どおり★ ══ */
T('★全員／全期間に戻すと 1行目から 順に 並ぶ★', () => {
  const sh = 表を作る(['白石正人', '長野孝'], 46023, 365);
  当てる(sh, { 人: '長野孝', 月: '2026-03', 始まりの日: 21 });
  const 戻 = Kirikae.戻す();
  sh.kirikaeRows = 戻.隠す行; sh.kirikaeCols = 戻.隠す列;
  const g = 台(sh, {});
  eq(g.rowY(0), 22, '★1行目が 上に無い★');
  eq(g.rowY(10), 22 + 220, '★詰まったまま★');
  eq(g.cW(1), 80, '★列が 隠れたまま★');
});

console.log('');
console.log('  ' + pass + ' 緑 / ' + fail + ' 赤');

if (SELF) {
  const { spawnSync } = await import('node:child_process');
  const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'exally-egaku-'));
  console.log('');
  console.log('[self-test] わざと壊して 赤くなるかを数える（★repo は読むだけ★）');
  const BREAKS = [
    ['★隠れた行も 場所を取る（差し戻された 元の形）★',
      (s) => s.replace('  var 隠 = _隠れた数(_隠れ行の並び(), r);\n  var y = HDR_H - scrollTop + (r - 隠) * ROW_H * scale;',
                       '  var y = HDR_H - scrollTop + r * ROW_H * scale;')],
    ['★隠れた列も 場所を取る★',
      (s) => s.replace('  var 隠 = _隠れた数(_隠れ列の並び(), c);\n  var x = HDR_W - scrollLeft + (c - 隠) * COL_W * scale;',
                       '  var x = HDR_W - scrollLeft + c * COL_W * scale;')],
    ['★切り替えの隠しを 座標が 見ない★',
      (s) => s.replace("return _隠しの並び('行', [sh.hiddenRows, sh.kirikaeRows, sh.filterHidden]);",
                       "return _隠しの並び('行', [sh.hiddenRows, sh.filterHidden]);")],
    ['★隠れた行の 高さを 0にしない★',
      (s) => s.replace('  if(sh.kirikaeRows && sh.kirikaeRows[r]) return 0;               /* ★切り替えで 隠れている行（別の入れ物）★ */', '')],
    ['★隠れた列の 幅を 0にしない★',
      (s) => s.replace('  if(sh.kirikaeCols && sh.kirikaeCols[c]) return 0;', '')],
    ['★隠す物が 変わっても 数え直さない（前の並びを 使い回す）★',
      (s) => s.replace('  if(箱.印 === 印) return 箱.並び;', '  if(箱.並び.length) return 箱.並び;')],
  ];
  let red = 0;
  for (const [name, brk] of BREAKS) {
    const orig = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
    const bad = brk(orig);
    if (bad === orig) { console.log('  ★置換できず★  ' + name); continue; }
    const f = path.join(TMP, 'book.html');
    fs.writeFileSync(f, bad, 'utf8');
    const env = Object.assign({}, process.env, { EXALLY_EGAKU_OVERRIDE: JSON.stringify({ 'book.html': f }) });
    const r = spawnSync(process.execPath, [path.join(__dirname, 'kirikae-egaku.test.mjs')], { encoding: 'utf8', env });
    if (r.status !== 0) { red++; console.log('  赤くなった  ' + name); }
    else console.log('  ★素通り★  ' + name);
  }
  try { fs.rmSync(TMP, { recursive: true, force: true }); } catch (e) { /* 消せなくても検査は済んでいる */ }
  console.log('');
  console.log('  ' + red + '/' + BREAKS.length + ' 通りで赤くなった');
  process.exit(red === BREAKS.length ? 0 : 1);
}

process.exit(fail ? 1 : 0);
