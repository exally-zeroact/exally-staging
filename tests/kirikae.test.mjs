/* kirikae.test.mjs — ★切り替え＝見る人・見る月を 変える（増えない）★
 * =============================================================================
 * ★司さんの言葉（2026-08-28）★「綺麗な形で切り替えれるん？」
 *   ＝1人分を出すたび シートが増えるのは ★切り出し★。★切り替え★は 増えない。
 *
 * ★指示役の検証要件（2026-08-28）★
 *   ・人を切り替えて 8人とも 出る／★シートが1枚も増えない★（数で）
 *   ・月を切り替えて 12か月とも 出る／★1日から／10日から／21日から の3通り★
 *   ・★「全員／全期間」に戻ると 元の表に 1セルも違わず 戻る★（★中身のSHAで突合★）
 *   ・★切り替えても 元の表は1セルも変わっていない★／★AIは 0回★
 *   ・★数が合う★＝★画面に描かれた字を1行ずつ足す★（★中の値で閉じない★）
 *
 * ★描かれた字★の作り方 … 画面の描く側と ★同じ関数(fmtForDisplay)★ を 本物の book.html から
 *   切り出して 通す。うちの切り替えが 隠した所は 描かない（rH/cW と 同じ判定）。
 *
 * 使い方: node tests/kirikae.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import crypto from 'node:crypto';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';
import { 注記を外す } from '../scripts/lib/chuki.mjs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const SELF = process.argv.includes('--self-test');
const OVERRIDE = process.env.EXALLY_KIRIKAE_OVERRIDE ? JSON.parse(process.env.EXALLY_KIRIKAE_OVERRIDE) : {};
const 読む = (rel) => fs.readFileSync(OVERRIDE[rel] || path.join(ROOT, rel), 'utf8');
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const Kirikae = require_(OVERRIDE['lib/kirikae.js'] || path.join(ROOT, 'lib/kirikae.js'));
const Recipe = require_(path.join(ROOT, 'lib/recipe.js'));
const GOLDEN = JSON.parse(fs.readFileSync(path.join(ROOT, 'tests/fixtures/shindan-golden.json'), 'utf8'));

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ok   ' + n); } catch (e) { fail++; console.log('  NG   ' + n + '\n       ' + (e && e.message)); } };
const ok = (c, m) => { if (!c) throw new Error(m || 'expected truthy'); };
const eq = (a, b, m) => { if (a !== b) throw new Error((m ? m + ': ' : '') + '期待=' + JSON.stringify(b) + ' 実際=' + JSON.stringify(a)); };
const SHA = (o) => crypto.createHash('sha256').update(JSON.stringify(o, Object.keys(o).sort())).digest('hex').slice(0, 16);

const book = 読む('book.html');

console.log('');
console.log('[kirikae] ★見る人・見る月を 変える（シートは 増えない）★');

/* ── 作り物の表（形は 実物と同じ：1行目=題／2行目=見出し／3行目〜=日付と金額）── */
const 表を作る = (人たち, 始まりの通し番号, 日数) => {
  const d = { '0,0': { v: 'ZERO代行　給料表' }, '1,0': { v: '日付' } };
  人たち.forEach((n, i) => { d['1,' + (i + 1)] = { v: n }; });
  for (let r = 0; r < 日数; r++) {
    /* ★画面と同じ形★＝式のセルは v が空で 計算した値は d（字） */
    d[(r + 2) + ',0'] = { v: '', f: '=A', d: String(始まりの通し番号 + r), numFmt: 'm/d;@' };
    人たち.forEach((n, i) => { d[(r + 2) + ',' + (i + 1)] = { v: '', f: '=X', d: String((i + 1) * 100 + r) }; });
  }
  return { name: '給料表', data: d };
};

/* ══ ①この表は 切り替えられるか ══ */
T('★人が横に並ぶ表なら 使える（見出しは2行目）★', () => {
  const 姿 = Kirikae.見る(表を作る(['白石正人', '長野孝'], 46023, 40));
  ok(姿.つかえる, '★使えないと 言っている：' + 姿.なぜ);
  eq(姿.見出しの行, 1, '★1行目を 見出しにしている★');
  eq(姿.日付の列, 0);
  eq(姿.人たち.map((x) => x.名).join(','), '白石正人,長野孝');
  eq(姿.月たち.join(','), '2026-01,2026-02');
});
T('★日付の列が 無い表では 使えない（出来ない物を 見せない）★', () => {
  const 姿 = Kirikae.見る({ name: 'あ', data: { '0,0': { v: '名前' }, '0,1': { v: '点' }, '1,0': { v: 'ア' }, '1,1': { v: 1 } } });
  ok(!姿.つかえる, '★使えると 言っている★');
  ok(姿.なぜ, '★理由が 無い★');
});
T('★人の列が 無い表では 使えない★', () => {
  const d = { '0,0': { v: '日付' } };
  for (let r = 1; r < 10; r++) d[r + ',0'] = { v: 46022 + r };
  ok(!Kirikae.見る({ name: 'あ', data: d }).つかえる);
});
T('★在る月だけ 出す（空振りの月を 選ばせない）★', () => {
  const 姿 = Kirikae.見る(表を作る(['白石正人'], 46023, 31));
  eq(姿.月たち.join(','), '2026-01');
});

/* ══ ②切り替える（隠すだけ・1セルも 書き換えない） ══ */
T('★人を切り替えると その人以外の列だけ 隠れる★', () => {
  const sh = 表を作る(['白石正人', '長野孝', '長野真道'], 46023, 40);
  const 姿 = Kirikae.見る(sh);
  const 出 = Kirikae.見る所を決める(sh, 姿, { 人: '長野孝' });
  eq(Object.keys(出.隠す列).sort().join(','), '1,3', '★日付の列や その人の列まで 隠している★');
  eq(Object.keys(出.隠す行).length, 0, '★人だけ選んだのに 行を 隠している★');
});
T('★月を切り替えると 期間の外の行だけ 隠れる★', () => {
  const sh = 表を作る(['白石正人'], 46023, 60);
  const 姿 = Kirikae.見る(sh);
  const 出 = Kirikae.見る所を決める(sh, 姿, { 月: '2026-01', 始まりの日: 1 });
  eq(出.見える行, 31);
  eq(Object.keys(出.隠す行).length, 29, '★2月の29日ぶんが 隠れる★');
  ok(!出.隠す行[0] && !出.隠す行[1], '★表の題と 見出しの行を 隠している★');
  eq(Object.keys(出.隠す列).length, 0, '★月だけ選んだのに 列を 隠している★');
});
T('★1日から／10日から／21日から の3通り★', () => {
  const sh = 表を作る(['白石正人'], 45900, 200);   /* 2025-09-13 から 200日 */
  const 姿 = Kirikae.見る(sh);
  const a = Kirikae.見る所を決める(sh, 姿, { 月: '2026-01', 始まりの日: 1 });
  const b = Kirikae.見る所を決める(sh, 姿, { 月: '2026-01', 始まりの日: 10 });
  const c = Kirikae.見る所を決める(sh, 姿, { 月: '2026-01', 始まりの日: 21 });
  eq(a.期間.from + '〜' + a.期間.to, '2026-01-01〜2026-01-31');
  eq(b.期間.from + '〜' + b.期間.to, '2025-12-10〜2026-01-09');
  eq(c.期間.from + '〜' + c.期間.to, '2025-12-21〜2026-01-20');
  eq(a.見える行, 31); eq(b.見える行, 31); eq(c.見える行, 31);
});
T('★今 何を見ているかを 1行 出す★', () => {
  const sh = 表を作る(['白石正人', '長野孝'], 46023, 60);
  const 姿 = Kirikae.見る(sh);
  eq(Kirikae.見る所を決める(sh, 姿, { 人: '白石正人', 月: '2026-01', 始まりの日: 1 }).いま,
    '白石正人／1月分（1/1〜1/31）／31行');
  eq(Kirikae.見る所を決める(sh, 姿, { 月: '2026-01', 始まりの日: 1 }).いま, '全員／1月分（1/1〜1/31）／31行');
  eq(Kirikae.戻す().いま, '全員／全期間');
});
T('★元の表は 1セルも 変わらない（切り替えは 見え方だけ）★', () => {
  const sh = 表を作る(['白石正人', '長野孝'], 46023, 60);
  const 姿 = Kirikae.見る(sh);
  const 前 = SHA(sh.data);
  Kirikae.見る所を決める(sh, 姿, { 人: '白石正人', 月: '2026-01', 始まりの日: 1 });
  Kirikae.見る所を決める(sh, 姿, { 人: '長野孝', 月: '2026-02', 始まりの日: 10 });
  Kirikae.戻す();
  eq(SHA(sh.data), 前, '★中身が 変わった★');
});
T('★戻すと 隠す物が 0になる★', () => {
  const 戻 = Kirikae.戻す();
  eq(Object.keys(戻.隠す行).length + Object.keys(戻.隠す列).length, 0);
});

/* ══ ③境界 ══ */
T('★同じ名前が 2人 居たら 断る（どちらか 分からない）★', () => {
  const sh = 表を作る(['白石正人', '白石正人'], 46023, 40);
  const 出 = Kirikae.見る所を決める(sh, Kirikae.見る(sh), { 人: '白石正人' });
  ok(出.なぜ && 出.なぜ.indexOf('2列') > 0, '★どちらかを 勝手に選んでいる★');
});
T('★居ない人は 断る★', () => {
  const sh = 表を作る(['白石正人'], 46023, 40);
  const 出 = Kirikae.見る所を決める(sh, Kirikae.見る(sh), { 人: '居ない人' });
  ok(出.なぜ && 出.なぜ.indexOf('見つかりません') > 0);
});
T('★期間に 1日も無い時は 断る（空の表を 見せない）★', () => {
  const sh = 表を作る(['白石正人'], 46023, 20);
  const 出 = Kirikae.見る所を決める(sh, Kirikae.見る(sh), { 月: '2026-06', 始まりの日: 1 });
  ok(出.なぜ && 出.なぜ.indexOf('1日も') > 0, '★空のまま 見せている★');
});
T('★日付でない行（合計の行など）は 隠さない★', () => {
  const sh = 表を作る(['白石正人'], 46023, 40);
  sh.data['42,0'] = { v: '' }; sh.data['42,1'] = { v: 999999 };
  const 出 = Kirikae.見る所を決める(sh, Kirikae.見る(sh), { 月: '2026-01', 始まりの日: 1 });
  ok(!出.隠す行[42], '★合計の行を 隠している★');
});
T('★月をまたぐ期間でも 途切れない★', () => {
  const sh = 表を作る(['白石正人'], 45900, 200);
  const 出 = Kirikae.見る所を決める(sh, Kirikae.見る(sh), { 月: '2026-01', 始まりの日: 21 });
  eq(出.見える行, 31);
});

/* ══ ④画面の作り（本物の字を読む） ══ */
T('★切り替えと 切り出しを 言葉で 分けている★', () => {
  const b = 注記を外す(book, { html: true });
  ok(b.indexOf('全員／全期間に戻す') > 0, '★戻す口が 無い★');
  ok(b.indexOf('この分を 新しいシートに 出す') > 0, '★切り出しの言葉が 無い★');
  ok(/id="kkHito"[^>]*onchange="切り替える\(\)"/.test(b), '★人▼が その場で 効かない★');
  ok(/id="kkTsuki"[^>]*onchange="切り替える\(\)"/.test(b), '★月▼が その場で 効かない★');
  ok(/id="kkShime"[^>]*onchange="切り替える\(\)"/.test(b), '★締めの決まりが その場で 効かない★');
});
T('★切り替えは 別の入れ物に 隠す（手で隠した行・絞り込みを 巻き込まない）★', () => {
  const b = 注記を外す(book, { html: true });
  ok(b.indexOf('sh.kirikaeRows && sh.kirikaeRows[r]') > 0, '★行の隠しが 効いていない★');
  ok(b.indexOf('sh.kirikaeCols && sh.kirikaeCols[c]') > 0, '★列の隠しが 効いていない★');
  ok(b.indexOf('sh.kirikaeRows = 出.隠す行; sh.kirikaeCols = 出.隠す列;') > 0, '★置く所が 無い★');
  ok(b.indexOf('hiddenRows[r]') > 0 && b.indexOf('filterHidden[r]') > 0, '★前からの物を 消している★');
});
T('★シートを作る所を 通らない（増えない）★', () => {
  const b = 注記を外す(book, { html: true });
  const i = b.indexOf('function 切り替える()');
  const 所 = b.slice(i, b.indexOf('function 切り替えを戻す'));
  for (const だめ of ['sheets.push(', 'addSheet(', 'setCell(']) {
    ok(所.indexOf(だめ) < 0, '★切り替えの中で ' + だめ + ' をしている★');
  }
});
T('★当てはまる表の時だけ 帯を出す★', () => {
  const b = 注記を外す(book, { html: true });
  ok(/<div id="kirikaeBar" hidden>/.test(b), '★最初から 出ている★');
  ok(/#kirikaeBar\[hidden\]\{ display:none !important; \}/.test(b), '★[hidden] の1行が 無い★');
  ok(b.indexOf('if(!_切り替えの姿){ 帯.hidden = true; return; }') > 0, '★使えない表でも 出している★');
});
T('★シートを変えたら 帯を 作り直す（前のシートの人が 並んだまま にしない）★', () => {
  const b = 注記を外す(book, { html: true });
  const i = b.indexOf('function switchSheet(idx)');
  ok(i > 0);
  ok(b.slice(i, i + 1400).indexOf('切り替えを整える()') > 0, '★作り直していない★');
});
T('★スマホで 勝手に拡大しない（選ぶ所は 16px）★', () => {
  const b = 注記を外す(book, { html: true });
  const i = b.indexOf('#kirikaeBar select');
  ok(i > 0, '★選ぶ所の 見た目が 無い★');
  ok(b.slice(i, i + 200).indexOf('font-size:16px') > 0, '★16px 未満（iPhoneが 勝手に拡大する）★');
});
T('★帯の字が 1文字ずつ 縦に割れない（前科3回）★', () => {
  const b = 注記を外す(book, { html: true });
  const i = b.indexOf('#kirikaeBar {');
  const 所 = b.slice(i, i + 400);
  ok(所.indexOf('white-space:nowrap') > 0, '★折り返して 縦に割れる★');
  ok(所.indexOf('overflow-x:auto') > 0, '★はみ出した時に 読めない★');
});
T('★スマホでは「今 何を見ているか」を 丸ごと 2行目に出す★', () => {
  /* ★390pxで 実測（2026-08-28）★＝1行に並べると 画面の外へ出て 横に動かさないと 読めなかった。 */
  const b = 注記を外す(book, { html: true });
  const i = b.indexOf('@media (max-width: 600px) {');
  ok(i > 0, '★スマホの決まりが 無い★');
  const 所 = b.slice(i, i + 400);
  ok(所.indexOf('#kirikaeBar { flex-wrap:wrap; }') > 0, '★次の行へ 回らない★');
  ok(/#kirikaeBar \.kk-now \{ order:9; flex-basis:100%;/.test(所), '★今 見ている物が 1行に ならない★');
});
T('★客に見せる字に ★ を書かない（切り替えの所）★', () => {
  const b = 注記を外す(book, { html: true });
  const i = b.indexOf('var _切り替えの姿 = null;');
  /* ★注記を外した後は 注記を 目印に出来ない★（-1 になって 全部を 拾っていた）
     ⇒ ★コードの目印★で 切る（2026-08-28 実際に 踏んだ）。 */
  const j = b.indexOf('function 覚えた手順を読む(', i);
  ok(j > i, '★切り替えの所の 終わりが 見つからない★');
  const 所 = b.slice(i, j);
  const 字 = (所.match(/'[^']*'/g) || []).join('');
  ok(字.indexOf('★') < 0, '★客の字に ★ が出ている★');
});

/* ══ ⑤実物（司さんの .xlsb）★描かれた字で 数える★ ══ */
const 実物 = GOLDEN.本.場所;
if (!fs.existsSync(実物)) {
  console.log('  ★未測定★ 実物が無い機械です（0件・異常なしにしない）');
} else {
  const XLSX = require_(path.join(ROOT, 'lib/xlsx.full.min.js'));
  const ZipSurgeon = require_(path.join(ROOT, 'lib/zip-surgeon.js'));
  const TableRefs = require_(path.join(ROOT, 'lib/table-refs.js'));
  const bytes = new Uint8Array(fs.readFileSync(実物));
  const wb = XLSX.read(bytes, { type: 'array', cellFormula: true });
  const rr = await TableRefs.resolve(bytes, 'xlsb', wb, ZipSurgeon);
  const fixes = (rr && rr.fixes) || {};
  const sheets = wb.SheetNames.map((name) => {
    const ws = wb.Sheets[name] || {}; const data = {};
    for (const k of Object.keys(ws)) {
      if (k[0] === '!') continue;
      const a = XLSX.utils.decode_cell(k); const c = ws[k];
      const fx = fixes[name + '|' + a.r + ',' + a.c];
      data[a.r + ',' + a.c] = { v: c.v, f: fx !== undefined ? fx : (c.f ? '=' + c.f : undefined), d: c.w, numFmt: c.z };
    }
    return { name, data };
  });
  const 給料表 = sheets.find((s) => s.name === '給料表');
  const 姿 = Kirikae.見る(給料表);

  /* ★描かれた字★＝隠していない行・列だけを 1行ずつ 字にする（中の値で 閉じない） */
  const 描く = (隠す行, 隠す列) => {
    const 行 = [];
    for (let r = 0; r <= 470; r++) {
      if (隠す行[r]) continue;
      const a = 給料表.data[r + ',0'];
      const 日 = a ? (a.d !== undefined && a.d !== '' ? a.d : a.v) : '';
      /* ★描かれる字は 機械によって 2通り★＝通し番号（46023）と 日付の字（1/1）。
         ★どちらも 日付の行として 数える★（描かれた物を 足すのが 目的）。 */
      if (!(/^\d{5}(\.\d+)?$/.test(String(日)) || /^\d{1,2}\/\d{1,2}/.test(String(日)))) continue;
      const 列 = [];
      for (let c = 0; c <= 姿.最大列; c++) {
        if (隠す列[c]) continue;
        const cell = 給料表.data[r + ',' + c];
        const v = cell ? (cell.d !== undefined && cell.d !== '' ? cell.d : cell.v) : '';
        列.push(String(v === undefined ? '' : v));
      }
      行.push(列);
    }
    return 行;
  };

  T('★実物で 8人とも 切り替えられる（シートは 1枚も 増えない）★', () => {
    const 前 = sheets.length;
    const 元SHA = SHA(給料表.data);
    const 出た = [];
    for (const 人 of 姿.人たち) {
      const 出 = Kirikae.見る所を決める(給料表, 姿, { 人: 人.名, 月: '2026-01', 始まりの日: 1 });
      ok(!出.なぜ, '★' + 人.名 + ' が 切り替えられない：' + 出.なぜ);
      const 描 = 描く(出.隠す行, 出.隠す列);
      eq(描.length, 31, '★' + 人.名 + ' の行数が 違う★');
      eq(描[0].length, 2, '★' + 人.名 + ' で 2列に なっていない★');
      const 合計 = 描.reduce((s, ln) => s + (Number(String(ln[1]).replace(/[,¥￥\s]/g, '')) || 0), 0);
      出た.push(人.名 + '=' + Math.round(合計).toLocaleString() + '円');
    }
    eq(出た.length, 8);
    eq(sheets.length, 前, '★シートが 増えた★');
    eq(SHA(給料表.data), 元SHA, '★元の表が 変わった★');
    console.log('       … 1月分（1/1〜1/31）: ' + 出た.join(' / '));
  });

  T('★描かれた字の合計が 切り出しと 一致する（2つの道で 同じ数）★', () => {
    const 出 = Kirikae.見る所を決める(給料表, 姿, { 人: '白石正人', 月: '2026-01', 始まりの日: 1 });
    const 描 = 描く(出.隠す行, 出.隠す列);
    const 切り替えの合計 = Math.round(描.reduce((s, ln) => s + (Number(String(ln[1]).replace(/[,¥￥\s]/g, '')) || 0), 0));
    const 切り出し = Recipe.切り出す(給料表, { 人: '白石正人', 月: '2026-01', 始まりの日: 1 });
    let 切り出しの合計 = 0;
    for (let i = 1; i <= 切り出し.何行; i++) 切り出しの合計 += Number(切り出し.data[i + ',1'].v) || 0;
    eq(描.length, 切り出し.何行, '★行数が 合わない★');
    /* ★1円までは 合う（丸めの差）★
       ＝描かれた字は ★1行ずつ 丸めた数★（10,638）。中の値は 丸めていない（10,637.5…）。
       ★同じ数になれと言うと 嘘になる★ので、★差は 行数（1行あたり1円）まで★と決める。
       実測 2026-08-28 … 描かれた字 137,550円／中の値 137,549円（31行・差1円）。 */
    const 差 = Math.abs(切り替えの合計 - Math.round(切り出しの合計));
    ok(差 <= 描.length, '★合計が 合わない（差 ' + 差 + '円・行数 ' + 描.length + '）★');
    console.log('       … 描かれた字 ' + 描.length + '行・' + 切り替えの合計.toLocaleString() + '円'
      + '／中の値 ' + Math.round(切り出しの合計).toLocaleString() + '円（差 ' + 差 + '円＝1行ずつ 丸めた分）');
  });

  T('★12か月とも 出る（1年ぶん 足すと 365日）★', () => {
    let 合計 = 0;
    for (let m = 1; m <= 12; m++) {
      const 出 = Kirikae.見る所を決める(給料表, 姿, { 人: '白石正人', 月: '2026-' + String(m).padStart(2, '0'), 始まりの日: 1 });
      ok(!出.なぜ, '★' + m + '月が 出ない：' + 出.なぜ);
      合計 += 描く(出.隠す行, 出.隠す列).length;
    }
    eq(合計, 365, '★1年ぶんに ならない（' + 合計 + '日）★');
  });

  T('★締めの決まり 3通り（1日から／10日から／21日から）★', () => {
    const 数 = ['1', '10', '21'].map((d) => {
      const 出 = Kirikae.見る所を決める(給料表, 姿, { 人: '白石正人', 月: '2026-03', 始まりの日: +d });
      return d + '日から=' + 描く(出.隠す行, 出.隠す列).length + '行(' + 出.期間.from.slice(5) + '〜' + 出.期間.to.slice(5) + ')';
    });
    eq(数[0], '1日から=31行(03-01〜03-31)');
    eq(数[1], '10日から=28行(02-10〜03-09)');
    eq(数[2], '21日から=28行(02-21〜03-20)');
    console.log('       … ' + 数.join(' / '));
  });

  T('★全員／全期間に戻すと 元の表に 1セルも違わず 戻る★', () => {
    const 元SHA = SHA(給料表.data);
    Kirikae.見る所を決める(給料表, 姿, { 人: '白石正人', 月: '2026-01', 始まりの日: 10 });
    const 戻 = Kirikae.戻す();
    eq(Object.keys(戻.隠す行).length + Object.keys(戻.隠す列).length, 0, '★隠れが 残っている★');
    eq(描く(戻.隠す行, 戻.隠す列).length, 365, '★全期間で 365日 出ない★');
    eq(SHA(給料表.data), 元SHA, '★中身のSHAが 違う★');
  });
}

console.log('');
console.log('  ' + pass + ' 緑 / ' + fail + ' 赤');

if (SELF) {
  const { spawnSync } = await import('node:child_process');
  const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'exally-kirikae-'));
  console.log('');
  console.log('[self-test] わざと壊して 赤くなるかを数える（★repo は読むだけ★）');
  const BREAKS = [
    ['lib/kirikae.js', '★人を切り替えても 何も隠さない★',
      (s) => s.replace('if (姿.人たち[i].名 !== 人の名) 隠す列[姿.人たち[i].列] = true;', '')],
    ['lib/kirikae.js', '★日付の列まで 隠す★',
      (s) => s.replace('if (c === 日付の列.列) continue;', '')],
    ['lib/kirikae.js', '★期間の外を 隠さない★',
      (s) => s.replace('if (期間 && !KI.期間の中か(ymd, 期間)) { 隠す行[r] = true; continue; }', '')],
    ['lib/kirikae.js', '★見出しの行から 隠し始める（何の表か 分からなくなる）★',
      (s) => s.replace('for (var r = 姿.見出しの行 + 1; r <= 姿.最大行; r++) {', 'for (var r = 0; r <= 姿.最大行; r++) {')],
    ['lib/kirikae.js', '★日付でない行（合計の行）も 隠す★',
      (s) => s.replace('if (!ymd) { continue; }', 'if (!ymd) { 隠す行[r] = true; continue; }')],
    ['lib/kirikae.js', '★同じ名前が2人でも 勝手に 選ぶ★',
      (s) => s.replace('if (当たり.length > 1) {', 'if (false) {')],
    ['lib/kirikae.js', '★1日も無い期間でも 見せる★',
      (s) => s.replace('if (期間 && !見える行) {', 'if (false) {')],
    ['lib/kirikae.js', '★今 何を見ているかを 言わない★',
      (s) => s.replace("いま: (人の名 || '全員') + '／' + (期間 ? 期間.言い方 : '全期間') + '／' + 見える行 + '行',", "いま: '',")],
    ['lib/kirikae.js', '★戻しても 隠れが 残る★',
      (s) => s.replace("return { 隠す行: {}, 隠す列: {}, いま: '全員／全期間' };", "return { 隠す行: { 3: true }, 隠す列: {}, いま: '全員／全期間' };")],
    ['lib/kirikae.js', '★日付の列が 無い表でも 使えると言う★',
      (s) => s.replace("if (日付の列.列 < 0 || !日付の列.数) return { つかえる: false, なぜ: '日付の列が ありません' };", '')],
    ['lib/kirikae.js', '★1行目を 見出しにする（来月 当たらなくなる）★',
      (s) => s.replace('var 見出しの行 = (み.行 >= 0 && み.数 > 0) ? み.行 : 0;', 'var 見出しの行 = 0;')],
    ['book.html', '★切り替えで シートを 増やす★',
      (s) => s.replace('  sh.kirikaeRows = 出.隠す行; sh.kirikaeCols = 出.隠す列;',
                       '  sheets.push({name:"ふえた",data:{},colW:{},rowH:{}});\n  sh.kirikaeRows = 出.隠す行; sh.kirikaeCols = 出.隠す列;')],
    ['book.html', '★隠しても 描く側に 効かない（行）★',
      (s) => s.replace('  if(sh.kirikaeRows && sh.kirikaeRows[r]) return 0;               /* ★切り替えで 隠れている行（別の入れ物）★ */', '')],
    ['book.html', '★隠しても 描く側に 効かない（列）★',
      (s) => s.replace('  if(sh.kirikaeCols && sh.kirikaeCols[c]) return 0;', '')],
    ['book.html', '★使えない表でも 帯を 出す★',
      (s) => s.replace('if(!_切り替えの姿){ 帯.hidden = true; return; }', 'if(!_切り替えの姿){ 帯.hidden = false; return; }')],
    ['book.html', '★最初から 帯を 出す★',
      (s) => s.replace('<div id="kirikaeBar" hidden>', '<div id="kirikaeBar">')],
    ['book.html', '★選ぶ所を 16px 未満にする（iPhoneが 勝手に拡大する）★',
      (s) => s.replace('font-family:inherit; font-size:16px;      /* ★iPhoneで 勝手に拡大しない＝16px★ */', 'font-family:inherit; font-size:13px;')],
    ['book.html', '★帯の字を 折り返させる（1文字ずつ 縦に割れる）★',
      (s) => s.replace('font-size:13px; color:#333333; white-space:nowrap; overflow-x:auto; -webkit-overflow-scrolling:touch;', 'font-size:13px; color:#333333;')],
    ['book.html', '★シートを変えても 帯を 作り直さない★',
      (s) => s.replace("  if(typeof 切り替えを整える === 'function') 切り替えを整える();\n", '')],
    ['book.html', '★戻す口を 消す★', (s) => s.replace('全員／全期間に戻す', '戻す')],
    ['book.html', '★スマホでも 1行に 詰め込む（今 見ている物が 画面の外）★',
      (s2) => s2.replace('  #kirikaeBar { flex-wrap:wrap; }', '')],
    ['book.html', '★今 見ている物を 2行目に 出さない★',
      (s2) => s2.replace('  #kirikaeBar .kk-now { order:9; flex-basis:100%; padding-top:2px; }', '')],
  ];
  let red = 0;
  for (const [rel, name, brk] of BREAKS) {
    const orig = fs.readFileSync(path.join(ROOT, rel), 'utf8');
    const bad = brk(orig);
    if (bad === orig) { console.log('  ★置換できず★  ' + name); continue; }
    const f = path.join(TMP, rel.replace(/[\\/]/g, '_'));
    fs.writeFileSync(f, bad, 'utf8');
    const env = Object.assign({}, process.env, { EXALLY_KIRIKAE_OVERRIDE: JSON.stringify({ [rel]: f }) });
    const r = spawnSync(process.execPath, ['--max-old-space-size=4096', path.join(__dirname, 'kirikae.test.mjs')], { encoding: 'utf8', env });
    if (r.status !== 0) { red++; console.log('  赤くなった  ' + name); }
    else console.log('  ★素通り★  ' + name);
  }
  try { fs.rmSync(TMP, { recursive: true, force: true }); } catch (e) { /* 消せなくても検査は済んでいる */ }
  console.log('');
  console.log('  ' + red + '/' + BREAKS.length + ' 通りで赤くなった');
  process.exit(red === BREAKS.length ? 0 : 1);
}

process.exit(fail ? 1 : 0);
