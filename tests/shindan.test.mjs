/* shindan.test.mjs — ★5 E2診断 1本目「消えた参照が IFERROR で隠れている」★
 *
 *  ★指示役 2026-08-25★
 *    ・客の画面には ★空しか出ない＝一生 気づけない★
 *    ・★競合が見つけられるのは 69個（Excel自身がエラーを出す物）だけ★
 *    ・司さんの実物で ★122本★
 *
 *  ★2通りで数えて 突き合わせる（自分の答えで閉じない）★
 *    ① lib/shindan.js（★括弧を数えて 1つ目の引数の中だけ見る★）
 *    ② 素朴なやり方（/IFERROR/ と /#REF!/ が同じ式に在るか）
 *    ⇒ ★実物では 同じ122★／★作り物では ②が 多く数える（包まれていない物まで拾う）★
 *
 *  使い方: node tests/shindan.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const SELF = process.argv.includes('--self-test');
const OVERRIDE = process.env.EXALLY_SHINDAN_OVERRIDE ? JSON.parse(process.env.EXALLY_SHINDAN_OVERRIDE) : {};
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const Shindan = require_(OVERRIDE['lib/shindan.js'] || path.join(ROOT, 'lib/shindan.js'));
const GOLDEN = JSON.parse(fs.readFileSync(path.join(ROOT, 'tests/fixtures/shindan-golden.json'), 'utf8'));

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ok   ' + n); } catch (e) { fail++; console.log('  NG   ' + n + '\n       ' + (e && e.message)); } };
const ok = (c, m) => { if (!c) throw new Error(m || 'expected truthy'); };
const eq = (a, b, m) => { if (a !== b) throw new Error((m ? m + ': ' : '') + '期待=' + JSON.stringify(b) + ' 実際=' + JSON.stringify(a)); };

const 本 = (式たち) => [{
  name: '計算',
  data: 式たち.reduce((a, f, i) => { a[i + ',0'] = { f: f, v: '' }; return a; }, {}),
}];

console.log('');
console.log('[shindan] ★消えた参照が IFERROR で隠れている★');

/* ══ 形を1つずつ ══ */
const 形 = [
  ['包まれている（IFERROR）', '=IFERROR(#REF!,"")', 1],
  ['包まれている（IFNA）', '=IFNA(VLOOKUP(A1,#REF!,2,0),"")', 1],
  ['包まれている（引き算の途中）', '=IFERROR((K1-Q1-#REF!)+C1,"")', 1],
  ['★2つ隠れている（のべ2・式は1本）★', '=IFERROR(#REF!+#REF!,"")', 2],
  ['入れ子（中のIFERRORが包んでいる）', '=IF(A1>0,IFERROR(#REF!,""),0)', 1],
  ['★包まれていない（同じ式に IFERROR は在る）★', '=IFERROR(A1,0)+INDEX(#REF!,1)', 0],
  ['★2つ目の引数の側（だめだった時に出す値）★', '=IFERROR(A1,#REF!)', 0],
  ['IFERROR が無い（Excelがエラーを出す＝競合でも見つかる）', '=INDEX(#REF!,1)', 0],
  ['★字の中の #REF!（式ではない）★', '=IFERROR("#REF!","")', 0],
  ['壊れていない', '=IFERROR(A1,"")', 0],
  ['大文字小文字を混ぜる', '=IfErRoR(#REF!,"")', 1],
  ['空白を挟む', '=IFERROR (#REF!,"")', 1],
];
for (const [名, f, 待つ] of 形) {
  T('★' + 名 + ' → ' + 待つ + '★', () => {
    const 数 = Shindan.隠している所(f).reduce((a, x) => a + x.数, 0);
    eq(数, 待つ, f);
  });
}
T('★式でない物（打った字）は 見ない★', () => {
  const r = Shindan.調べる([{ name: 'あ', data: { '0,0': { f: 'IFERROR(#REF!,"")', v: 'IFERROR(#REF!,"")' } } }]);
  eq(r.式の本数, 0, '★= で始まらない字を 式と読んでいる★');
});
T('★のべ と 式の本数 を 混ぜない★', () => {
  const r = Shindan.調べる(本(['=IFERROR(#REF!+#REF!,"")', '=IFERROR(#REF!,"")']));
  eq(r.のべ, 3);
  eq(r.式の本数, 2);
});
T('★場所（シート・セル）を出す（出さないと 見に行けない）★', () => {
  const r = Shindan.調べる(本(['=IFERROR(#REF!,"")']));
  eq(r.見つけた[0].シート, '計算');
  eq(r.見つけた[0].セル, 'A1');
});
T('★今 何が出ているかを 出す（空だと分かる）★', () => {
  const r = Shindan.調べる([{ name: 'あ', data: { '0,0': { f: '=IFERROR(#REF!,"")', v: '' } } }]);
  eq(r.見つけた[0].いま出ている物, '（空）');
});
T('★直す手掛かり（残っている参照）を 出す★', () => {
  const r = Shindan.調べる(本(['=IFERROR(INDEX(#REF!,MATCH(B1,計算!A67:A97,0)),"")']));
  eq(r.見つけた[0].手掛かり.join(','), '計算!A67:A97');
});
T('★列の字（27列目は AA）★', () => {
  eq(Shindan.列の字(0), 'A'); eq(Shindan.列の字(25), 'Z'); eq(Shindan.列の字(26), 'AA'); eq(Shindan.列の字(701), 'ZZ');
});

/* ══ 客に見せる言葉 ══ */
T('★客に見せる字に ★ を書かない★', () => {
  const w = Shindan.言葉('kakureta_ref', 122);
  ok(w, '言葉が無い');
  for (const k of ['題', '本文', 'つぎ']) ok(w[k].indexOf('★') < 0, '★客の字に ★ が出ている★：' + w[k]);
});
T('★「壊れています」と言わず、何が起きているかを言う★', () => {
  const w = Shindan.言葉('kakureta_ref', 122);
  ok(w.本文.indexOf('122か所') >= 0, '★数を 言っていない★');
  ok(w.本文.indexOf('空') >= 0, '★何が出ているかを 言っていない★');
  ok(w.本文.indexOf('合計') >= 0, '★何が起きるかを 言っていない★');
  ok(w.つぎ.length > 5, '★次に何をすればよいかを 言っていない★');
  for (const だめ of ['壊れて', 'エラーです', '異常']) ok(w.題.indexOf(だめ) < 0 && w.本文.indexOf(だめ) < 0, '★「' + だめ + '」と言っている★');
});

/* ══ 固まらない（小分け）══ */
T('★小分けで動く（画面を固めない）★', () => {
  const 式 = [];
  for (let i = 0; i < 5000; i++) 式.push(i % 50 === 0 ? '=IFERROR(#REF!,"")' : '=A1+1');
  const it = Shindan.調べる途中(本(式), { 一度に: 500 });
  let 回 = 0, r;
  do { r = it.next(); 回++; } while (!r.done && 回 < 100);
  ok(回 >= 5, '★1回で全部やっている（固まる）★：' + 回 + '回');
  eq(r.value.式の本数, 100);
});
T('★1回の小分けは 短い（実測）★', () => {
  const 式 = [];
  for (let i = 0; i < 20000; i++) 式.push('=IFERROR(#REF!,"")');
  const it = Shindan.調べる途中(本(式), { 一度に: 3000 });
  let 最長 = 0, r;
  do {
    const t = Date.now(); r = it.next(); 最長 = Math.max(最長, Date.now() - t);
  } while (!r.done);
  ok(最長 <= 200, '★1回が ' + 最長 + 'ミリ秒（200を超えると 固まって見える）★');
  console.log('       … 20,000本で 1回の最長 ' + 最長 + 'ミリ秒');
});

T('★「調べた時間」と「出るまでの時間」を 混ぜない★', () => {
  /* ★ブラウザで実測（2026-08-25）★ 仕事12ミリ秒 ／ 出るまで5.6秒＝★待っていた時間★。
     混ぜると「うちは遅い」という嘘にも「速い」という嘘にもなる。 */
  const 式 = [];
  for (let i = 0; i < 4000; i++) 式.push('=IFERROR(#REF!,"")');
  const it = Shindan.調べる途中(本(式), { 一度に: 1000 });
  let r;
  do { r = it.next(); } while (!r.done);
  ok(typeof r.value.かかった秒 === 'number', '★調べた時間を 出していない★');
  ok(typeof r.value.出るまでの秒 === 'number', '★出るまでの時間を 出していない★');
  ok(r.value.かかった秒 <= r.value.出るまでの秒 + 0.01, '★調べた時間が 出るまでの時間より 長い★');
});

/* ══ ★実物（司さんの .xlsb）★ ══ */
const 実物 = GOLDEN.本.場所;
if (!fs.existsSync(実物)) {
  console.log('  ★未測定★ 実物が無い機械です（0件・異常なしにしない）: ' + 実物);
} else {
  const XLSX = require_(path.join(ROOT, 'lib/xlsx.full.min.js'));
  const ZipSurgeon = require_(path.join(ROOT, 'lib/zip-surgeon.js'));
  const TableRefs = require_(path.join(ROOT, 'lib/table-refs.js'));
  const bytes = new Uint8Array(fs.readFileSync(実物));
  T('★実物のバイト数が ゴールデンと同じ（別のファイルを見ていない）★', () => {
    eq(bytes.length, GOLDEN.本.バイト);
  });
  const wb = XLSX.read(bytes, { type: 'array', cellFormula: true });
  const rr = await TableRefs.resolve(bytes, 'xlsb', wb, ZipSurgeon);
  const fixes = (rr && rr.fixes) || {};
  const sheets = wb.SheetNames.map((name) => {
    const ws = wb.Sheets[name] || {}; const data = {};
    for (const k of Object.keys(ws)) {
      if (k[0] === '!') continue;
      const c = ws[k]; const a = XLSX.utils.decode_cell(k); const key = name + '|' + a.r + ',' + a.c;
      data[a.r + ',' + a.c] = { v: c.v, f: fixes[key] !== undefined ? fixes[key] : (c.f ? '=' + c.f : undefined), t: c.t };
    }
    return { name, data };
  });
  const r = Shindan.調べる(sheets);
  T('★実物で 122本 見つける（ゴールデンと一致）★', () => {
    eq(r.式の本数, GOLDEN.本.式の本数);
    eq(r.のべ, GOLDEN.本.のべ);
  });
  T('★出たシートも 一致（散らばっていない）★', () => {
    const bys = {};
    for (const x of r.見つけた) bys[x.シート] = (bys[x.シート] || 0) + 1;
    eq(JSON.stringify(bys), JSON.stringify(GOLDEN.本.シート別));
  });
  T('★もう1通りの数え方でも 同じ（自分の答えで閉じない）★', () => {
    let 素朴 = 0;
    for (const sh of sheets) {
      for (const rc of Object.keys(sh.data)) {
        const f = sh.data[rc].f;
        if (f && f.charAt(0) === '=' && /IFERROR|IFNA/i.test(f) && /#REF!/.test(f)) 素朴++;
      }
    }
    eq(素朴, GOLDEN.本.素朴なやり方でも, '★2通りの数が 違う（どちらかが 嘘）★');
  });
  T('★競合が見つけられる数（Excel自身がエラー）とは 別物★', () => {
    let e = 0;
    for (const nm of wb.SheetNames) {
      const ws = wb.Sheets[nm] || {};
      for (const k of Object.keys(ws)) { if (k[0] !== '!' && ws[k].t === 'e') e++; }
    }
    eq(e, GOLDEN.本.Excel自身がエラーのセル);
    ok(r.式の本数 > e, '★診断の数が エラーの数以下＝新しく見つけた物が無い★');
  });
  T('★実際に 合計が黙って小さくなっている所を 数える（この診断の値打ち）★', () => {
    const 隠れ = r.見つけた.filter((x) => x.シート === '計算');
    const A1 = (s) => { const m = /^\$?([A-Z]{1,3})\$?([0-9]{1,7})$/.exec(s); if (!m) return null; let c = 0; for (const ch of m[1]) c = c * 26 + (ch.charCodeAt(0) - 64); return { r: +m[2] - 1, c: c - 1 }; };
    const 計算 = sheets.find((s) => s.name === '計算');
    const 出 = [];
    for (const rc of Object.keys(計算.data)) {
      const f = 計算.data[rc].f;
      if (!f || !/SUBTOTAL|SUM\(/i.test(f)) continue;
      const m = /(\$?[A-Z]{1,3}\$?[0-9]{1,7}):(\$?[A-Z]{1,3}\$?[0-9]{1,7})/.exec(f);
      if (!m) continue;
      const a = A1(m[1]), b = A1(m[2]);
      if (!a || !b) continue;
      const 何行 = 隠れ.filter((x) => x.r >= Math.min(a.r, b.r) && x.r <= Math.max(a.r, b.r) && x.c >= Math.min(a.c, b.c) && x.c <= Math.max(a.c, b.c)).length;
      if (何行) 出.push({ セル: Shindan.セルの名(rc), 何行: 何行, いまの値: 計算.data[rc].v });
    }
    eq(JSON.stringify(出), JSON.stringify(GOLDEN.本.黙って小さくなっている合計),
      '★実物の合計の壊れ方が 変わった（数え直すこと）★');
    console.log('       … ' + 出.map((x) => x.セル + '=' + JSON.stringify(x.いまの値) + '（' + x.何行 + '行が空）').join(' / '));
  });
  T('★速い（0円・AIを1回も呼ばない）★', () => {
    ok(r.かかった秒 <= GOLDEN.本.何秒以内, '★' + r.かかった秒 + '秒（' + GOLDEN.本.何秒以内 + '秒以内）★');
    console.log('       … 見たセル ' + r.見たセル.toLocaleString() + ' ／ ' + r.かかった秒 + '秒');
  });
}

console.log('');
console.log('  ' + pass + ' 緑 / ' + fail + ' 赤');

if (SELF) {
  const { spawnSync } = await import('node:child_process');
  const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'exally-shindan-'));
  console.log('');
  console.log('[self-test] わざと壊して 赤くなるかを数える（★repo は読むだけ★）');
  const BREAKS = [
    ['★包まれていない物まで 数える（素朴なやり方に戻す）★',
      (s) => s.replace('  function 隠している所(f) {', '  function 隠している所(f) {\n    if (/IFERROR|IFNA/i.test(String(f)) && /#REF!/.test(String(f))) return [{ 関数: "IFERROR", 中身: "", 数: (String(f).match(/#REF!/g) || []).length }];')],
    ['★2つ目の引数の側まで 数える★', (s) => s.replace("else if (ch === ',' && 深さ === 1) break;", '')],
    ['★式でない字（打った字）まで 見る★', (s) => s.replace("if (!f || String(f).charAt(0) !== '=') continue;", 'if (!f) continue;')],
    ['★のべ と 式の本数を 同じにする★', (s) => s.replace('式の本数: Object.keys(本).length,', '式の本数: のべ,')],
    ['★場所を出さない★', (s) => s.replace('セル: セルの名(rc),', "セル: '',")],
    ['★今 何が出ているかを 出さない★', (s) => s.replace("いま出ている物: (cell.v === undefined || cell.v === null || cell.v === '') ? '（空）' : String(cell.v),", "いま出ている物: '',")],
    ['★手掛かりを 出さない★', (s) => s.replace('手掛かり: 手掛かり(f),', '手掛かり: [],')],
    ['★小分けをやめる（画面が固まる）★', (s) => s.replace('var 一度に = opt.一度に || 3000;', 'var 一度に = 1e9;')],
    ['★客の字に ★ を書く★', (s) => s.replace("題: '見えない所で 空になっています',", "題: '★見えない所で 空になっています★',")],
    ['★「壊れています」と言う★', (s) => s.replace("題: '見えない所で 空になっています',", "題: '壊れています',")],
    ['★何が起きるか（合計が小さくなる）を 言わない★', (s) => s.replace("なぜ: 'そのまま合計すると、合計だけが 静かに 小さくなります。',", "なぜ: '',")],
    ['★列の字を 間違える（AAにならない）★', (s) => s.replace('c = Math.floor(c / 26) - 1;', 'c = Math.floor(c / 26);')],
    ['★調べた時間と 出るまでの時間を 混ぜる★',
      (s) => s.replace('かかった秒: Math.max(0.01, Math.round(仕事 / 10) / 100),   /* ★実際に調べた時間★ */', 'かかった秒: Math.round((t1 - t0) / 10) / 100,')
        .replace('出るまでの秒: Math.round((t1 - t0) / 10) / 100,            /* ★待っていた時間も含む★ */', '')],
  ];
  let red = 0;
  for (const [name, brk] of BREAKS) {
    const orig = fs.readFileSync(path.join(ROOT, 'lib/shindan.js'), 'utf8');
    const bad = brk(orig);
    if (bad === orig) { console.log('  ★置換できず★  ' + name); continue; }
    const f = path.join(TMP, 'shindan.js');
    fs.writeFileSync(f, bad, 'utf8');
    const env = Object.assign({}, process.env, { EXALLY_SHINDAN_OVERRIDE: JSON.stringify({ 'lib/shindan.js': f }) });
    const r2 = spawnSync(process.execPath, ['--max-old-space-size=4096', path.join(__dirname, 'shindan.test.mjs')], { encoding: 'utf8', env });
    if (r2.status !== 0) { red++; console.log('  赤くなった  ' + name); }
    else console.log('  ★素通り★  ' + name);
  }
  try { fs.rmSync(TMP, { recursive: true, force: true }); } catch (e) { /* 消せなくても検査は済んでいる */ }
  console.log('');
  console.log('  ' + red + '/' + BREAKS.length + ' 通りで赤くなった');
  process.exit(red === BREAKS.length ? 0 : 1);
}

process.exit(fail ? 1 : 0);
