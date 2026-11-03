/* vba.test.mjs — ★マクロ（VBA）を 読む・見立てる★
 * =============================================================================
 * ★決まり（指示役 2026-08-28）★
 *   ・★試験用の .xlsm は うちで作る★（客の実物は 使わない）
 *     … tests/fixtures/vba-sample.xlsm ／ 元の字は tests/fixtures/vba-src/*.bas（cp932）
 *   ・★「読めた」と「正しく読めた」を分ける★
 *     ＝取り出した中身を ★もう一度 圧縮して 元のバイトと 1バイトずつ 比べる★
 *   ・★VBAは 動かさない★。読むだけ。
 *   ・★AIには 要約だけ★＝渡す物に ★客のコードが 1文字も 入っていない★事を 数える
 *
 * ★踏んだ穴（測り方の話・2026-08-28）★
 *   ・zip の中身を ZipSurgeon.toBytes(z, e) で数えた ＝ toBytes は「字→バイト」の道具で、
 *     String(z) の "[object Object]" の ★15文字★を ずっと数えていた。
 *     ⇒ ★中身は z.bytes(名) で取る★。この試験は それを 1本 固定する。
 *   ・中身は Unicode では 置かれていない（日本語Excelは cp932）。
 *     1バイトずつ 読むと 日本語が 化ける。⇒ dir の 文字の種類で 読む。
 *
 * 使い方: node tests/vba.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';
import { spawnSync } from 'node:child_process';
import { 入れ物を組み立てる } from '../scripts/make-vba-fixture.mjs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const SELF = process.argv.includes('--self-test');
const OVERRIDE = process.env.EXALLY_VBA_OVERRIDE ? JSON.parse(process.env.EXALLY_VBA_OVERRIDE) : {};
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const V = require_(OVERRIDE['lib/vba.js'] || path.join(ROOT, 'lib/vba.js'));
const M = require_(OVERRIDE['lib/vba-mikata.js'] || path.join(ROOT, 'lib/vba-mikata.js'));
const XLSX = require_(path.join(ROOT, 'lib/xlsx.full.min.js'));
const ZS = require_(path.join(ROOT, 'lib/zip-surgeon.js'));

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ok   ' + n); } catch (e) { fail++; console.log('  NG   ' + n + '\n       ' + (e && e.message)); } };
const ok = (c, m) => { if (!c) throw new Error(m || 'expected truthy'); };
const eq = (a, b, m) => { if (a !== b) throw new Error((m ? m + ': ' : '') + '期待=' + JSON.stringify(b) + ' 実際=' + JSON.stringify(a)); };

console.log('');
console.log('[vba] ★マクロを 読む・見立てる（AIは0回）★');

/* ══ ①詰め方（MS-OVBA）の 往復＝★境界を 実物で 測る★ ══════════ */
console.log('  -- 詰め方の往復（端っこを 全部）--');
/* 4096 は 1かたまりの 境目。0/1 は 空と1文字。★端で 落ちるのが この形の癖★ */
const 端 = [0, 1, 2, 3, 4095, 4096, 4097, 8191, 8192, 8193, 12000];
for (const n of 端) {
  T('往復 ' + n + 'バイト（同じ字が続く＝よく縮む）', () => {
    const 元 = new Uint8Array(n).fill(0x41);
    const r = V.解凍(V.圧縮(元));
    ok(r.ok, '解凍できない: ' + r.なぜ);
    eq(r.出.length, n, '長さ');
    for (let i = 0; i < n; i++) if (r.出[i] !== 元[i]) throw new Error(i + '番目が 違う');
  });
}
T('往復 30000バイト（ばらばら＝縮まない）', () => {
  /* ★毎回 同じ物を作る★（Math.random は 使わない＝落ちた時に 同じ物で 追えなくなる） */
  const n = 30000, 元 = new Uint8Array(n);
  let x = 12345;
  for (let i = 0; i < n; i++) { x = (x * 1103515245 + 12345) & 0x7FFFFFFF; 元[i] = x & 0xFF; }
  const r = V.解凍(V.圧縮(元));
  ok(r.ok, '解凍できない: ' + r.なぜ);
  eq(r.出.length, n, '長さ');
  for (let i = 0; i < n; i++) if (r.出[i] !== 元[i]) throw new Error(i + '番目が 違う');
});
T('印（0x01）が無い物は 読めないと言う（黙って0本にしない）', () => {
  const r = V.解凍(new Uint8Array([0x99, 0x00, 0x00]));
  eq(r.ok, false);
  ok(r.なぜ && r.なぜ.length > 0, 'なぜ が 空');
});

/* ══ ①-b 決め手になる 2つの道具を 直に 押す ═══════════════════
   ★ここを 押さないと 壊しても 赤くならない★（2026-08-28 実測で 2本 素通りした） */
console.log('  -- 決め手の道具 --');
T('★長さが 同じで 中身が 違う物を 同じと言わない★', () => {
  eq(V.同じバイト列(Uint8Array.from([1, 2, 3]), Uint8Array.from([1, 2, 3])), true, '同じ物');
  eq(V.同じバイト列(Uint8Array.from([1, 2, 3]), Uint8Array.from([1, 2, 4])), false, '最後だけ 違う');
  eq(V.同じバイト列(Uint8Array.from([9, 2, 3]), Uint8Array.from([1, 2, 3])), false, '頭だけ 違う');
  eq(V.同じバイト列(Uint8Array.from([1, 2]), Uint8Array.from([1, 2, 3])), false, '長さが 違う');
  eq(V.同じバイト列(null, Uint8Array.from([1])), false, '片方が 無い');
});
T('★字を読む道具が 無い時は 確か=false と言う★', () => {
  const 元 = globalThis.TextDecoder;
  /* big5（950）だけ 使えない機械の ふりをする＝この道は 機械によって 通る／通らないが 変わる */
  globalThis.TextDecoder = function (名, o) {
    if (String(名) === 'big5') throw new Error('この機械には 無い');
    return new 元(名, o);
  };
  try {
    const r = V.中身を字にする(Uint8Array.from([0x41, 0x42]), 950);
    eq(r.確か, false, '確か');
    eq(r.字づかい, 'big5', '字づかい');
    eq(r.文, 'AB', '中身（1バイトずつでも 読めた物は 出す）');
  } finally { globalThis.TextDecoder = 元; }
});
T('知らない字づかいの番号も 確か=false（番号を そのまま 言う）', () => {
  const r = V.中身を字にする(Uint8Array.from([0x41]), 1);
  eq(r.確か, false);
  eq(r.字づかい, '1');
});

/* ══ ②試験用 .xlsm を 読む ═════════════════════════════════ */
console.log('  -- 試験用の .xlsm（うちで作った物）--');
const 見本 = path.join(ROOT, 'tests/fixtures/vba-sample.xlsm');
ok(fs.existsSync(見本), '見本が 無い: node scripts/make-vba-fixture.mjs で 作る');
const z = ZS.read(new Uint8Array(fs.readFileSync(見本)));
const bin = await z.bytes('xl/vbaProject.bin');

T('マクロの入れ物が 中に 在る（★字→バイトの道具で 数えない★）', () => {
  ok(bin.length > 2000, 'vbaProject.bin が 小さすぎる: ' + bin.length + 'バイト');
  eq(bin[0], 0xD0, 'CFBの印(1)'); eq(bin[1], 0xCF, 'CFBの印(2)');
});

const 読み = V.読む(bin, XLSX.CFB);
T('読めた', () => { ok(読み.ok, 読み.なぜ); });
T('モジュールが 4本', () => eq(読み.モジュール.length, 4));
T('★正しく読めた 4/4（証拠3つが 全部 通る）★', () => {
  eq(読み.確かめ.一致, 4, '正しく読めた');
  eq(読み.確かめ.詰め方も同じ, 4, '元のバイトと 1バイトずつ 同じ');
  eq(読み.確かめ.名前一致, 4, '目次の名前と 中身の1行目');
  eq(読み.確かめ.往復, 4, '往復');
  eq(読み.確かめ.全部, 4, '全部');
  eq(読み.確かめ.目次, 4, '目次が 言う本数');
});
T('種類を 分けている（手続き3・クラス1）', () => {
  eq(読み.モジュール.filter((m) => m.種類 === '手続き').length, 3);
  eq(読み.モジュール.filter((m) => m.種類 === 'クラス').length, 1);
});
T('★日本語が 化けていない★（cp932で 読む）', () => {
  const m1 = 読み.モジュール.find((m) => m.名 === 'Module1');
  ok(m1, 'Module1 が 無い');
  eq(m1.字づかい, 'shift_jis', '字づかい');
  ok(m1.中身.indexOf('Sub 月次締め()') >= 0, '「Sub 月次締め()」が 出ない');
  ok(m1.中身.indexOf('MsgBox "締めました"') >= 0, '「締めました」が 出ない');
  ok(m1.中身.indexOf('�') < 0, '化けた字（置き換え文字）が 混ざっている');
});
T('★元の .bas と 1文字ずつ 同じ★（読んだ物 と 作った元）', () => {
  for (const 名 of ['Module1', 'Module2', 'Module3', 'Sheet1']) {
    const 元 = fs.readFileSync(path.join(ROOT, 'tests/fixtures/vba-src', 名 + '.bas'));
    const 期待 = new TextDecoder('shift_jis').decode(元);
    const 実 = 読み.モジュール.find((m) => m.名 === 名).中身;
    if (実 !== 期待) {
      for (let i = 0; i < Math.max(実.length, 期待.length); i++) {
        if (実[i] !== 期待[i]) throw new Error(名 + ' の ' + i + '文字目が 違う（期待=' + JSON.stringify(期待[i]) + ' 実際=' + JSON.stringify(実[i]) + '）');
      }
    }
  }
});
T('中身が 無い時は「マクロが ありません」と言う（0本と言わない）', () => {
  const r = V.読む(new Uint8Array(0), XLSX.CFB);
  eq(r.ok, false);
  ok(r.なぜ.indexOf('ありません') >= 0, r.なぜ);
});

/* ══ ②-b ★「読めた」と「正しく読めた」は 別物★ ══════════════
   ★中身は同じ・詰め方だけ 違う物★ と ★読めない字づかいの物★ を その場で組み立てて 測る。
   （組み立てる道具は 見本を作る script と 同じ物を 呼ぶ＝写しを 作らない） */
console.log('  -- 読めた ≠ 正しく読めた --');
T('★詰め方が違う物は「詰め方が違う」と 分けて 言う（0本に しない）★', () => {
  /* ★実物では ここが 普通★＝Excelの詰め方と うちの詰め方は 違う。
     司さんの実物 59本 中 ★元のバイトと一致 0本／名前が合う 59本★（2026-08-28 実測）。
     だから ★元のバイトと一致 だけを「正しく読めた」にしない★。 */
  const r = V.読む(new Uint8Array(入れ物を組み立てる({ 詰めない: true })), XLSX.CFB);
  ok(r.ok, r.なぜ);
  eq(r.モジュール.length, 4, '本数');
  eq(r.確かめ.詰め方も同じ, 0, '元のバイトと 1バイトずつ 同じ');
  eq(r.確かめ.往復, 4, '往復');
  eq(r.確かめ.名前一致, 4, '名前');
  eq(r.確かめ.一致, 4, '正しく読めた（往復＋名前＋字づかい）');
  for (const m of r.モジュール) {
    eq(m.確か, true, m.名);
    eq(m.詰め方も同じ, false, m.名 + ' の 詰め方');
    eq(m.なぜ, '中身は同じ・詰め方が違う', m.名);
  }
  const m1 = r.モジュール.find((m) => m.名 === 'Module1');
  ok(m1.中身.indexOf('Sub 月次締め()') >= 0, '中身が 読めていない');
});
T('★目次の名前と 中身の1行目が 合わなければ 未測定★（読み違えたら 合わない）', () => {
  const r = V.読む(new Uint8Array(入れ物を組み立てる({ 名前をずらす: true })), XLSX.CFB);
  ok(r.ok, r.なぜ);
  eq(r.確かめ.名前一致, 0, '名前');
  eq(r.確かめ.一致, 0, '正しく読めた');
  for (const m of r.モジュール) {
    eq(m.確か, false, m.名);
    ok(m.なぜ.indexOf('1行目') >= 0, m.名 + ' の なぜ: ' + m.なぜ);
  }
  eq(M.見立てる(r.モジュール).本数, 0, '見立ててしまっている');
});
T('★目次が 途中で 読めなくなったら そう言う（0本と 言わない）★', () => {
  const r = V.読む(new Uint8Array(入れ物を組み立てる({ 目次を壊す: true })), XLSX.CFB);
  eq(r.ok, false, '読めたと 言っている');
  eq(r.モジュール.length, 0, '本数');
  ok(r.なぜ.indexOf('読めなく') >= 0, 'なぜ: ' + r.なぜ);
});
T('★目次の 大きさの 嘘（版の記録）を 直せている★', () => {
  /* この記録は「4」と書いて 中身は 6バイト。直さないと ここから先が 総崩れになる。 */
  eq(読み.確かめ.目次, 4, '目次が 言う本数');
  eq(読み.モジュール.length, 4, '取れた本数');
  eq(読み.なぜ, '', 'なぜ: ' + 読み.なぜ);
});
T('★読めない字づかいは 確か=false・理由を言う（未測定にする）★', () => {
  const r = V.読む(new Uint8Array(入れ物を組み立てる({ 字の種類: 1 })), XLSX.CFB);
  ok(r.ok, r.なぜ);
  eq(r.確かめ.一致, 0, '正しく読めた');
  for (const m of r.モジュール) {
    eq(m.確か, false, m.名 + ' が 確か になっている');
    ok(m.なぜ.indexOf('字づかい') >= 0, m.名 + ' の なぜ: ' + m.なぜ);
  }
  /* ★そのまま 見立てに 流さない★ */
  eq(M.見立てる(r.モジュール).本数, 0, '見立ててしまっている');
});

/* ══ ③見立て（機械が 分ける・AIは0回） ═══════════════════════ */
console.log('  -- 見立て（分ける・可否）--');
const 見立て = M.見立てる(読み.モジュール);
T('手続きに 5本 切れた', () => eq(見立て.本数, 5));
T('切った所（名前・行数）が 合っている', () => {
  const 名 = 見立て.手続き.map((v) => v.名).join(',');
  eq(名, '月次締め,印刷,CSV取り込み,月次の並べ替え,Worksheet_Change');
  eq(見立て.手続き[1].行数, 3, '印刷の行数');
});
T('印刷は「そのまま できる」', () => {
  const v = 見立て.手続き.find((x) => x.名 === '印刷');
  eq(v.可否, 'できる');
  ok(v.うちのやり方.indexOf('印刷') >= 0, v.うちのやり方);
});
T('★きっかけ（Worksheet_Change）を 先に見る★', () => {
  const v = 見立て.手続き.find((x) => x.名 === 'Worksheet_Change');
  eq(v.分類[0].名, '自動で動く（きっかけ）', '先頭が きっかけ ではない');
  eq(v.可否, 'かえる');
});
T('★段取りの話（ちらつき止め）を 先頭に 出さない★（数だけで 並べない）', () => {
  const 文 = ['Sub 集計()', '  Application.ScreenUpdating = False',
    '  Range("A1:D100").Sort Key1:=Range("A1")',
    '  Application.ScreenUpdating = True', 'End Sub'].join('\n');
  const v = M.見立てる1本(M.手続きに切る(文)[0]);
  /* ちらつき止めは 2回・並べ替えは 1回＝数だけで 並べると 段取りが 先頭に来る */
  eq(v.分類[0].名, '並べ替え', '先頭が 仕事の中身ではない（' + v.分類.map((a) => a.名).join(',') + '）');
  ok(v.何をしているか.indexOf('表を 並べ替えています。') === 0, v.何をしているか);
});
T('取り込みは GetOpenFilename で 当てている', () => {
  const v = 見立て.手続き.find((x) => x.名 === 'CSV取り込み');
  ok(v.分類.some((a) => a.key === 'torikomi'), '取り込みが 出ない');
});
T('★End が 無いまま 次が 始まる物を 1本に まとめない★（実物に 在る）', () => {
  /* 司さんの実物 5件＝End Sub が 無い Sub の 次に 別の Sub が 始まっていた。
     まとめると ★次の手続きの中身が 前の物として 数えられる★。 */
  const 文 = ['Attribute VB_Name = "a"', 'Private Sub 先()', 'X: 1', '',
    'Private Sub 後(ByVal T As Range)', '    T.Value = 1', 'End Sub', ''].join('\r\n');
  const たち = M.手続きに切る(文);
  eq(たち.length, 2, '本数');
  eq(たち[0].名, '先'); eq(たち[0].閉じていない, true, '先が 閉じていない と 言えていない');
  eq(たち[1].名, '後'); eq(たち[1].閉じていない, false, '後まで 閉じていない 扱い');
  /* ★前の物に 次の中身を 混ぜていない★ */
  ok(たち[0].中身.indexOf('T.Value') < 0, '次の手続きの中身を 前の物に 混ぜている');
  ok(たち[1].中身.indexOf('T.Value') >= 0, '次の手続きの中身が 落ちている');
});
T('★別の道★を 別の道と言う（メール・外のアプリ・ファイル直・DB）', () => {
  const 例 = [
    ['Sub a()\n Set o = CreateObject("Outlook.Application")\n o.Send\nEnd Sub', 'メールを出す'],
    ['Sub b()\n Set fso = CreateObject("Scripting.FileSystemObject")\n fso.CopyFile x, y\nEnd Sub', 'ファイルを直に触る'],
    ['Sub c()\n Set cn = CreateObject("ADODB.Connection")\n cn.Open s\nEnd Sub', 'データベースにつなぐ'],
    ['Sub d()\n Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long\nEnd Sub', 'Windowsの機能を直に呼ぶ'],
  ];
  for (const [文, 名] of 例) {
    const v = M.見立てる1本(M.手続きに切る(文)[0]);
    eq(v.可否, 'べつの道', 名 + ' の可否');
    ok(v.分類.some((a) => a.名 === 名), 名 + ' が 出ない（出たのは ' + v.分類.map((a) => a.名).join(',') + '）');
  }
});
T('★一番 重い物に合わせる★（できる＋別の道 なら 別の道）', () => {
  const v = M.見立てる1本(M.手続きに切る('Sub e()\n Range("A1").Sort\n Set o = CreateObject("Outlook.Application")\nEnd Sub')[0]);
  eq(v.可否, 'べつの道');
});
T('★分からない物は「わからない」と言う★（当てずっぽうで 分けない）', () => {
  const v = M.見立てる1本(M.手続きに切る('Sub f()\n Dim x As Long\n x = 1 + 2\nEnd Sub')[0]);
  eq(v.可否, 'わからない');
  eq(v.分類.length, 0);
  ok(v.何をしているか.indexOf('読み取れません') >= 0, v.何をしているか);
});
T('★正しく読めなかった物は 見立てない（未測定にする）★', () => {
  const r = M.見立てる([{ 名: 'X', 確か: false, なぜ: '往復で 中身が 変わった', 中身: 'Sub g()\n Range("A1").Sort\nEnd Sub' }]);
  eq(r.本数, 0, '見立ててしまっている');
  eq(r.未測定.length, 1);
  ok(M.知らせの字(r).indexOf('正しく読めた') >= 0, M.知らせの字(r));
});

/* ══ ④AIに渡す物＝★要約だけ★ ══════════════════════════════ */
console.log('  -- AIに渡す形（要約だけ）--');
const 渡す = JSON.stringify(M.AIに渡す形(見立て));
T('★客のコードが 1文字も 入っていない★', () => {
  /* 見本の中に 出てくる 客の側の字（手続きの名前は 除く＝名前は 要約に要る） */
  /* ★中身は 見立ての中に 持っている★（手順を 取り出す側が 使う）ので、
     ★AIに渡す形が それを 拾っていないか★を ここで 数える。 */
  const だめ = ['明細', '最終行', '取込', 'Sheets(', 'Range(', 'MsgBox "', 'End Sub', 'Cells(', '1.1',
    'Attribute', 'Sub ', 'Option Explicit'];
  const 出た = だめ.filter((w) => 渡す.indexOf(w) >= 0);
  eq(出た.join(','), '', '入ってしまった字');
});
T('渡す物は 元のコードより 短い', () => {
  const 元 = 読み.モジュール.reduce((a, m) => a + m.中身.length, 0);
  ok(渡す.length < 元, '要約 ' + 渡す.length + '文字 ≧ 元 ' + 元 + '文字');
});
T('★手掛かりは うちの言葉だけ（表に書いた印と 一致）★', () => {
  const 印 = new Set(M.分け方.map((d) => d.印));
  for (const v of 見立て.手続き) for (const h of v.手掛かり) ok(印.has(h), 'うちの表に無い字: ' + JSON.stringify(h));
});
T('分け方の 印は 全部 在って ASCIIだけ（客の字が 混ざらない形）', () => {
  for (const d of M.分け方) {
    ok(d.印 && d.印.length, d.key + ' に 印が 無い');
    /* eslint-disable no-control-regex */
    ok(/^[\x20-\x7E]+$/.test(d.印), d.key + ' の印に ASCII以外: ' + d.印);
  }
});
T('分け方の keyが 重なっていない', () => {
  const k = M.分け方.map((d) => d.key);
  eq(new Set(k).size, k.length, '同じkeyが 2つ');
});

/* ══ ⑤客に見せる字 ═══════════════════════════════════════ */
console.log('  -- 客に見せる字 --');
T('1行の 知らせが 本数と 内訳を言う', () => {
  const s = M.知らせの字(見立て);
  ok(s.indexOf('5本') >= 0, s);
  ok(s.indexOf('できる') >= 0, s);
});
T('★客に見せる字に ★ を書かない★', () => {
  const 見る = [M.知らせの字(見立て)];
  for (const v of 見立て.手続き) { 見る.push(v.何をしているか); 見る.push(v.うちのやり方); }
  for (const m of 読み.モジュール) 見る.push(m.なぜ || '');
  for (const s of 見る) ok(s.indexOf('★') < 0, '★が 入っている: ' + s);
});
T('★「壊れています」と言わない★', () => {
  const 見る = [M.知らせの字(見立て)].concat(見立て.手続き.map((v) => v.何をしているか));
  for (const s of 見る) ok(s.indexOf('壊れ') < 0, s);
});

console.log('');
console.log('  ' + pass + '緑 / ' + fail + '赤');

/* ══ ⑥壊して 赤くなるか（★1本ずつ★） ══════════════════════ */
if (SELF) {
  console.log('');
  console.log('[vba --self-test] ★壊したら 赤くなるか★');
  const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'exally-vba-'));
  const BREAKS = [
    /* ★当てる所を 間違えると 素通りする★＝2026-08-28、この2本は
       「決めていない行」を 壊していた（次の行で 数え直していた／その前で return していた）。 */
    ['vba.js', '★確かめを 全部 素通りさせる★',
      (s) => s.replace('var 正しい = !!(往復同じ && 名前が合う && よ.確か);', 'var 正しい = true;')],
    ['vba.js', '★目次の名前と 中身の1行目を 見ない★',
      (s) => s.replace('var 名前が合う = !!(頭 && 頭[1] === x.名);', 'var 名前が合う = true;')],
    /* ★「往復を 見ない」は 壊しても 赤くならない★＝往復は うちの 圧縮と 解凍が
       食い違った時だけ 落ちる物で、それは ①の 往復（端っこ12通り）で 直に 測っている。
       ここに 置いても 一度も 赤くならない＝★置かない★（在るだけの 見張りを 作らない）。 */
    ['vba.js', '★目次の 行き止まりを 黙って 飲み込む★',
      (s) => s.replace("行き止まり = '目次の 途中で 読めなくなりました';", '')],
    ['vba.js', '★目次の 大きさの 嘘（0x0009）を 直さない★',
      (s) => s.replace('if (id === 0x0009) 大きさ = 6;', '')],
    ['vba.js', '★0本でも ok と言う★',
      (s) => s.replace("      return { ok: false, モジュール: [], なぜ: 目次.行き止まり || '目次に マクロが 1本も 書かれていません' };",
        '      return { ok: true, モジュール: [], なぜ: 0, 確かめ: { 一致: 0, 全部: 0 } };')],
    ['vba.js', '★1バイトずつ 比べるのを やめる（長さだけ見る）★',
      (s) => s.replace('for (var i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;', '')],
    ['vba.js', '★かたまりの頭の 大きさを 1 ずらす★',
      (s) => s.replace('var 大きさ = (頭 & 0x0FFF) + 3;', 'var 大きさ = (頭 & 0x0FFF) + 4;')],
    ['vba.js', '★字づかいを 見ずに 1バイトずつ読む（日本語が化ける）★',
      (s) => s.replace('var よ = 中身を字にする(r.出, 字づかい);', 'var よ = { 文: 字にする(r.出), 字づかい: "1バイト", 確か: true };')],
    ['vba.js', '★読めない字づかいでも 確か と言う★',
      (s) => s.replace("if (!名) return { 文: 字にする(b), 字づかい: (cp ? String(cp) : '不明'), 確か: false };",
        "if (!名) return { 文: 字にする(b), 字づかい: (cp ? String(cp) : '不明'), 確か: true };")],
    ['vba.js', '★道具が 無い時も 確か と言う★',
      (s) => s.replace('if (!d) return { 文: 字にする(b), 字づかい: 名, 確か: false };', 'if (!d) return { 文: 字にする(b), 字づかい: 名, 確か: true };')],
    ['vba.js', '★中身が無い時に 空で ok と言う★',
      (s) => s.replace("if (!bin || !bin.length) return { ok: false, モジュール: [], なぜ: 'マクロが ありません' };", 'if (!bin || !bin.length) return { ok: true, モジュール: [], なぜ: 0 };')],
    ['vba-mikata.js', '★正しく読めていない物も 見立てる★',
      (s) => s.replace('if (!m || !m.確か) {', 'if (false) {')],
    ['vba-mikata.js', '★分からない物を「できる」と言う★',
      (s) => s.replace("        分類: [], 可否: 'わからない',", "        分類: [], 可否: 'できる',")],
    ['vba-mikata.js', '★可否を 一番 軽い物に合わせる★',
      (s) => s.replace('if (重さ[当たり[k].可否] > 重さ[可否]) 可否 = 当たり[k].可否;', 'if (重さ[当たり[k].可否] < 重さ[可否]) 可否 = 当たり[k].可否;')],
    ['vba-mikata.js', '★AIに渡す形に 手続きの中身を 入れる★',
      (s) => s.replace('        名: v.名, 種類: v.種類, 行数: v.行数, 可否: v.可否,',
        '        名: v.名, 種類: v.種類, 行数: v.行数, 可否: v.可否, 中身: v.中身,')],
    ['vba-mikata.js', '★AIに 中身を そのまま渡す★',
      (s) => s.replace('        分類: v.分類.map(function (a) { return a.名; }),', '        分類: v.分類.map(function (a) { return a.名; }), 中身: v.何をしているか + JSON.stringify(v),')],
    ['vba-mikata.js', '★手掛かりに 客のコードを 入れる★',
      (s) => s.replace("数: n, 手掛かり: d.印 || d.key,", "数: n, 手掛かり: 中.slice(0, 60),")],
    ['vba-mikata.js', '★End が 無い物を 次と 1本に まとめる★',
      (s) => s.replace('if (m && 今) 閉じる(true);', '')],
    ['vba-mikata.js', '★段取りの話を 先頭に 出す（数だけで 並べる）★',
      (s) => s.replace('var 順 = function (x) { return x.先に見る ? 0 : (x.後回し ? 2 : 1); };',
        'var 順 = function (x) { return 0; };')],
    ['vba-mikata.js', '★きっかけを 後回しにする★',
      (s) => s.replace('return x.先に見る ? 0 : (x.後回し ? 2 : 1);', 'return x.後回し ? 2 : 1;')],
    ['vba-mikata.js', '★客に見せる字に ★ を書く★',
      (s) => s.replace("var s = 'マクロが '", "var s = '★マクロが '")],
  ];
  let red = 0;
  for (const [対象, name, brk] of BREAKS) {
    const 元 = fs.readFileSync(path.join(ROOT, 'lib', 対象), 'utf8');
    const 壊 = brk(元);
    if (壊 === 元) { console.log('  ★置換できず★  ' + name); continue; }
    const f = path.join(TMP, 対象);
    fs.writeFileSync(f, 壊, 'utf8');
    const env = Object.assign({}, process.env, { EXALLY_VBA_OVERRIDE: JSON.stringify({ ['lib/' + 対象]: f }) });
    const r = spawnSync(process.execPath, [path.join(__dirname, 'vba.test.mjs')], { encoding: 'utf8', env });
    if (r.status !== 0) { red++; console.log('  赤くなった  ' + name); }
    else console.log('  ★素通り★  ' + name);
  }
  try { fs.rmSync(TMP, { recursive: true, force: true }); } catch (e) { /* 消せなくても 検査は済んでいる */ }
  console.log('');
  console.log('  ' + red + '/' + BREAKS.length + ' 通りで赤くなった');
  process.exit(red === BREAKS.length ? 0 : 1);
}

process.exit(fail ? 1 : 0);
