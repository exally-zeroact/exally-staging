/* hyou-no-soto.test.mjs — ★診断2本目「ほかの表の その行を 見ている」★ 2026-09-04
 *
 *  ★なぜ 作ったか（司さんの 実物で 実際に 出た）★
 *    2026-09-03 に 8月の 行を コピーして 9月を 作った時、
 *    ★AF・AH・AJ の 3列 × 31行 が 8月の 表(Table23) の [@列名] を 指したまま★ 残っていた。
 *    ⇒ 「その行」は ★自分の表の 中でしか 決まらない★ので Excel も 答えを 出せない
 *    ⇒ IFERROR に 包まれているので ★画面は 0★＝★エラーが 出ず 誰も 気づけない★
 *    ⇒ ★数字を 入れても ずっと 0 のまま★
 *
 *  ★司さんの 指摘（2026-09-04）★
 *    「分かったんなら ★理由と 修正方法まで★ 教えれないかんやろが
 *      それが 出来ずに Exally内でも ユーザーに 聞くつもりか？」
 *    ⇒ ★聞かない★。★見つけて・なぜかを 言って・直し方（前→後）まで 出す★。
 *
 *  ★ここでは 直さない★＝式を 書き換えるのは 客の 決め。★言うだけ★。
 *
 *  走らせ方: node tests/hyou-no-soto.test.mjs [--self-test]
 */
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const TR = require_(path.join(ROOT, 'lib/table-refs.js'));
const Shindan = require_(path.join(ROOT, 'lib/shindan.js'));

let pass = 0, fail = 0;
const T = (n, よい, 添え) => {
  if (よい) { pass++; console.log('  ok   ' + n); }
  else { fail++; console.log('  NG   ' + n + (添え ? '\n       ' + 添え : '')); }
};

/* ★実物と 同じ 形で 作る★（計算シート・8月=Table23・9月=Table24）
   行は 0 から 数える。実物 … Table23 行233〜265 ／ Table24 行267〜299（見出し1・合計1） */
/* ★名前は 司さんの 実物と 同じ★＝★客が 数式バーで 見る 名前は R8.8 / R8.9★
   （SheetJS が 書く Table23 / Table24 は ★客の 画面には 出ない★＝それで 助言すると 探せない） */
/* ★列の 一覧★＝32番目が 標本七ｈ・34番目が 雛形八一ｈ（実物と 同じ 並び） */
const 列たち = (function () { var a = []; for (var i = 0; i < 38; i++) a.push('列' + i); a[32] = '標本七ｈ'; a[34] = '雛形八一ｈ'; a[5] = '売上'; return a; })();
const 表 = [
  { id: 23, name: 'R8.8', sheet: '計算', rw1: 233, rw2: 265, cl1: 0, cl2: 37, header: 1, totals: 1, cols: 列たち },
  { id: 24, name: 'R8.9', sheet: '計算', rw1: 267, rw2: 299, cl1: 0, cl2: 37, header: 1, totals: 1, cols: 列たち },
];
const byId = { 23: 表[0], 24: 表[1] };
/* 見出しの 字（列名を 出す為）。Table23 と Table24 の 33列目＝AG＝「標本七ｈ」 */
const wb = { Sheets: { 計算: { AG234: { v: '標本七ｈ' }, AG268: { v: '標本七ｈ' } } } };
const その行 = (id, col) => ({ id: id, code: 16, coltype: 1, colFirst: col, colLast: col });
const 中身 = (id, col) => ({ id: id, code: 0, coltype: 1, colFirst: col, colLast: col });

console.log('');
console.log('[hyou-no-soto] ★ほかの表の「その行」を 見ている所★');

/* ── ① 実物と 同じ形＝見つける ── */
{
  /* 計算!AF269（0から数えると 行268・列31）が Table23[@標本七ｈ] を 指す */
  const x = TR._断りを仕分ける('計算', 268, 31, [その行(24, 12), その行(23, 32)], byId, 表, wb);
  T('★ほかの表の その行を 指していたら 見つける★', !!x, '見つけていない');
  T('★種類が thisrow_other_table★', x && x.種類 === 'thisrow_other_table');
  T('★場所が セルの名前で 出る（AF269）★', x && x.セル === 'AF269', x && x.セル);
  T('★客が 見る 名前で 出す（R8.8 → R8.9／Table23 では ない）★',
    x && x.指した表 === 'R8.8' && x.自分の表 === 'R8.9',
    x && (x.指した表 + ' / ' + x.自分の表));
  T('★列名を 見出しから 取る（標本七ｈ）★', x && x.列名 === '標本七ｈ', x && String(x.列名));
  T('★直し方が 前→後 で 出る★',
    x && x.直し方 && x.直し方.前 === 'R8.8' && x.直し方.後 === 'R8.9');
  T('★直し方の 字＝何を 何に 変えるか（中の言葉を 出さない）★',
    Shindan.直し方の字(x) === 'R8.8 を R8.9 に 変える（標本七ｈ の 所）',
    Shindan.直し方の字(x));
  T('★行と列を そのまま 出す（画面が その場所へ 飛べる）★', x && x.r === 268 && x.c === 31);
}

/* ── ② 出しては いけない物（★空振りで 客を 驚かせない★）── */
T('★自分の表の その行は 出さない★',
  TR._断りを仕分ける('計算', 268, 31, [その行(24, 32)], byId, 表, wb) === null);
T('★その行では ない（[#Data]）物は 出さない★',
  TR._断りを仕分ける('計算', 268, 31, [中身(23, 32)], byId, 表, wb) === null);
T('★どの表にも 入っていない セルは 出さない（自分の表が 分からない）★',
  TR._断りを仕分ける('計算', 500, 31, [その行(23, 32)], byId, 表, wb) === null);
T('★知らない表を 指している時は 出さない（当て推量しない）★',
  TR._断りを仕分ける('計算', 268, 31, [その行(99, 32)], byId, 表, wb) === null);
{
  /* ★ほかの表でも その行が 中に 在るなら 出さない★（Excel は 答えを 出せる） */
  const 重なり = [
    { id: 23, name: 'R8.8', sheet: '計算', rw1: 233, rw2: 299, cl1: 0, cl2: 37, header: 1, totals: 1, cols: 列たち },
    表[1],
  ];
  T('★ほかの表でも その行が 中身の 行の 中なら 出さない★',
    TR._断りを仕分ける('計算', 268, 31, [その行(23, 32)], { 23: 重なり[0], 24: 表[1] }, 重なり, wb) === null);
}
/* ★自己確認が 教えてくれた 抜け その1★
   ★自分の表の その行でも、合計の行から 呼べば 中身の 行の 外★になる。
   その時に 出してしまうと 直し方が「Table24→Table24」＝★意味の 無い 助言★に なる。 */
T('★自分の表なら 合計の行から 呼んでも 出さない（前＝後の 助言を 出さない）★',
  TR._断りを仕分ける('計算', 299, 31, [その行(24, 32)], byId, 表, wb) === null);
/* ★自己確認が 教えてくれた 抜け その2★
   ★ほかの表でも その行が 中身の 中に 在るなら Excel は 答えを 出せる＝出しては いけない★
   （自分の表と 重ならない 形で 試す＝列を ずらす） */
T('★ほかの表でも その行が 中身の 中なら 出さない（列をずらして 確かめる）★',
  (() => {
    const 長い = { id: 23, sheet: '計算', rw1: 233, rw2: 299, cl1: 0, cl2: 10, header: 1, totals: 1 };
    return TR._断りを仕分ける('計算', 268, 31, [その行(23, 5)], { 23: 長い, 24: 表[1] }, [表[1], 長い], wb) === null;
  })());
T('★別のシートの 表を 指していたら 出す（行が 合っていても 別の紙）★',
  (() => {
    const 別 = [{ id: 23, name: 'R8.8', sheet: '売上表', rw1: 233, rw2: 299, cl1: 0, cl2: 37, header: 1, totals: 1, cols: 列たち }, 表[1]];
    const x = TR._断りを仕分ける('計算', 268, 31, [その行(23, 32)], { 23: 別[0], 24: 表[1] }, 別, wb);
    return !!x && x.指した表 === 'R8.8';
  })());

/* ── ②-b ★.xlsx / .xlsm でも 同じ★（形式で 出たり 出なかったり させない）── */
{
  /* ★短い形（画面の 見え方）と 長い形（実Excel が 保存する 形）の 両方★ */
  const f = '=IFERROR(MAX(R8.9[@売上]*2, 1000*R8.8[@標本七ｈ]), 0)';
  const f長 = '=IFERROR(MAX(R8.9[[#This Row],[売上]]*2, 1000*R8.8[[#This Row],[標本七ｈ]]), 0)';
  const x = TR._断りを仕分ける文字('計算', { r: 268, c: 31 }, f, 表, 表[1]);
  T('★.xlsx の 文字の 式でも 見つける★', !!x);
  T('★.xlsx でも 同じ 直し方（R8.8 → R8.9）★',
    x && x.直し方.前 === 'R8.8' && x.直し方.後 === 'R8.9' && x.列名 === '標本七ｈ',
    x && JSON.stringify(x.直し方) + ' ' + x.列名);
  T('★.xlsx でも 種類と 場所は 同じ形★',
    x && x.種類 === 'thisrow_other_table' && x.セル === 'AF269');
  T('★.xlsx で 自分の表だけなら 出さない★',
    TR._断りを仕分ける文字('計算', { r: 268, c: 31 }, '=R8.9[@標本七ｈ]', 表, 表[1]) === null);
  T('★.xlsx で 知らない名前は 出さない（当て推量しない）★',
    TR._断りを仕分ける文字('計算', { r: 268, c: 31 }, '=しらない表[@標本七ｈ]', 表, 表[1]) === null);
  const x長 = TR._断りを仕分ける文字('計算', { r: 268, c: 31 }, f長, 表, 表[1]);
  T('★★実Excel が 保存する 長い形でも 見つける（[[#This Row],[列名]]）★★', !!x長,
    '★字で 探す 作りだと ここが 0件に なる（2026-09-04 実際に 踏んだ）★');
  T('★長い形でも 同じ 直し方★',
    x長 && x長.直し方.前 === 'R8.8' && x長.直し方.後 === 'R8.9' && x長.列名 === '標本七ｈ');
  T('★.xlsx で その行で ない（[列名]）物は 出さない★',
    TR._断りを仕分ける文字('計算', { r: 268, c: 31 }, '=SUM(R8.8[標本七ｈ])', 表, 表[1]) === null);
  /* ★自己確認が 教えてくれた 抜け★（xlsb 側と 同じ 2つを xlsx 側にも） */
  T('★.xlsx で 自分の表なら 合計の行から 呼んでも 出さない★',
    TR._断りを仕分ける文字('計算', { r: 299, c: 31 }, '=R8.9[@標本七ｈ]', 表, 表[1]) === null);
  T('★.xlsx で ほかの表でも その行が 中身の 中なら 出さない★',
    (() => {
      const 長い = { id: 23, name: 'R8.8', sheet: '計算', rw1: 233, rw2: 299, cl1: 0, cl2: 10, header: 1, totals: 1, cols: 列たち };
      return TR._断りを仕分ける文字('計算', { r: 268, c: 31 }, '=R8.8[@標本七ｈ]', [表[1], 長い], 表[1]) === null;
    })());
}

/* ── ②-c ★★実Excel が 作った 見本の ブックで 実際に 押す★★ ──
   ★作り物の 引数で 試すだけでは 足りない★（2026-09-04 実際に 踏んだ）：
   実Excel が 保存する 形は ★R8.8[[#This Row],[標本七ｈ]]★ で、
   ★[@標本七ｈ] の 短い形では ない★＝字で 探す 作りは ★1件も 見つけられなかった★。
   ⇒ ★実Excel に 作らせた 見本を repo に 置いて 毎回 通す★
      tools/make-mihon.ps1 で 作った（司さんの 実物は 1バイトも 触っていない）
      中身 … 表2つ（R8.8＝A1:C5 ／ R8.9＝A8:C12）、R8.9 の C列が R8.8 の その行を 指す
      ★Excel 自身も 答えを 出せず 値は 0★（エラーは 出ない＝客は 気づけない） */
{
  const fs = await import('node:fs');
  const XLSX = require_(path.join(ROOT, 'lib/xlsx.full.min.js'));
  const ZipSurgeon = require_(path.join(ROOT, 'lib/zip-surgeon.js'));
  const TR2 = require_(path.join(ROOT, 'lib/table-refs.js'));
  for (const [形, 名] of [['xlsx', 'hyou-no-soto-sample.xlsx'], ['xlsb', 'hyou-no-soto-sample.xlsb']]) {
    const 道 = path.join(ROOT, 'tests/fixtures', 名);
    if (!fs.existsSync(道)) { T('★見本が 在る（' + 名 + '）★', false, '★無い＝この 試験は 空振り★'); continue; }
    const bytes = new Uint8Array(fs.readFileSync(道));
    const wb = XLSX.read(bytes, { type: 'array', cellFormula: true });
    const rr = await TR2.resolve(bytes, 形, wb, ZipSurgeon);
    const 出 = (rr.断り || []).filter((x) => x.種類 === 'thisrow_other_table');
    T('★' + 形 + ' の 見本で 見つける（4か所）★', 出.length === 4, '見つけた ' + 出.length + '件');
    T('★' + 形 + ' の 直し方が 正しい（R8.8[@標本七ｈ] → R8.9[@標本七ｈ]）★',
      出.length > 0 && 出[0].直し方.前 === 'R8.8' && 出[0].直し方.後 === 'R8.9' && 出[0].列名 === '標本七ｈ',
      出.length ? (出[0].直し方.前 + '→' + 出[0].直し方.後 + ' [@' + 出[0].列名 + ']') : '(0件)');
    T('★' + 形 + ' の 場所が 正しい（計算!C9 から）★',
      出.length > 0 && 出[0].シート === '計算' && 出[0].セル === 'C9', 出.length ? 出[0].セル : '(0件)');
    T('★' + 形 + ' Excel 自身も 0 を 返している（エラーが 出ない＝客は 気づけない）★',
      (wb.Sheets['計算'].C9 || {}).v === 0, JSON.stringify((wb.Sheets['計算'].C9 || {}).v));
    /* ★★答えを 実Excel に 合わせる★★（2026-09-04＝★絵を 開いて 見つけた★）
       断ると 式ごと 読めず ★IFERROR まで 道連れで #ERROR★＝実Excel の 0 と 違う。
       ⇒★Excel と 同じ エラーの 印を 置く★＝IFERROR が 拾って 0 に なる。 */
    const 直し = rr.fixes['計算|8,2'] || '';
    T('★' + 形 + ' その セルを 直している（断って 放り出していない）★', !!直し, '(直していない)');
    T('★' + 形 + ' ★Excel と 同じく エラーの 印を 置く（IFERROR が 拾える）★★',
      直し.indexOf('#VALUE!') >= 0 && 直し.indexOf('IFERROR') >= 0, 直し.slice(0, 120));
    /* ★置き換えた 数 ＝ 診断に 出す 数★（1個でも 黙らせたら 赤） */
    let 置換 = 0;
    for (const k of Object.keys(rr.fixes || {})) {
      if (String(rr.fixes[k]).indexOf('#VALUE!') < 0) continue;
      const [sn, rc] = k.split('|'); const [r, c] = rc.split(',').map(Number);
      const も = (wb.Sheets[sn] && wb.Sheets[sn][XLSX.utils.encode_cell({ r, c })] || {}).f || '';
      if (('=' + も).indexOf('#VALUE!') >= 0) continue;
      置換++;
    }
    T('★' + 形 + ' ★置き換えた 数 ＝ 診断に 出す 数（黙らせない）★★',
      置換 === (rr.断り || []).length, '置き換え ' + 置換 + ' / 診断 ' + (rr.断り || []).length);
  }
}

/* ── ②-d ★★答えを 変えた 所は 必ず 診断に 出す（黙って 0 に しない）★★ ──
   ★指示役の 問い（2026-09-04）★
     「Excel と 同じ エラーの 印を 置く を 入れた 事で、
       今まで #ERROR だった 他の 式が 黙って 0 に 変わっていませんか」
   ⇒★1つ 在った★＝★表の 外から 自分の表の「同じ行」を 見ている 式★
     （直す 名前が 無いので 前は 診断に 出さなかった＝★黙って 0★）
   ⇒★型を 分けて 必ず 出す ようにした★
   ★見本★ tests/fixtures/hyou-no-soto-soto-sample.xlsx（tools/make-mihon-soto.ps1）
     計算!B20 ＝ 表(R8.8 は A1:C5)の 外の 行から R8.8 の「同じ行」を 見る
     ★Excel の 答えも 0★（エラーは 出ない） */
{
  const fs = await import('node:fs');
  const XLSX = require_(path.join(ROOT, 'lib/xlsx.full.min.js'));
  const ZipSurgeon = require_(path.join(ROOT, 'lib/zip-surgeon.js'));
  const TR3 = require_(path.join(ROOT, 'lib/table-refs.js'));
  const 道 = path.join(ROOT, 'tests/fixtures/hyou-no-soto-soto-sample.xlsx');
  if (!fs.existsSync(道)) {
    T('★見本が 在る（表の外から 見ている 物）★', false, '★無い＝この 試験は 空振り★');
  } else {
    const bytes = new Uint8Array(fs.readFileSync(道));
    const wb = XLSX.read(bytes, { type: 'array', cellFormula: true });
    const rr = await TR3.resolve(bytes, 'xlsx', wb, ZipSurgeon);
    const 出 = rr.断り || [];
    T('★★答えを 変えた 所が 診断に 出る（黙って 0 に しない）★★', 出.length === 1,
      '見つけた ' + 出.length + '件');
    T('★型が 分かれている（表の外から）★',
      出.length > 0 && 出[0].種類 === 'thisrow_outside_table', 出.length ? 出[0].種類 : '(0件)');
    T('★場所が 出る（計算!B20）★', 出.length > 0 && 出[0].セル === 'B20');
    T('★直す 名前が 無い事を そのまま 出す（前＝後の 嘘を 言わない）★',
      出.length > 0 && 出[0].直し方 === null);
    T('★一覧の 字が 出る★',
      出.length > 0 && Shindan.直し方の字(出[0]) === 'R8.8 の「同じ行」を 表の外から 見ている（標本七ｈ の 所）',
      出.length ? Shindan.直し方の字(出[0]) : '(0件)');
    T('★Excel と 同じ 答えに している（エラーの印を 置く）★',
      String(rr.fixes['計算|19,1'] || '').indexOf('#VALUE!') >= 0, rr.fixes['計算|19,1']);
    T('★Excel 自身も 0★', (wb.Sheets['計算'].B20 || {}).v === 0);
    const w = Shindan.言葉('thisrow_outside_table', 1);
    T('★この型の 言葉が 在る★', !!w && w.つぎ.indexOf('表の中へ 移す') >= 0, w && w.つぎ);
    T('★この型の 言葉にも 中の言葉を 出さない★',
      (w.題 + w.本文 + w.つぎ).indexOf('[@') < 0 && (w.題 + w.本文 + w.つぎ).indexOf('IFERROR') < 0
      && (w.題 + w.本文 + w.つぎ).indexOf('★') < 0);
    /* ★★どちらの 型も 数えて、置き換えた 数と 合うか★★（1個でも 黙っていたら 赤） */
    let 置換 = 0;
    for (const k of Object.keys(rr.fixes || {})) {
      if (String(rr.fixes[k]).indexOf('#VALUE!') < 0) continue;
      const [sn, rc] = k.split('|'); const [r, c] = rc.split(',').map(Number);
      const 元 = (wb.Sheets[sn] && wb.Sheets[sn][XLSX.utils.encode_cell({ r, c })] || {}).f || '';
      if (('=' + 元).indexOf('#VALUE!') >= 0) continue;
      置換++;
    }
    T('★★置き換えた 数 ＝ 診断に 出す 数（1個も 黙らせない）★★', 置換 === 出.length,
      '置き換え ' + 置換 + ' / 診断 ' + 出.length);
  }
}

/* ── ③ 客に 見せる 言葉 ── */
{
  const w = Shindan.言葉('thisrow_other_table', 62);
  const 客に出る字 = (w.題 + w.本文 + w.つぎ + Shindan.直し方の字({ 列名: 'あ', 直し方: { 前: 'A', 後: 'B' } }));
  T('★言葉が 在る★', !!w);
  T('★何か所か を 言う★', w && w.本文.indexOf('62か所') >= 0);
  T('★「0 のまま」と 言う（気づけない事を 言う）★', w && w.本文.indexOf('0 のまま') >= 0);
  T('★どこを 押すかを 言う（式の欄）★', w && w.つぎ.indexOf('式の欄') >= 0, w && w.つぎ);
  T('★何を 変えるかを 言う（表の名前だけ）★', w && w.つぎ.indexOf('表の名前だけ') >= 0);
  T('★1か所 直したら 下へ コピー、と 言う★', w && w.つぎ.indexOf('下へコピー') >= 0);
  T('★金額らしき数を 出さない★', w && (w.本文 + w.つぎ).match(/\d{6,}/) === null);
  /* ★★2026-09-04（司さん「分かりやすく 説明しとんか？」）★★
     ★客に 中の言葉を 出さない★＝会社の 決まり。前の 文には 入っていた。 */
  for (const 中の言葉 of ['[@', 'IFERROR', 'IFNA', '#REF!', '#VALUE!', 'その行', 'band', 'thisRow']) {
    T('★客に 中の言葉を 出さない（' + 中の言葉 + '）★', 客に出る字.indexOf(中の言葉) < 0,
      '「' + 中の言葉 + '」が 客の 文に 出ている');
  }
  T('★客に見せる字に ★ を書かない（★は うちの覚え書きの印）★', 客に出る字.indexOf('★') < 0, 客に出る字);
}

/* ── ④ 画面に つないである（台帳に 書いただけで 出ていない、を 防ぐ）── */
{
  const fs = await import('node:fs');
  const html = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
  T('★画面が 断りを 受け取っている★', html.indexOf('res.表の断り') > 0);
  T('★画面が 種類で 選り分けている★', html.indexOf("'thisrow_other_table'") > 0);
  T('★画面が 直し方の字を 出している★', html.indexOf('Shindan.直し方の字') > 0);
  T('★ボタンの 数が 2本の 合計★', html.indexOf("'危ない所 ' + (n1 + n2) + 'か所'") > 0);
  const open = fs.readFileSync(path.join(ROOT, 'js/book-open.js'), 'utf8');
  T('★読み込みが 断りを 渡している★', open.indexOf('表の断り: tr断り') > 0);
  T('★前のファイルの物を 残さない（毎回 入れ替える）★',
    html.indexOf('_hyouNoSoto = (res.表の断り') > 0);
}

/* ── ⑤ ★機械が 見つけた 物を AIへ 渡しているか★（2026-09-04 司さん「なんで できてないんど」）── */
{
  const fs = await import('node:fs');
  const Chizu = require_(path.join(ROOT, 'lib/chizu.js'));
  const Horu = require_(path.join(ROOT, 'lib/horu.js'));
  const AiGenkai = require_(path.join(ROOT, 'lib/ai-genkai.js'));
  const 断り = [{ 種類: 'thisrow_other_table', シート: '計算', r: 268, c: 31, セル: 'AF269',
    指した表: 'R8.8', 自分の表: 'R8.9', 列名: '標本七ｈ', 直し方: { 前: 'R8.8', 後: 'R8.9' } }];
  const m = Chizu.作る([{ name: '計算', data: { '0,0': { v: 1 } } }], null, null, { 表の断り: 断り });
  T('★地図に 2本目が 載る（見つけたのに 渡さない、を 防ぐ）★',
    m.字.indexOf('機械が見つけた 危ない所（AIは呼んでいない）その2') > 0
    && m.字.indexOf('ほかの表の「その行」を見ている式') > 0, m.字.slice(-260));
  /* ★地図は AI 向け★＝中の 書き方で よい（客に 出るのは 窓の 字）。★混ぜない★ */
  T('★地図に 直し方が 載る（前→後・AI向け）★', m.字.indexOf('R8.8[@標本七ｈ]→R8.9[@標本七ｈ]') > 0);
  T('★地図に 何か所かが 載る★', m.字.indexOf('1か所') > 0);
  T('★1つも 無い時は 地図に 出さない（0件で 騒がない）★',
    Chizu.作る([{ name: '計算', data: { '0,0': { v: 1 } } }], null, null, {}).字.indexOf('その2') < 0);
  const html = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
  T('★画面が 地図に 2本目を 渡している★', html.indexOf('表の断り: (typeof _hyouNoSoto') > 0);

  /* ★掘れる回数＝ブックを 掘りきれる 回数★（★5回の 決め打ちに 戻らない為★） */
  const 作る = (n) => { const d = {}; for (let i = 0; i < n; i++) d[i + ',0'] = { v: 1 }; return [{ name: 's', data: d }]; };
  T('★小さいブックでも 5回は 掘れる★', Horu.掘れる回数を決める(作る(10)) === 5);
  /* ★2026-09-05★ ここは ★司さん 08-22 の「何度でも掘れる＝実質 全部見られる」★を
     守っていた 見張りです。★弱めていません。事実が 変わったので 書き直しました★。
     ★何が あったか★ … サーバは 1分に 10回で 止める（api/claude.js 事故止め・指示役 08-25）。
       画面は 実物ブックで ★15回★ 叩こうとしていた ⇒★11回目で 止まり 掘りを 全部 捨てていた★
       ⇒ 数を lib/ai-genkai.js の 1か所に 集め、★叩ける 8回★に 頭打ちした。
     ★その結果 21,204セルは 掘りきれません（6回＝約12,000セル）★＝★後退です★。
     ★隠さない★＝下の 3本目で ★客に そう 言える事★を 見張ります。
     ★戻す 手（材料を 足し続けない／掘った分を 使い回す）は 指示役が 司さんへ★。 */
  T('★上限に 当たるまでは 大きいほど 多く 掘れる（決め打ちに なっていない）★',
    Horu.掘れる回数を決める(作る(10000)) > Horu.掘れる回数を決める(作る(2000)),
    Horu.掘れる回数を決める(作る(2000)) + ' → ' + Horu.掘れる回数を決める(作る(10000)));
  T('★叩ける 回数を 超えない（超えると 掘りを 全部 捨てる）★',
    Horu.掘れる回数を決める(作る(21204)) + 2 <= AiGenkai.限界.分の回数,
    '掘れる' + Horu.掘れる回数を決める(作る(21204)) + '＋2 vs 止める' + AiGenkai.限界.分の回数);
  /* ★★上限を 上げたら 掘れる回数も 増える★★（2026-09-05 指示役の 問い）
     ★なぜ 要るか★ … 上の 2本は「掘れる ≦ 上限」しか 見ていない＝★上限を 下げるほど 通る★。
     ★それでは 5回の 決め打ちに 戻しても 気づけない★（＝④と 同じ 死に方）。
     ⇒★頭打ちが 外れたら ★ブックの 大きさに 応じて 増える★事★を 見張る。
     ⇒★司さん 08-22「何度でも掘れる」へ 戻る 道が 生きている事★の 見張り。 */
  (function 戻る道が生きているか() {
    var 元 = AiGenkai.限界.一問で叩ける;
    try {
      AiGenkai.限界.一問で叩ける = 30;                 /* 上限を 上げてみる */
      var 増えた = Horu.掘れる回数を決める(作る(21204));
      AiGenkai.限界.一問で叩ける = 元;
      T('★上限を 上げたら 掘れる回数も 増える（決め打ちに なっていない）★',
        増えた > Horu.掘れる回数を決める(作る(21204)) && 増えた * 2000 >= 21204,
        '上限8→' + Horu.掘れる回数を決める(作る(21204)) + '回 ／ 上限30→' + 増えた + '回');
    } finally { AiGenkai.限界.一問で叩ける = 元; }
  })();

  T('★★掘りきれない時は 客に そう 言う（黙って「全部見た」顔を しない）★★',
    !AiGenkai.掘りきれるか(AiGenkai.要る回数(21204, 2000))
    && typeof Horu.掘りきれない時の客への言葉 === 'function'
    && html.indexOf('掘りきれない時の客への言葉') > 0,
    '実物21,204セル → 掘れる' + Horu.掘れる回数を決める(作る(21204)) + '回（★全部は 見られない★）');
  T('★画面が その回数を 使っている（12回の 決め打ちで ない）★',
    html.indexOf('Horu.掘れる回数を決める(sheets)') > 0 && html.indexOf('回 < 掘れる + 2') > 0);
}

/* ── わざと 壊して 赤に なるか ── */
/* ★「出さない」を もう一度 呼ぶだけでは 自己確認に ならない★（同じ物を 2回 見るだけ）
   ⇒ ★本物の 部品を わざと 壊して、この 試験が 赤に なるか を 見る★ */
if (process.argv.includes('--self-test')) {
  console.log('');
  console.log('★本物の 部品を わざと 壊して 赤に なるか★');
  const fs = await import('node:fs');
  const { execFileSync } = await import('node:child_process');
  const 道 = path.join(ROOT, 'lib/table-refs.js');
  const 元 = fs.readFileSync(道, 'utf8');
  const 壊す = [
    ['★自分の表かどうかを 見なくする★',
      (t) => t.replace('      if (L.id === 自分.id) continue;', '      if (false) continue;')],
    ['★その行（[@列名]）以外も 拾う★',
      (t) => t.replace("      if (BAND_BY_CODE[L.code] !== 'thisRow') continue;", '      if (false) continue;')],
    ['★中身の 行の 中か どうかを 見なくする★',
      (t) => t.replace('      if (先.sheet === sn && 行 >= 中身の頭 && 行 <= 中身の尻) continue;', '      if (false) continue;')],
    ['★直し方（後）を 取り違える★',
      (t) => t.replace('後: 自分の名 },', '後: 先の名 },')],
    ['★番号の 名前（Table23）で 助言してしまう★',
      (t) => t.replace("var 先の名 = (先.name || ('Table' + L.id));", "var 先の名 = 'Table' + L.id;")],
    ['★.xlsx の 側で 自分の表かどうかを 見なくする★',
      (t) => t.replace('      if (名 === own.name) continue;', '      if (false) continue;')],
    ['★.xlsx の 側で 中身の 行か どうかを 見なくする★',
      (t) => t.replace('      if (先.sheet === sn && rc.r >= 中身の頭 && rc.r <= 中身の尻) continue;', '      if (false) continue;')],
    ['★列名を 出さなくする★',
      (t) => t.replace('列名 = String(hc.v);', '列名 = null;')],
  ];
  /* ★lib/horu.js と lib/chizu.js の 側も 壊して 見る★（渡す所が 抜けたら 気づけない） */
  const 別の壊し = [
    /* ★★本当に 動いている 道を 壊す★★（2026-09-04）
       ＝実Excel の 見本を 通す 道。上の 8通りは ★断った 時の 受け皿★を 壊している。 */
    ['lib/table-refs.js', '★実物の 道＝控えるのを やめる（答えは 合うが 誰も 気づけない）★',
      (t) => t.replace('    ctx.断り.push(元);', '    if (false) ctx.断り.push(元);')],
    ['lib/table-refs.js', '★表の外から の 型を 黙らせる（前の 作りに 戻す）★',
      (t) => t.replace("      元.種類 = 'thisrow_outside_table';", '      return;')],
    ['lib/table-refs.js', '★直す 名前が 無いのに 在る事に する（前＝後の 嘘）★',
      (t) => t.replace('      元.直し方 = null;', "      元.直し方 = { 前: 先の名, 後: 先の名 };")],
    ['lib/table-refs.js', '★実物の 道＝その行が 中か 外かを 見なくする★',
      (t) => t.replace('    return ctx.row >= (def.rw1 + def.header) && ctx.row <= (def.rw2 - def.totals);', '    return true;')],
    ['lib/table-refs.js', '★実物の 道＝Excel と 同じ 答えに しない（#ERROR に 戻す）★',
      (t) => t.replace("        out += formula.slice(prev, h.at) + '#VALUE!';", '        return null;')],
    /* ★2026-09-05★ 掘る数を lib/ai-genkai.js で 頭打ちする ようにしたので、
       ★掴む 字が 変わった★（前の 字は もう 無い＝★素通り★に なっていた）。
       ⇒ 今の 実物に 合わせる。★見張りを 減らしていない★ */
    ['lib/horu.js', '★掘れる回数を 5回の 決め打ちに 戻す★',
      (t) => t.replace('var 欲しい = Math.max(下限, Math.ceil(全部 / Math.max(1, 一度のセル)) + 2);', 'var 欲しい = 下限;')],
    ['lib/chizu.js', '★機械が 見つけた 2本目を 地図に 載せない★',
      (t) => t.replace("      行.push('## 機械が見つけた 危ない所（AIは呼んでいない）その2');", '')],
    ['lib/chizu.js', '★地図に 直し方を 載せない★',
      (t) => t.replace("      行.push('- 例: ' + 直し.join(' / ')", "      行.push('- 例: ' + ''")],
  ];
  for (const [名, f] of 壊す) {
    const 壊れ = f(元);
    if (壊れ === 元) { console.log('  ★素通り★  ' + 名 + '（印が 古い＝直せ）'); fail++; continue; }
    fs.writeFileSync(道, 壊れ);
    let 赤 = false;
    try { execFileSync(process.execPath, [path.join(ROOT, 'tests', 'hyou-no-soto.test.mjs')], { stdio: 'pipe' }); }
    catch (e) { 赤 = true; }
    fs.writeFileSync(道, 元);                 /* ★必ず 戻す★ */
    console.log((赤 ? '  赤くなった  ' : '  ★素通り★  ') + 名);
    if (!赤) fail++;
  }
  T('★本物は 壊していない（戻した）★', fs.readFileSync(道, 'utf8') === 元);
  for (const [部品, 名, f] of 別の壊し) {
    const み = path.join(ROOT, 部品);
    const もと = fs.readFileSync(み, 'utf8');
    const こわれ = f(もと);
    if (こわれ === もと) { console.log('  ★素通り★  ' + 名 + '（印が 古い＝直せ）'); fail++; continue; }
    fs.writeFileSync(み, こわれ);
    let 赤 = false;
    try { execFileSync(process.execPath, [path.join(ROOT, 'tests', 'hyou-no-soto.test.mjs')], { stdio: 'pipe' }); }
    catch (e) { 赤 = true; }
    fs.writeFileSync(み, もと);                 /* ★必ず 戻す★ */
    console.log((赤 ? '  赤くなった  ' : '  ★素通り★  ') + 名 + '（' + 部品 + '）');
    if (!赤) fail++;
    if (fs.readFileSync(み, 'utf8') !== もと) { console.log('  ★戻せていない★ ' + 部品); fail++; }
  }
  T('★言葉を 知らない 種類で 呼んだら null（黙って 何か 出さない）★',
    Shindan.言葉('しらない種類', 1) === null);
  T('★直し方の 字は 中身が 無ければ 空★', Shindan.直し方の字(null) === '');
}

console.log('');
console.log('hyou-no-soto: ' + pass + ' 緑 / ' + fail + ' 赤');
process.exit(fail ? 1 : 0);
