/* formula-nokori.test.mjs — ★答えは 全部 実Excel の 実測★（2026-09-06）
 *
 *  ★司さん★「きっちり ファイルに するのと ★細胞レベルで 対応できるように しろ★」
 *
 *  ★★この 試験の 決まり★★
 *    ・期待値は ★1つも 私が 考えていない★
 *    ・全部 `docs/measured/golden-2026-09-06.tsv` に 在る
 *      ＝★実Excel 16.0 build 20326（日本語UI 1041）に COM で 打たせて 読んだ 字★
 *    ・★紙の 数と 試した 数が 合っているか★も 見る（拾い漏らしを 止める）
 *
 *  使い方: node tests/formula-nokori.test.mjs
 *          node tests/formula-nokori.test.mjs --self-test
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(import.meta.url);
const F = require_(path.join(ROOT, 'lib/formula-nokori.js'));
let pass = 0, fail = 0;
const ok = (名, 条件, 添え) => {
  if (条件) { pass++; console.log('  ok   ' + 名); }
  else { fail++; console.log('  ★NG★ ' + 名 + (添え !== undefined ? '  … ' + 添え : '')); }
};
const 同じ = (名, 出, 期待) => ok(名 + '  ＝ ' + JSON.stringify(期待), String(出) === String(期待), '今 ' + JSON.stringify(出));

/* ══ 実Excel の 答えを 読む（★repo に 在る★） ══════════════ */
const 金 = fs.readFileSync(path.join(ROOT, 'docs/measured/golden-2026-09-06.tsv'), 'utf8')
  .split(/\r?\n/).filter(Boolean).slice(1)
  .filter((l) => !l.startsWith('#'))                 /* ★# の行は 答えでは ない★ */
  .map((l) => { const a = l.split('\t'); return { 関数: a[0], 式: a[1], 答: a[2] }; });

/* ★材料は 紙から 読む★＝ここで 手で 書かない（写し間違いを 無くす）
   ★前は ここに 手で 書いて 在りました★ … const A = [[1],[4],[3],[2]] など。
   ⇒ 同じ 材料を 別の 道具が もう一度 写す事に なり、★写し間違えれば 紙は 緑・本番は 赤★。
   ⇒ 2026-09-14 に `docs/measured/golden-2026-09-06.tsv` へ 移した。★紙が 正本★。 */
const 材料 = {};
for (const l of fs.readFileSync(path.join(ROOT, 'docs/measured/golden-2026-09-06.tsv'), 'utf8').split(/\r?\n/)) {
  if (!l.startsWith('#材料')) continue;
  const a = l.split('\t');
  if (a.length >= 3) 材料[a[1].trim()] = Number(a[2]);
}
const 四角 = (左上, 右下) => {
  const m = /^([A-Z])(\d+)$/.exec(左上), n = /^([A-Z])(\d+)$/.exec(右下);
  const 出 = [];
  for (let r = +m[2]; r <= +n[2]; r++) {
    const 段 = [];
    for (let c = m[1].charCodeAt(0); c <= n[1].charCodeAt(0); c++) {
      const 名 = String.fromCharCode(c) + r;
      if (!(名 in 材料)) throw new Error('★紙に 材料が 無い★: ' + 名);
      段.push(材料[名]);
    }
    出.push(段);
  }
  return 出;
};
const 答 = (式) => { const r = 金.find((x) => x.式 === 式); if (!r) throw new Error('★紙に その式が 無い★: ' + 式); return r.答; };

if (process.argv.includes('--self-test')) {
  console.log('\n[formula-nokori --self-test] わざと壊して赤になるか');
  ok('★紙が 空でない（空を 緑に しない）', 金.length >= 30, 金.length + '行');
  ok('★紙に 無い 式を 聞いたら 落ちる', (() => {
    try { 答('=NAI()'); return false; } catch (e) { return true; }
  })());
  ok('★バイト数＝全角は 2（実Excel と 同じ数え方）', F.バイト数('あいう') === 6 && F.バイト数('abc') === 3);
  ok('★逆行列が 無い時は #VALUE!', (F.逆行列([[1, 2], [2, 4]]) || {}).誤り === 'VALUE');
  console.log('\n' + pass + ' passed, ' + fail + ' failed');
  process.exit(fail ? 1 : 0);
}

console.log('\n[formula-nokori] ★答えは 全部 実Excel の 実測★（' + 金.length + '通り）');

console.log('\n[① 字の 幅で 数える（B が 付く 物）]');
同じ('FINDB("b","abc")', F.探すB('b', 'abc'), 答('=FINDB("b","abc")'));
同じ('FINDB("い","あいう")', F.探すB('い', 'あいう'), 答('=FINDB("い","あいう")'));
同じ('SEARCHB("B","abc")', F.探すB大小なし('B', 'abc'), 答('=SEARCHB("B","abc")'));
同じ('SEARCHB("い","あいう")', F.探すB大小なし('い', 'あいう'), 答('=SEARCHB("い","あいう")'));
同じ('REPLACEB("abcdef",2,3,"X")', F.入れ替えB('abcdef', 2, 3, 'X'), 答('=REPLACEB("abcdef",2,3,"X")'));
同じ('REPLACEB("あいう",3,2,"X")', F.入れ替えB('あいう', 3, 2, 'X'), 答('=REPLACEB("あいう",3,2,"X")'));

console.log('\n[② 区切って 分ける]');
同じ('TEXTSPLIT の 1つ目', F.区切って分ける('a,b,c', ',')[0][0], 答('=TEXTSPLIT("a,b,c",",")'));
同じ('TEXTSPLIT の 2つ目', F.区切って分ける('a,b,c', ',')[0][1], 答('=INDEX(TEXTSPLIT("a,b,c",","),1,2)'));
ok('★形を 保つ（横 3つ）★', F.区切って分ける('a,b,c', ',')[0].length === 3);

console.log('\n[③ 正規表現の 3つ]');
同じ('REGEXTEST 当たる', F.正規で調べる('abc123', '[0-9]+') ? 'TRUE' : 'FALSE', 答('=REGEXTEST("abc123","[0-9]+")'));
同じ('REGEXTEST 当たらない', F.正規で調べる('abc', '[0-9]+') ? 'TRUE' : 'FALSE', 答('=REGEXTEST("abc","[0-9]+")'));
同じ('REGEXEXTRACT', F.正規で取り出す('abc123', '[0-9]+'), 答('=REGEXEXTRACT("abc123","[0-9]+")'));
同じ('REGEXREPLACE', F.正規で入れ替える('abc123', '[0-9]+', '#'), 答('=REGEXREPLACE("abc123","[0-9]+","#")'));

console.log('\n[④ 別の 列で 並べる]');
const A = 四角('A1', 'A4'), B = 四角('B1', 'B4');   /* ★紙から 読む★（手で 書かない） */
同じ('SORTBY 小さい順の 1つ目', F.別の列で並べる(A, B, 1)[0][0], 答('=INDEX(SORTBY(A1:A4,B1:B4,1),1,1)'));
同じ('SORTBY 大きい順の 1つ目', F.別の列で並べる(A, B, -1)[0][0], 答('=INDEX(SORTBY(A1:A4,B1:B4,-1),1,1)'));

console.log('\n[⑤ 行列]');
同じ('MUNIT(3) の (1,1)', F.単位行列(3)[0][0], 答('=INDEX(MUNIT(3),1,1)'));
同じ('MUNIT(3) の (1,2)', F.単位行列(3)[0][1], 答('=INDEX(MUNIT(3),1,2)'));
同じ('MINVERSE の (1,1)', F.逆行列(四角('D1', 'E2'))[0][0], 答('=INDEX(MINVERSE(D1:E2),1,1)'));
同じ('MINVERSE の (1,2)', F.逆行列(四角('D1', 'E2'))[0][1], 答('=INDEX(MINVERSE(D1:E2),1,2)'));

console.log('\n[⑥ 順位の 割合]');
同じ('PERCENTRANK.INC', F.順位の割合(A, 3, undefined, true), 答('=PERCENTRANK.INC(A1:A4,3)'));
同じ('PERCENTRANK.EXC', F.順位の割合(A, 3, undefined, false), 答('=PERCENTRANK.EXC(A1:A4,3)'));
ok('★3桁で ★切り捨て★（四捨五入では ない）', String(F.順位の割合(A, 3, undefined, true)) === '0.666');

console.log('\n[⑦ 確率]');
同じ('PROB(値,確率,2,3)', F.確率(A, 四角('F1', 'F4'), 2, 3), 答('=PROB(A1:A4,F1:F4,2,3)'));

console.log('\n[⑧ かたまりの 数]');
同じ('AREAS 1つ', F.かたまりの数([1]), 答('=AREAS(A1:B3)'));
同じ('AREAS 2つ', F.かたまりの数([1, 2]), 答('=AREAS((A1:B3,D1:D2))'));

console.log('\n[⑨ 誤りの 番号]');
同じ('ERROR.TYPE(1/0)', F.誤りを番号に('DIV_BY_ZERO'), 答('=ERROR.TYPE(1/0)'));
同じ('ERROR.TYPE(NA())', F.誤りを番号に('NA'), 答('=ERROR.TYPE(NA())'));
ok('★誤りで ない 物は #N/A★', (F.誤りを番号に('') || {}).誤り === 'NA');

console.log('\n[⑩ 台帳と 紙が 揃っている]');
const 台 = F.数える();
ok('★台帳が 空でない★', 台.足す.length >= 15, 台.足す.length + '個');
const 紙の関数 = [...new Set(金.map((x) => x.関数))];
const 足りない = 台.足す.filter((f) => 紙の関数.indexOf(f) < 0);
ok('★台帳の 全部に 実Excel の 答えが 在る（考えた 値が 混ざらない）★',
  足りない.length === 0, 足りない.join(' / '));

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
