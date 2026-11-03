/* shiki-hyou-hidzuke.test.mjs — ★字の 日付・時刻を 数に する★（2026-09-15）
 *
 *  ★★なぜ 要るか★★
 *    2026-09-15、自前の 土台と 借り物に 同じ 式を 押させて 突き合わせたら
 *    ★借り物だけが 正しく 答える 10通り★が 出ました。その うち ★3本★が これです：
 *      `=AVERAGE("12:30")` ／ `=MIN("12:30")` ／ `=PRODUCT("12:30")`
 *      紙（実Excel）★0.5208333333333333★ ／ うち ★#VALUE!★
 *    ⇒★実Excel は 字の 時刻・日付を 数に します★。うちは していませんでした。
 *
 *  ★土台は 前から 口を 開けて 待って いました★
 *    `lib/shiki-keisan.js:139`「★日付の 字（"2026/1/1"→46023）は この 台では 読みません★」
 *    ＝`手.字を日付に` という 口。★誰も 渡して いなかった★だけ。
 *    ⇒ `lib/shiki-hyou.js` が ★既定の 手★として 渡す ように した。
 *
 *  ★★この 見張りが 見る 物★★
 *    ①実測の 値（"12:30" → 0.5208…／"2024-01-15" → 45306）
 *    ②★読まない 形は 読まない★（当て推量で 広げて いないか）
 *    ③★通し番号の 出し方が 2か所に 在る★＝`lib/formula-soto.js` の `日から数` と
 *      ★同じ 数を 返すか★（ずれたら 赤）
 *
 *  ★★見ていない 範囲★★
 *    ・`2024-01-15 12:30` の ような ★日付と 時刻の 組み合わせ★＝★測って いません★（読みません）
 *    ・`Jan 15, 2024` の ような 英語の 書き方／和暦／`1/15`（年なし）＝読みません
 *    ★★⇒ ただし 下の 3つは「読んで しまって いる」のに 未測定です★★（2026-09-15 に 自分で 押して 見つけた）
 *      `"24:00"`      … うち ★1★（1日 ちょうど）… 実Excel も 1 だと 思うが ★紙に 行が 無い★
 *      `"0000-01-01"` … うち ★3★           … ★実Excel は 断ると 思うが 紙に 行が 無い★
 *      `"2024-1-5"`   … うち ★45296★       … 月日が 1桁（★これは 読む 形に 入れて 在る★）
 *    ⇒★実Excel を 打てる 時に 測る★（棚 … `docs/measured/karimono-hazushi-no-tana.md`）
 *    ⇒★「たぶん こう」で 直しません★＝★今 直すと 当て推量に なる★
 *
 *  使い方: node tests/shiki-hyou-hidzuke.test.mjs
 *          node tests/shiki-hyou-hidzuke.test.mjs --self-test
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const H = require_(path.join(ROOT, 'lib/shiki-hyou.js'));
const SOTO = require_(path.join(ROOT, 'lib/formula-soto.js'));

let pass = 0, fail = 0;
const T = (n, f) => {
  try { f(); pass++; console.log('  ✓ ' + n); }
  catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); }
};
const 近い = (a, b) => Math.abs(Number(a) - Number(b)) <= Math.max(1e-9, Math.abs(Number(b)) * 1e-9);

console.log('\n[shiki-hyou-hidzuke] ★字の 日付・時刻を 数に する★（実Excel の 実測）');

/* ★紙から 引く★（手で 写さない） */
const 紙 = path.join(ROOT, 'docs/measured/kansuu46/golden-346-2026-09-08.tsv');
const 中 = fs.readFileSync(紙, 'utf8').split(/\r?\n/);
const 引く = (式) => {
  for (const l of 中) {
    if (l.startsWith('#')) continue;
    const c = l.split('\t');
    if ((c[1] || '').trim() === 式) return (c[2] || '').trim();
  }
  throw new Error('★紙に その 式が 無い★: ' + 式);
};

T('★紙から 引けて いる★（紙を 見ずに 緑に しない）', () => {
  const v = 引く('=ABS("12:30")');
  if (!近い(v, 0.5208333333333333)) throw new Error('紙の 値が 違う … ' + v);
});

/* ★押すのは「うちが 持って いる 関数」だけ★
   （`=ABS("12:30")` も 紙に 在りますが ★ABS は まだ 書いて いません★＝#NAME? に なる。
     ★まだ 無い 関数で 突き合わせると「日付が 読めない」と 読み違えます★） */
for (const 式 of ['=AVERAGE("12:30")', '=MIN("12:30")', '=PRODUCT("12:30")']) {
  T('★実Excel と 同じ★  ' + 式, () => {
    const 正 = 引く(式);
    const h = H.表();
    h.打つ('Z1', 式);
    const 出 = h.字('Z1');
    if (!近い(出, 正)) throw new Error('紙 `' + 正 + '` ／ うち `' + 出 + '`');
  });
}

T('★★借り物だけが 答えて いた 3本★★（=AVERAGE/MIN/PRODUCT の "12:30"）', () => {
  const h = H.表();
  for (const 式 of ['=AVERAGE("12:30")', '=MIN("12:30")', '=PRODUCT("12:30")']) {
    h.打つ('Z1', 式);
    const 出 = h.字('Z1');
    if (!近い(出, 0.5208333333333333)) throw new Error(式 + ' … うち `' + 出 + '`');
  }
});

T('★★読まない 形は 読まない★★（当て推量で 広げて いない）', () => {
  const h = H.表();
  for (const 式 of ['=SUM("あ")', '=SUM("25:99")', '=SUM("12:60")', '=SUM("2024-02-30")',
    '=SUM("2024-13-01")', '=SUM("Jan 15, 2024")', '=SUM("2024-01-15 12:30")']) {
    h.打つ('Z1', 式);
    const 出 = h.字('Z1');
    if (出 !== '#VALUE!') throw new Error('★読んで しまった★ ' + 式 + ' … `' + 出 + '`');
  }
});

T('★★通し番号の 出し方が 2か所で 同じ★★（lib/formula-soto.js の 日から数 と）', () => {
  const h = H.表();
  for (const [y, m, d] of [[2024, 1, 15], [1900, 1, 1], [1900, 3, 1], [1999, 12, 31],
    [2026, 1, 1], [2000, 2, 29], [2100, 12, 31]]) {
    const 字 = y + '-' + String(m).padStart(2, '0') + '-' + String(d).padStart(2, '0');
    h.打つ('Z1', '=SUM("' + 字 + '")');
    const うち = Number(h.字('Z1'));
    const 向こう = SOTO.日から数(y, m, d);
    if (うち !== 向こう) {
      throw new Error('★ずれた★ ' + 字 + ' … 土台 ' + うち + ' ／ formula-soto ' + 向こう);
    }
  }
});

T('★★通し番号は 3か所とも 同じ★★（shiki-hyou ／ formula-soto ／ formula-kane）', () => {
  const KANE = require_(path.join(ROOT, 'lib/formula-kane.js'));
  const h = H.表();
  /* ★★　1900年の 日を 入れて いませんでした★★（2026-09-16）
       ⇒★ 4か所 全部が 同じだけ 違って いても 緑に なって いた★
       ⇒★★「揃って いる」と 「合って いる」は 別★★＝実Excel の 数とも 比べます */
  for (const [y, m, d] of [[2024, 1, 15], [2026, 1, 1], [2000, 2, 29], [1999, 12, 31],
    [1900, 1, 1], [1900, 2, 28], [1900, 3, 1]]) {
    const 字 = y + '-' + String(m).padStart(2, '0') + '-' + String(d).padStart(2, '0');
    h.打つ('Z1', '=SUM("' + 字 + '")');
    const 三 = [Number(h.字('Z1')), SOTO.日から数(y, m, d), KANE.日から数(y, m, d)];
    if (三[0] !== 三[1] || 三[1] !== 三[2]) {
      throw new Error('★ずれた★ ' + 字 + ' … 土台 ' + 三[0] + ' ／ soto ' + 三[1] + ' ／ kane ' + 三[2]);
    }
  }
});

T('★★実Excel の 数と 合う★★（4か所とも・紙から 引く）', () => {
  const KANE = require_(path.join(ROOT, 'lib/formula-kane.js'));
  /* ★物差し★ `docs/measured/golden-1900-hidzuke-2026-09-15.tsv`
       `=DATE(1900,1,1)`→★1★ `=DATE(1900,2,28)`→★59★ `=DATE(1900,3,1)`→★61★ */
  const 紙 = path.join(ROOT, 'docs/measured/golden-1900-hidzuke-2026-09-15.tsv');
  const 行 = fs.readFileSync(紙, 'utf8').split(/\r?\n/);
  const 引 = (式) => {
    for (const l of 行) {
      if (l.startsWith('#')) continue;
      const c = l.split('\t');
      if ((c[0] || '').trim() === 式) return (c[1] || '').trim();
    }
    throw new Error('★紙に その 式が 無い★: ' + 式);
  };
  const h = H.表();
  for (const [y, m, d] of [[1900, 1, 1], [1900, 2, 28], [1900, 3, 1]]) {
    const 正 = Number(引('=DATE(' + y + ',' + m + ',' + d + ')'));
    const 字 = y + '-' + String(m).padStart(2, '0') + '-' + String(d).padStart(2, '0');
    h.打つ('Z1', '=SUM("' + 字 + '")');
    const 四 = [Number(h.字('Z1')), SOTO.日から数(y, m, d), KANE.日から数(y, m, d)];
    for (const v of 四) {
      if (v !== 正) throw new Error('★実Excel と 違う★ ' + 字 + ' … 紙 ' + 正 + ' ／ うち ' + 四.join('/'));
    }
  }
});

T('★★現代の 日付は 往復する★★（数 → 日 → 数 で 戻る）', () => {
  const KANE = require_(path.join(ROOT, 'lib/formula-kane.js'));
  for (const n of [45292, 45306, 46023, 40000, 30000, 61]) {
    const p = KANE.数から日(n);
    const 戻 = KANE.日から数(p.y, p.m, p.d);
    if (戻 !== n) throw new Error('★戻らない★ ' + n + ' → ' + p.y + '-' + p.m + '-' + p.d + ' → ' + 戻);
  }
});

/* ★★1900年の 辺りは 往復しません＝実測済み（2026-09-15）★★
   数から日(1)=1900-01-01 ／ ★日から数(1900,1,1)=3★ … ★1〜60 の 全部で 戻りません★
   数から日(60) と 数から日(61) が ★同じ 日★＝★60 に 当たる 日が 無い★
   ★元★＝`日から数` は n<61 で +1 するが `数から日` は 引いて いない
   ★どちらが 正しいかは 実Excel を 打つまで 分かりません★＝★ここでは 押しません★
   ★でも「うちの 中で 矛盾して いる」事は 書いて 残します★（棚 ⑨） */
T('★1900年の 往復が 合わない事が 棚に 書いて 在る★（黙って 通さない）', () => {
  const 棚 = fs.readFileSync(path.join(ROOT, 'docs/measured/karimono-hazushi-no-tana.md'), 'utf8');
  if (!/往復（数→日→数）が 合って いません/.test(棚)) throw new Error('★棚に 書いて いない★');
});

T('★マスに 打った 字は 今まで どおり★（数に 化けて いない）', () => {
  const h = H.表();
  h.打つ('A1', '12:30');
  const v = h.値('A1');
  if (!v || v.型 !== '字') throw new Error('★マスの 中身が 字で なく なった★ … ' + JSON.stringify(v));
});

if (process.argv.includes('--self-test')) {
  console.log('\n[shiki-hyou-hidzuke --self-test] ★わざと 壊したら 赤に なるか★');
  T('★紙に 無い 式を 引いたら 落ちる★', () => {
    let 落ちた = false;
    try { 引く('=NAIshiki("12:30")'); } catch (e) { 落ちた = true; }
    if (!落ちた) throw new Error('落ちなかった');
  });
  /* ★★1900年の 辺りは「未測定」です★★
     私は 一度 ここで「1900-01-01 は 1 の はず」と ★紙を 見ずに 書きました★。
     ⇒ 実際は うちは ★3★ を 返します（`lib/formula-soto.js` の `日から数` も 同じ 3）。
     ⇒ ★紙に 1900年の 行は 1本も 在りません★＝★どちらが 正しいか 分かりません★。
     ⇒★押しません★。★2か所が 同じ 数を 返す事だけ★を 上の 見張りで 見ます。
       （実Excel を 打てる 時に 測る＝棚） */
  T('★1900年の 辺りは 未測定と 書いて 在る★（当て推量で 押さない）', () => {
    const 字 = fs.readFileSync(path.join(ROOT, 'tests/shiki-hyou-hidzuke.test.mjs'), 'utf8');
    if (!/1900年の 辺りは「未測定」/.test(字)) throw new Error('★未測定と 書いて いない★');
  });
  T('★時刻は 1日を 1 と する 割合★', () => {
    const h = H.表();
    h.打つ('Z1', '=SUM("00:00")');
    if (Number(h.字('Z1')) !== 0) throw new Error('00:00 は 0');
    h.打つ('Z1', '=SUM("12:00")');
    if (!近い(h.字('Z1'), 0.5)) throw new Error('12:00 は 0.5');
    h.打つ('Z1', '=SUM("23:59:59")');
    if (!近い(h.字('Z1'), 86399 / 86400)) throw new Error('23:59:59');
  });
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
