/* shoshiki.test.mjs — ★書式の 台★（2026-09-15）
 *
 *  ★★なぜ 台を 1つに したか★★（指示役1 の 決め 2026-09-15）
 *    ★TEXT() も 画面も 同じ 書式を 解きます★。2か所で 解くと ★道が 2本★。
 *    ⇒★台は `lib/shoshiki.js` の 1つ★。
 *
 *  ★★押す 紙★★ `docs/measured/kansuu46/golden-shoshiki-dai-2026-09-15.tsv`（192行）
 *    ★実物が 実際に 使って いる 16種 × 代表の 値 12個★
 *    （どの 書式を 測るかは ★実物を 数えて 決めました★＝
 *      TEXT 740回＝aaa 731／m/d 9 ／ 画面 67,542マス＝18種）
 *
 *  ★★この 紙で 立つ 一番 大事な 事★★
 *    ★192行 中 162行は マスの 見た目と TEXT() が 同じ 字★
 *    ★違う 30行は 全部「その 値を その 書式で 出せない」時★
 *      ⇒ TEXT() は #VALUE! ／ 画面は #### ⇒★決めるのは 呼ぶ側★
 *      ⇒★だから 台は 1つで 足ります★（★これが measured で 立った★）
 *
 *  ★★この 見張りが 守る 事★★
 *    ①★紙の 192行を 1本も 飛ばさない★
 *    ②★「出せない」を「出せる」に しない★（★嘘の 字を 出さない★）
 *    ③★呼ぶ側の 分かれ（#VALUE! と ####）が 残って いる★
 *
 *  使い方: node tests/shoshiki.test.mjs
 *          node tests/shoshiki.test.mjs --self-test
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const S = require_(path.join(ROOT, 'lib/shoshiki.js'));

let pass = 0, fail = 0;
const T = (n, f) => {
  try { f(); pass++; console.log('  ✓ ' + n); }
  catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); }
};

console.log('\n[shoshiki] ★書式の 台★ … 実Excel の 紙 192行と 1本ずつ 突き合わせる');

const 紙 = path.join(ROOT, 'docs/measured/kansuu46/golden-shoshiki-dai-2026-09-15.tsv');
const 生 = fs.readFileSync(紙, 'utf8').split(/\r?\n/);

/* ★★1列目は 番号★★（★書式の 字は `#` で 始まります★＝
   `# で 始まる 行は 覚え書き` と ぶつかって ★16種の うち 7種が 黙って 消えました★。
   ⇒ 紙の 1列目を 番号に して 直した＝★ここも 番号で 拾います★） */
const 行たち = 生.map((l) => l.split('\t'))
  .filter((c) => c.length >= 8 && /^[0-9]+$/.test(c[0]));

let 全 = 0, 合 = 0, 違 = 0, 出せない = 0;
const 違う行 = [], 見た書式 = new Set();
for (const c of 行たち) {
  const [, 書, , 名, 値, 見た目, , 同] = c;
  見た書式.add(書);
  全++;
  const r = S.当てる(Number(値), 書);
  /* ★紙の 7列目が「同じ」＝マスの 見た目と TEXT() が 同じ★
     ★「★違う★」＝その 値を その 書式で 出せない★（マスは #### ／ TEXT は #VALUE!） */
  const 期待 = (同 === '同じ') ? 見た目 : null;
  if (期待 === null) 出せない++;
  const ok = (期待 === null) ? !r.出せる : (r.出せる && r.字 === 期待);
  if (ok) { 合++; continue; }
  違++;
  違う行.push('    ' + 書.padEnd(28) + ' ' + 名.padEnd(11)
    + ' 紙 `' + (期待 === null ? '(出せない)' : 期待) + '` ／ うち `'
    + (r.出せる ? r.字 : '(出せない)') + '`');
}

console.log('');
console.log('  ★押した★ … ' + 全 + '行（書式 ' + 見た書式.size + '種）');
console.log('  ★合った★ … ' + 合 + '行 ／ ★合わない★ … ' + 違 + '行');
console.log('  ★その うち「出せない」が 正しい 行★ … ' + 出せない + '行');
if (違う行.length) 違う行.forEach((x) => console.log(x));

T('★紙の 192行を 1本も 飛ばして いない★（書式 16種 全部）', () => {
  if (全 !== 192) throw new Error('押した 行が 192で ない … ' + 全);
  if (見た書式.size !== 16) throw new Error('書式が 16種で ない … ' + 見た書式.size);
});

T('★★紙と 1行も 違わない★★', () => {
  if (違) throw new Error(違 + '行 違う');
});

T('★「出せない」行が ちゃんと 在る★（全部 出せる なら 紙を 読めて いない）', () => {
  if (出せない !== 30) throw new Error('出せない 行が 30で ない … ' + 出せない);
});

T('★★呼ぶ側で 分かれる★★（TEXT は #VALUE! ／ 画面は ####）', () => {
  /* ★負の 数に 日付の 書式★＝実測 … マスは #### ／ TEXT は #VALUE! */
  if (S.字にする(-1234.5, 'm/d') !== null) throw new Error('TEXT は null（呼ぶ側が #VALUE!）');
  const e = S.画面の字(-1234.5, 'm/d', 8);
  if (e.字 !== '########') throw new Error('画面は 幅ぶんの # … ' + e.字);
});

T('★色は 名前で 返す★（塗るのは 呼ぶ側）', () => {
  const r = S.当てる(-1234, '#,##0_);[赤](#,##0)');
  if (r.色 !== 'Red') throw new Error('色 … ' + r.色);
  if (r.字 !== '(1,234)') throw new Error('字 … ' + r.字);
});

T('★実物の TEXT 740回が 通る★（aaa 731 ／ m/d 9）', () => {
  if (S.字にする(45292, 'aaa') !== '月') throw new Error('aaa');
  if (S.字にする(45292, 'm/d') !== '1/1') throw new Error('m/d');
});

T('★★まだ 作って いない 形は 嘘の 字を 出さない★★（#VALUE! と 言う）', () => {
  /* ★和暦（ggge・ge）と 経過時間（[h]）は 2026-10-02 に 作った★＝下の 見張りへ 移した
     ★残る 棚は 分数と 指数★ */
  for (const f of ['# ?/?', '0.00E+00']) {
    if (S.字にする(45292, f) !== null) throw new Error('★字を 出して しまった★ … ' + f);
  }
});

/* ★★「和暦か経過か」の 判じ★★（2026-10-04・TEXT と 画面の 両方が この 1つを 呼ぶ）
     ★偽の 形を 必ず 押さえる★（経営者の 叩き⑴）＝当たると 今の 道（SSF・この層の 書式器）から 外れる
     ★10-02 の TEXT 側の 判じは `0.00E+00` の E を 和暦と 取り違えて #VALUE! に して いた★ */
T('★★判じが 偽の 形は 偽★★（指数の E・General・引用符・逃がし・[Red] 等）', () => {
  const 偽 = ['0.00E+00', '0e+0', '0.0E-3', 'General', 'G/標準', 'h:mm', 'hh:mm:ss', '[Red]0', '[$-411]yyyy/m/d',
    '[>=100]0', '[DBNum1]0', '"ge"0', '\\g0', '#,##0"円"', 'yyyy/m/d', 'm/d(aaa)', '0%', '_e0', '*e0', '[赤]#,##0', '@'];
  const 外 = 偽.filter((f) => S.和暦か経過か(f));
  if (外.length) throw new Error('★真に なった★ ' + 外.join(' / '));
});
/* ★★「日付の書式か」の 判じ★★（10-04 夕・画面の fmtForDisplay が 上限の 所で 呼ぶ・経営者の 叩き⑵） */
T('★★日付の書式か：偽の 形は 偽★★（General・指数の E・引用符・逃がし・[Red]・[$-411]・[DBNum1] 等）', () => {
  const 偽 = ['General', 'G/標準', '0.00E+00', '0e+0', '"d"0', '\\d0', '[Red]0', '[$-411]0', '#,##0"円"', '0%', '@',
    '[>=100]0', '[DBNum1]0', '0.00', '#,##0'];
  const 外 = 偽.filter((f) => S.日付の書式か(f));
  if (外.length) throw new Error('★真に なった★ ' + 外.join(' / '));
});
T('★★日付の書式か：真の 形は 真★★（日付・時刻・経過・和暦・区切り 付き）', () => {
  const 真 = ['yyyy/m/d', 'h:mm', 'hh:mm:ss', 'h:mm AM/PM', 'm/d', 'd', 'yyyy', 'mmm', 'mmmm', 'aaa', 'aaaa', 'ddd',
    'mm:ss', '[m]:ss', 'yyyy/m/d h:mm', 'ge.m.d', '[h]:mm:ss', '[s]', 'm/d;@', '[$-411]yyyy/m/d'];
  const 外 = 真.filter((f) => !S.日付の書式か(f));
  if (外.length) throw new Error('★偽に なった★ ' + 外.join(' / '));
});
T('★★日付の 上限★★（実Excel：2958465.99999 まで 字・2958466 から ####・単位1つの [h][s] は 上限 なし）', () => {
  const 組 = [[2958465.99999, 'yyyy/m/d', '9999/12/31'], [2958465.99999, 'h:mm AM/PM', '11:59 PM'], [2958465, 'mmm', 'Dec'],
    [2958465, 'ddd', 'Fri'], [2958466, '[h]', '71003184'], [2958466, '[s]', '255611462400']];
  for (const [v, f, 期待] of 組) {
    const 出 = S.字にする(v, f);
    if (出 !== 期待) throw new Error(f + ' ' + v + ' … ' + JSON.stringify(出) + '（実Excel ' + 期待 + '）');
  }
  for (const f of ['yyyy/m/d', 'h:mm', 'aaa', '[m]:ss', 'ge.m.d']) {
    if (S.字にする(2958466, f) !== null) throw new Error(f + ' 2958466 は 出せない（####）はず');
  }
});
T('★★判じが 真の 形は 真★★（和暦 g/e・経過 [h][m][s]・区切り 付き）', () => {
  const 真 = ['g', 'gg', 'ggg', 'e', 'ee', 'ge.m.d', 'ggge年m月d日', 'gggee年', '[$-ja-JP]ggge年', '[h]:mm', '[hh]:mm',
    '[m]:ss', '[s]', '[h]:mm:ss.00', 'ge.m.d;@', '[h]:mm;[Red]-[h]:mm'];
  const 外 = 真.filter((f) => !S.和暦か経過か(f));
  if (外.length) throw new Error('★偽に なった★ ' + 外.join(' / '));
});

T('★★和暦と 経過時間は 実Excel の 答え★★（2026-10-02 に 作った）', () => {
  /* ★答えは 実Excel★（この 台の 頭の 注記に 在る 実測 ＋ 実Excel 16.0.20430 の 真値 155本）
     ★全部の 突き合わせは tests/text-format.test.mjs と tests/xlsx-harness が します★ */
  const 組 = [
    [45292, 'ggge年m月d日', '令和6年1月1日'],
    [1.5, '[h]:mm', '36:00'],
    [0.999999, '[h]:mm:ss', '24:00:00'],          /* ★秒で 丸めて 繰り上がる★ */
    [0.999999, '[h]:mm:ss.00', '23:59:59.91'],    /* ★1/100秒で 丸めると 繰り上がらない★ */
    [-1.5, '[h]', '-36'],                          /* ★1つだけの 単位は 負も 出る★ */
  ];
  for (const [v, f, 期待] of 組) {
    const 出 = S.字にする(v, f);
    if (出 !== 期待) throw new Error(f + ' ... ' + JSON.stringify(出) + '（実Excel ' + 期待 + '）');
  }
  /* ★負で 単位が 2つ 以上は 出せない★（TEXT は #VALUE! ／ マスは ####） */
  if (S.字にする(-1.5, '[h]:mm') !== null) throw new Error('負の [h]:mm は 出せない はず');
});

if (process.argv.includes('--self-test')) {
  console.log('\n[shoshiki --self-test] ★わざと 壊したら 赤に なるか★');

  T('★`m` は 時の 隣なら 分★（月に すると 分かる）', () => {
    if (S.字にする(0.520833333333333, 'h:mm') !== '12:30') throw new Error('h:mm');
    if (S.字にする(45292, 'm/d') !== '1/1') throw new Error('m/d は 月');
  });

  T('★丸めて 0 に なったら 符号を 付けない★（-0 を 客に 見せない）', () => {
    if (S.字にする(-0.004, '#,##0_ ') !== '0 ') throw new Error('… ' + JSON.stringify(S.字にする(-0.004, '#,##0_ ')));
  });

  T('★1900年の 起点が 2つ★（通し 1 → 1/1 ／ 通し 0 → 1/0）', () => {
    if (S.字にする(1, 'm/d') !== '1/1') throw new Error('通し 1');
    if (S.字にする(0, 'm/d') !== '1/0') throw new Error('通し 0');
    if (S.字にする(61, 'm/d') !== '3/1') throw new Error('通し 61');
  });

  T('★曜日は 通し番号の 割り算★（通し 0→土・1→日・45292→月）', () => {
    for (const [n, w] of [[0, '土'], [1, '日'], [45292, '月']]) {
      if (S.字にする(n, 'aaa') !== w) throw new Error(n + ' … ' + S.字にする(n, 'aaa'));
    }
  });

  T('★`@` の 区画は 数に 使わない★（`m/d;@` の 負は 出せない）', () => {
    if (S.字にする(-1, 'm/d;@') !== null) throw new Error('負に @ を 当てて いる');
  });

  T('★`_x` は 空きを 1つ★', () => {
    if (S.字にする(1234.5, '#,##0_ ') !== '1,235 ') throw new Error('… ' + JSON.stringify(S.字にする(1234.5, '#,##0_ ')));
  });

  T('★TEXT は 世界共通の 字（General）を 受けない★（画面は 受ける）', () => {
    if (S.字にする(1234.5, 'General') !== null) throw new Error('TEXT が General を 受けた');
    if (S.当てる(1234.5, 'General').字 !== '1234.5') throw new Error('画面は 受ける');
  });

  T('★紙の 1列目が 番号★（書式の `#` で 行が 消えない）', () => {
    const 書式ら = new Set(行たち.map((c) => c[1]));
    let n = 0;
    for (const f of 書式ら) if (f.charAt(0) === '#') n++;
    if (n < 7) throw new Error('★`#` で 始まる 書式が ' + n + '種しか 拾えて いない★（7種 在るはず）');
  });
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
