/* kobore-basho.test.mjs — ★場所を 見る 関数と ROW/COLUMN の 向き★（2026-09-21）
 *
 *  ★★なぜ 要るか★★
 *    `tests/kobore-nakami.test.mjs` の 紙（09-20）は ★1つの 盤面★で 取って います。
 *    その 盤面には ★式の 入った マスが 1つも 有りません★
 *    ⇒`ISFORMULA` も `FORMULATEXT` も ★FALSE 側／#N/A 側しか 押せません★
 *    ⇒★「本当に 打って ある 字を 見て いるか」は 1本も 測れて いません★
 *    ⇒この 紙は ★A2 だけ 式（`=1+1`）★に して あります。
 *
 *    もう 1つ。09-20 の 紙は ★縦に 長い 範囲しか 有りません★（A1:A3）
 *    ⇒`COLUMN` は ★1マス★しか 出ず ★横に 溢れる 向きが 測れません★
 *    ⇒この 紙は ★E1:F2（2行2列）★を 持って います。
 *
 *  ★★物差し★★
 *    `docs/measured/golden-jitsu-excel-basho-kansuu-2026-09-21.tsv`
 *    ・実Excel 16.0 build 20326（経営者1 が 2026-09-21 に 取りました）
 *    ・★中身まで★ 在ります（`|` で マス ／ ` / ` で 行）
 *
 *  ★★盤面を どう 作るか★★
 *    ★この 紙には `#材料` の 行が 有りません★（頭の 注に 文で 書いて あります）
 *    ⇒★道具の 中に 書き写して います★＝★決め打ちです★
 *    ⇒★だから 紙が 持って いる ★対照★ で 検算します★
 *        `=SEQUENCE(2)` が 2マス ／ `=SUM(A1:A3)` が 5
 *      ⇒★盤面を 書き写し間違えたら まず ここが 赤に なります★
 *      ⇒★対照が 紙から 消えたら それも 赤に します★（★黙って 検算が 消えない★）
 *
 *  ★★見て いない 事★★
 *    ・★画面では ありません★（台で 押して います）
 *    ・`CELL` は ★溢れない★ 側として 1本 押して います（★入れると 逆に 壊れる★の 見張り）
 *
 *  使い方: node tests/kobore-basho.test.mjs
 *          node tests/kobore-basho.test.mjs --self-test
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const H = require_(path.join(ROOT, 'lib/shiki-hyou.js'));

let pass = 0, fail = 0;
const NL = String.fromCharCode(10);
const TB = String.fromCharCode(9);
const CR = String.fromCharCode(13);
const T = (n, f) => {
  try { f(); pass++; console.log('  ok   ' + n); }
  catch (e) { fail++; console.log('  NG   ' + n + NL + '       ' + e.message); }
};

const 紙道 = path.join(ROOT, 'docs/measured/golden-jitsu-excel-basho-kansuu-2026-09-21.tsv');
const 字 = fs.existsSync(紙道)
  ? fs.readFileSync(紙道, 'utf8').split(String.fromCharCode(0xFEFF)).join('')
  : '';

/* ★柱から 列を 探します★（★決め打ちに しない★＝列が 増えても 動く） */
let 式列 = 3, 数列 = 4, 中列 = 5;
for (const l of 字.split(NL)) {
  if (!l.startsWith('#')) continue;
  const h = l.replace('#', '').split(TB).map((z) => String(z).trim());
  const i0 = h.findIndex((z) => z === '式');
  const i1 = h.findIndex((z) => z.indexOf('埋まった数') >= 0);
  const i2 = h.findIndex((z) => z.indexOf('中身') >= 0);
  if (i0 < 0 || i1 < 0 || i2 < 0) continue;
  式列 = i0; 数列 = i1; 中列 = i2;
  break;
}

const 組 = [];
for (const l0 of 字.split(NL)) {
  const l = l0.split(CR).join('');
  if (!l || l.startsWith('#')) continue;
  const c = l.split(TB);
  const 式 = String(c[式列] || '').trim();
  if (!式.startsWith('=')) continue;
  const 数 = parseInt(c[数列], 10);
  if (!isFinite(数)) continue;
  組.push({
    組: String(c[0] || '').trim(),
    名: String(c[1] || '').trim(),
    式: 式,
    数: 数,
    中身: String(c[中列] || '').trim(),
  });
}

/* ★★盤面★★（紙の 頭の 注を 書き写した 物＝★決め打ち★）
     ★A2 だけ 式★＝ここが この 紙の 要です */
const 盤面 = [
  ['A1', '1'], ['A2', '=1+1'], ['A3', '2'],
  ['E1', '1'], ['F1', '2'], ['E2', '3'], ['F2', '4'],
];

function 押す(式) {
  const h = H.表();
  for (const 組み of 盤面) h.打つ(組み[0], 組み[1]);
  h.打つ('BZ1', 式);
  const v = h.値('BZ1');
  if (v && v.溢れ === true) {
    const 表 = (v.並び || []).map((段) => 段.map(
      (x) => (x && x.値 !== undefined ? String(x.値) : String(h.字('BZ1')))));
    return {
      数: 表.reduce((a, r) => a + r.length, 0),
      中身: 表.map((r) => r.join('|')).join(' / '),
    };
  }
  return { 数: 1, 中身: String(h.字('BZ1')) };
}

/* ★紙の 中身は 6行x3列の 窓★＝空の マスが `|` で 並びます。
     ⇒★埋まって いる 所だけ 取り出して 比べます★ */
function 削る(s) {
  const 出 = [];
  for (const 段 of String(s).split(' / ')) {
    const ま = 段.split('|').map((z) => String(z).trim()).filter((z) => z !== '');
    if (ま.length) 出.push(ま.join('|'));
  }
  return 出.join(' / ');
}
/* ★真偽の 字は 台と 紙で 書き方が 違います★（`false` と `FALSE`） */
function 揃える(s) {
  return String(s).split(' / ').map((段) => 段.split('|').map((z) => {
    const t = String(z).trim();
    if (t === 'false' || t === 'FALSE') return 'FALSE';
    if (t === 'true' || t === 'TRUE') return 'TRUE';
    return t;
  }).join('|')).join(' / ');
}

console.log('');
console.log('[kobore-basho] ★場所を 見る 関数と ROW/COLUMN の 向き★');
console.log('      ＝ 紙 ' + 組.length + '本 ／ 式=' + 式列 + '列目'
  + ' ／ 埋まった数=' + 数列 + '列目 ／ 中身=' + 中列 + '列目');

T('★紙が 読めて いる（空振りして いない）★', () => {
  if (!字) throw new Error('★紙が 無い★ ' + 紙道);
  if (組.length < 8) throw new Error('★' + 組.length + '本しか 読めない★');
  for (const f of ['=FORMULATEXT(A1:A3)', '=ISFORMULA(A1:A3)', '=COLUMN(E1:F2)', '=ROW(E1:F2)']) {
    if (!組.some((x) => x.式 === f)) throw new Error('★紙に ' + f + ' が 無い★');
  }
});

/* ══ ★★対照★★ ══（★盤面を 書き写し間違えて いないか★） */
const 対照 = 組.filter((x) => x.組 === 'taishou');
T('★★対照が 紙に 残って いる★★（★黙って 検算が 消えない★）', () => {
  if (対照.length < 2) throw new Error('★対照が ' + 対照.length + '本★（2本 の はず）');
});
T('★★盤面が 合って いる（対照で 検算）★★', () => {
  for (const q of 対照) {
    const 出 = 押す(q.式);
    if (出.数 !== q.数 || 揃える(出.中身) !== 揃える(削る(q.中身))) {
      throw new Error('★' + q.式 + '★ 実 ' + q.数 + 'マス ' + 削る(q.中身)
        + ' ／ 内 ' + 出.数 + 'マス ' + 出.中身
        + '／★盤面を 書き写し間違えて いませんか★');
    }
  }
});

/* ══ ★★本体★★ ══
     ★1本ずつ 名指しで 止めます★＝8本しか 無いので ★数で ごまかしません★ */
const 外れ = [];
for (const q of 組) {
  if (q.組 === 'taishou') continue;
  let 出;
  try { 出 = 押す(q.式); }
  catch (e) { 外れ.push(q.式 + ' ★投げた★ ' + e.message); continue; }
  const 実 = 削る(q.中身);
  if (出.数 !== q.数) { 外れ.push(q.式 + ' ★マスの数★ 実 ' + q.数 + ' / 内 ' + 出.数); continue; }
  if (揃える(出.中身) !== 揃える(実)) {
    外れ.push(q.式 + ' ★中身★ 実 ' + 実 + ' / 内 ' + 出.中身);
  }
}
T('★★紙の 1本ずつが 数も 中身も 合う★★', () => {
  if (外れ.length) {
    throw new Error('★' + 外れ.length + '本 外れ★' + NL + '       ' + 外れ.join(NL + '       '));
  }
});

console.log('');
console.log('  ★見て いない 事★ ... 画面では ありません（台で 押して います）');
console.log('kobore-basho: ' + pass + ' 緑 / ' + fail + ' 赤');

if (process.argv.includes('--self-test')) {
  console.log('');
  console.log('--self-test: ★わざと 外して 赤に なるか★（★lib は 1字も 触りません★）');
  let 悪 = 0;
  /* ㋐ ★盤面の A2 を 式で なく する★
       ＝`ISFORMULA` が 本当に ★打って ある 字★を 見て いるなら 中身が 変わる はず */
  const 元 = 盤面[1][1];
  盤面[1][1] = '2';
  const あ = 押す('=ISFORMULA(A1:A3)');
  盤面[1][1] = 元;
  if (揃える(あ.中身) === 'FALSE|TRUE|FALSE') {
    console.log('  NG   ★A2 を 式で 無くしたのに 同じ 答え★'); 悪++;
  } else {
    console.log('  ok   ★A2 を 式で 無くすと 中身が 変わる（打って ある 字を 見て います）★');
  }
  /* ㋑ ★紙の 中身を 書き換える★ ＝ 突き合わせが 本当に 中身を 見て いるか */
  const 的 = 組.find((x) => x.式 === '=FORMULATEXT(A1:A3)');
  if (!的) { console.log('  NG   ★FORMULATEXT の 行が 紙に 無い★'); 悪++; }
  else {
    const 前 = 的.中身;
    的.中身 = 前.split('#N/A').join('#VALUE!');
    const い = 押す(的.式);
    const 合 = 揃える(い.中身) === 揃える(削る(的.中身));
    的.中身 = 前;
    if (合) { console.log('  NG   ★紙を 書き換えたのに 合って しまいました★'); 悪++; }
    else console.log('  ok   ★紙を 書き換えると 合わなく なる（中身を 見て います）★');
  }
  process.exit(悪 ? 1 : 0);
}
process.exit(fail ? 1 : 0);
