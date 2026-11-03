/* xlfn-morenashi.test.mjs — ★`_xlfn.` の 一覧に 漏れが 無いか★（2026-09-16）
 *
 *  ★★なぜ 要るか★★
 *    `lib/xlsx-io.js` の 一覧に 1つでも 漏れが 有ると
 *    ★その 式だけ 実Excel で #NAME?★ に なります。
 *    ★紙（golden-*.tsv）が 全部 緑でも 見つかりません★
 *      ＝台の 中では 正しく 計算できるから。
 *
 *  ★★同じ 型で 3回 踏んで います★★
 *    ①PERMUTATIONA（2026-08-02）… 実Excel で その 式だけ #NAME?
 *    ②RANK.AVG（2026-08-01）… RANK.EQ は 入って いたのに これだけ 抜けて いた
 *    ③★2026-09-16 … ★91個★ 抜けて いた★（棚58〜67 で 足した 物＋前から 在った 物）
 *    ⇒★手で 足すと また 漏れる★＝★実Excel に 聞いた 紙と 突き合わせる★
 *
 *  ★★この 見張りが 見る 物★★
 *    ①紙が 読めて いる（★空振りして いない★）
 *    ②★紙に 在る 名前が 一覧に 全部 在る★（＝漏れが 無い）
 *    ③一覧の 数が 紙と 合う（★黙って 痩せない★）
 *
 *  ★★紙の 取り直し方★★
 *    node docs/measured/osu-xlfn-zenbu.mjs --namae
 *    pwsh -NoProfile -File docs/measured/toru-xlfn-zenbu.ps1
 *    node docs/measured/osu-xlfn-zenbu.mjs
 *
 *  使い方: node tests/xlfn-morenashi.test.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
let pass = 0, fail = 0;
const T = (n, f) => {
  try { f(); pass++; console.log('  ✓ ' + n); }
  catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); }
};

console.log('\n[xlfn-morenashi] ★`_xlfn.` の 一覧に 漏れが 無いか★');

const 紙道 = path.join(ROOT, 'docs/measured/golden-xlfn-2026-09-16.tsv');
const 紙 = fs.readFileSync(紙道, 'utf-8')
  .split(/\r?\n/).filter((l) => l && !l.startsWith('#')).map((s) => s.trim());

const io = fs.readFileSync(path.join(ROOT, 'lib/xlsx-io.js'), 'utf-8');
const m = /var XLFN = \[([\s\S]*?)\];/.exec(io);
const 一覧 = m ? [...new Set((m[1].match(/'[^']+'/g) || []).map((s) => s.slice(1, -1)))] : [];

console.log('      … 紙 ' + 紙.length + '個 ／ 一覧 ' + 一覧.length + '個');

T('★紙が 読めて いる（★空振りして いない★）★', () => {
  if (紙.length < 100) throw new Error('★' + 紙.length + '個しか 読めない★（紙が 壊れて いる）');
  /* ★実Excel が 付けると 測った 物が ちゃんと 入って いるか（名指し）★ */
  for (const n of ['NORM.DIST', 'CHISQ.DIST', 'CONFIDENCE.NORM', 'ISOMITTED', 'XLOOKUP']) {
    if (紙.indexOf(n) < 0) throw new Error('★紙に ' + n + ' が 無い★');
  }
});

T('★★一覧に 漏れが 無い（★漏れると その 式だけ 実Excel で #NAME?★）★★', () => {
  const 無 = 紙.filter((n) => 一覧.indexOf(n) < 0);
  if (無.length) {
    throw new Error('★' + 無.length + '個 漏れて いる★ … ' + 無.slice(0, 12).join(' ')
      + (無.length > 12 ? ' …' : ''));
  }
});

/* ══ ★★2026-09-20 に 足した 段＝★実Excel 自身が 書いた 字★から 数える★★ ══
     ★なぜ 要るか★
       上の 紙（09-16）は ★台が 知る 名前を 実Excel に 打たせた★ 物ですが、
       ★その 分母から 45個 漏れて いました★（`WRAPROWS` `WRAPCOLS` を 含む）
       ⇒★分母が 漏れて いれば 紙は 緑の まま★＝上の 3つでは 捕まりません。
     ★何が 起きたか★（2026-09-20 実Excel 16.0 build 20326 で 実測）
       `WRAPROWS` `WRAPCOLS` は ★台にも 皮にも 在る＝お客さんが 打てる★
       ⇒ でも 一覧に 無い ので ★裸で 書き出す★
       ⇒★相手の Excel で #NAME?（-2146826259）★
       ⇒★★溢れ先も 全部 空＝表が 丸ごと 消える★★
         （うちの 画面では 2行3列の 表／実Excel では 誤り 1マスだけ）
     ★ここが 見る 物★
       `golden-jitsu-excel-no-shirushi-2026-09-20.tsv` は
       ★実Excel が 自分で 書いた 字を そのまま 写した 紙★です（判じを して いない）
       ⇒★実Excel が `_xlfn.` を 付けた 名前は 一覧に 在る★
       ⇒★実Excel が 裸で 書いた 名前は 一覧に 無い★（付けると 逆に 壊れる）
     ★`_xlfn._xlws.` も `_xlfn.` として 数えます★
       ＝SORT / FILTER は 実Excel が `_xlfn._xlws.` と 書きますが、
         注31 が「`_xlfn.` でも 通る」と 実測して います（★その 裏取りは 別★） */
const 実紙道 = path.join(ROOT, 'docs/measured/golden-jitsu-excel-no-shirushi-2026-09-20.tsv');
const 実紙 = fs.existsSync(実紙道) ? fs.readFileSync(実紙道, 'utf-8') : '';
const 実が付ける = new Set();
const 実が裸 = new Set();
const 溢れる = new Set();   /* ★`cm` が 在る＝溢れる★＝割れると 表が 丸ごと 消える */
const 除いた = new Set();   /* ★黙って 除かない★＝下で 名前を 出します */
for (const 行 of 実紙.split(/\r?\n/)) {
  if (!行 || 行.startsWith('#')) continue;
  const 式 = (行.split('\t')[3] || '').trim();
  if (!式 || 式 === '(無し)') continue;
  /* ★★除く 物を ★名指しで★ 数えます★★（★黙って 除かない★）
       ㋐`_xlpm.` が 在る 式（LAMBDA の 一族）
         ＝`lib/xlsx-io.js` の `NEEDS_XLPM` が ★投げて 書き出しを 止めます★
         ＝2026-09-20 実測: `=MAP(A1:A3,LAMBDA(x,x*2))` 等 ★6個とも 投げた★
           「LAMBDA は引数名に _xlpm. が要るため、この書き出しでは未対応です」
         ⇒★壊れた 物を 作らない＝一覧に 入れる 必要が 無い★
       ㋑`_xleta.` が 在る 式（GROUPBY / PIVOTBY）
         ＝`lib/formula-extra.js` が「ピボットテーブル 自体が まだ 無い」と 書いて います
         ＝2026-09-20 実測: ★台が 持つ 376個の 中に 無い＝お客さんは 打てません★
       ★どちらも「打てて 書き出せる のに 一覧に 無い」物では 在りません★ */
  if (式.indexOf('_xlpm.') >= 0 || 式.indexOf('_xleta.') >= 0) {
    const 名 = /^_xlfn\.(?:_xlws\.)?([A-Z][A-Z0-9_.]*)\s*\(/.exec(式);
    if (名) 除いた.add(名[1]);
    continue;
  }
  const 頭 = /^_xlfn\.(?:_xlws\.)?([A-Z][A-Z0-9_.]*)\s*\(/.exec(式);
  if (頭) {
    実が付ける.add(頭[1]);
    /* ★紙の 2列目が `cm` の 有無★＝★重さが 違います★ */
    if ((行.split('	')[1] || '').indexOf('在') >= 0) 溢れる.add(頭[1]);
    continue;
  }
  const 裸 = /^([A-Z][A-Z0-9_.]*)\s*\(/.exec(式);
  if (裸) 実が裸.add(裸[1]);
}
実が付ける.delete('LAMBDA');
console.log('      ＝ ★実Excel が 付けた★ ' + 実が付ける.size + '個 ／ ★裸で 書いた★ ' + 実が裸.size + '個'
  + ' ／ ★除いた★ ' + 除いた.size + '個');
if (除いた.size) console.log('        ★除いた 物★ ＝ ' + [...除いた].join(' ')
  + '（★LAMBDA の 一族＝書き出しが 投げて 止まる／ピボットは まだ 無い★）');

T('★★実Excel の 紙が 読めて いる（空振りして いない）★★', () => {
  if (!実紙) throw new Error('★紙が 無い★ ' + 実紙道);
  if (実が付ける.size < 5) throw new Error('★' + 実が付ける.size + '個しか 読めない★');
  if (実が裸.size < 1) throw new Error('★裸の 名前を 1つも 読めて いない★（対照が 効いて いない）');
  for (const n of ['WRAPROWS', 'WRAPCOLS', 'SEQUENCE']) {
    if (!実が付ける.has(n)) throw new Error('★紙に ' + n + ' が 無い★');
  }
  if (!実が裸.has('TRANSPOSE')) throw new Error('★裸の 対照（TRANSPOSE）が 紙に 無い★');
});

T('★★実Excel が `_xlfn.` を 付けた 名前が 一覧に 全部 在る★★', () => {
  const 無 = [...実が付ける].filter((n) => 一覧.indexOf(n) < 0);
  if (無.length) {
    throw new Error('★' + 無.length + '個 漏れて いる★ ＝ ' + 無.join(' ')
      + '／★この まま 書き出すと 相手の Excel で #NAME? に なります★'
      /* ★重さを 分けて 出します★（2026-09-20）
           溢れる 物 ＝ ★表が 丸ごと 消えます★（WRAPROWS / WRAPCOLS / MODE.MULT）
           溢れない 物 ＝ ★その マスだけ★（FORMULATEXT）
         ★門の 字が 大げさを 間違えると 次の 人が 読み違えます★ */
      + (無.some((n) => 溢れる.has(n))
         ? '／★溢れる 物が 入って います＝表が 丸ごと 消えます★ '
           + 無.filter((n) => 溢れる.has(n)).join(' ')
         : '／★どれも 溢れません＝その マスだけ★'));
  }
});

T('★★実Excel が 裸で 書いた 名前は 一覧に 無い★★（付けると 逆に 壊れる）', () => {
  const 余 = [...実が裸].filter((n) => 一覧.indexOf(n) >= 0);
  if (余.length) {
    throw new Error('★' + 余.length + '個 余計★ ＝ ' + 余.join(' ')
      + '／★実Excel は 裸で 書いて います★');
  }
});

T('★一覧が 黙って 痩せて いない★', () => {
  /* ★数は 決め打ち★＝★減らしたら ここも 直す＝直さないと 赤で 止まる★
     ★2026-09-20★ 130 → ★132★（`WRAPROWS` `WRAPCOLS` を 足した）
       ＝★実Excel で #NAME? に なる のを 実測して から 足しました★
       ＝紙 `golden-wrap-excel-2026-09-20.tsv`
     ★2026-09-20★ 132 → ★133★（`MODE.MULT` を 足した）
       ＝★この 門が 自分で 見つけました★（紙が 増えた 瞬間 赤に なった）
     ★2026-09-20★ 133 → ★134★（`FORMULATEXT` を 足した）
       ＝★この 門が 自分で 見つけた ★２回目★
     ★2026-09-20★ 134 → ★135★（`ISFORMULA` を 足した）
       ＝★この 門が 自分で 見つけた ★３回目★
       ＝★紙の 分母が 62個 → 376個（総当たり）に なった 瞬間 */
  const 期待 = 135;
  if (一覧.length !== 期待) {
    throw new Error('★一覧が ' + 一覧.length + '個★（' + 期待 + '個の はず）'
      + '／足したなら ここも 直す');
  }
});

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (fail) process.exit(1);
