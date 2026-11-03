/* kansuu-tana.test.mjs — ★台帳の 棚が 本当か、engine に 押して 確かめる★（2026-09-05）
 *
 *  ★★なぜ 要るか（2026-09-05 に 実際に 起きた）★★
 *    `YEN` を ★「足さない（＝Exally で 動かない）」の 棚★に 置いていた。
 *    ⇒ AIへ 渡る 紙 `prompt/kansuu.md` の
 *      ★「Exally内では まだ 動かない（勧めては いけない）」の 列★に 並ぶ。
 *    ⇒★★実際は `=YEN(1234.5)` は ¥1,235 を 返す★★
 *      （`exally-formula.js` の `convertFormula` が 本名 `DOLLAR` に 直している）
 *    ⇒★AIが お客さんに「YEN は 動きません」と 言う★所だった。
 *
 *  ★★字だけ 見る 見張りでは 掴めなかった★★
 *    `tests/formula-extra.test.mjs` は ★理由の 字★を 見ていた
 *    `tests/prompt-file.test.mjs` は ★台帳と 紙が 揃っているか★を 見ていた
 *    ⇒ ★台帳が 間違っていると、揃っているほど 揃って 間違う★
 *    ⇒★★台帳の 外＝engine に 押して 答えを 見るしか ない★★
 *
 *  ★何を 押すか★
 *    ①「動かない」棚の 名前 … ★本番と 同じ道（convertFormula → engine）で #NAME? に なるか★
 *       なれば 正しい／答えが 出るなら ★棚が 嘘★＝赤
 *    ②「別名で動く」棚の 名前 … ★本当に 答えが 出るか★
 *       出なければ ★棚が 嘘★＝赤（今度は 逆向きの 嘘）
 *    ⇒★両方 見る★＝片側だけだと 「全部 動かない 棚に 入れる」で 緑に できてしまう
 *
 *  ★★既に 在る 物と 重なっていないか（作る前に 探した）★★
 *    `tests/xlsx-harness/alias.test.mjs`（2026-08-01）が
 *      ★JIS→DBCS / YEN→DOLLAR の ★直し そのもの★★を 見ている
 *      （入口 convertFormula と 出口 xlsx 書き出しの 2か所）
 *      ⇒★そこには `.FormulaLocal` ↔ `.Formula` の 事も 正しく 書いてあった★
 *      ⇒★08-31 に 私が 台帳へ 書いた「実Excel に 無い」は、
 *        ★repo が 既に 持っていた 正しい 記述と 逆★だった★
 *    ⇒★ここ(kansuu-tana)は 重ならない★＝見るのは ★直し★ではなく
 *      ★★台帳の 棚（AIに 何と 言わせるか）が 本当か★★
 *      ＝alias は「YEN が 動く」を 見る／ここは「棚が 嘘を ついていないか」を 見る
 *
 *  使い方: node tests/kansuu-tana.test.mjs
 *          node tests/kansuu-tana.test.mjs --self-test
 */
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(import.meta.url);
let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };

/** ★純関数★＝棚と「押した 答え」から 食い違いを 返す（self-test で 作り物を 通せる）
 *  @param 動かない 名前の 配列 ／ @param 別名 名前の 配列
 *  @param 押す (名前) => '#NAME?' か 答えの 字 */
export function 棚の嘘(動かない, 別名, 押す) {
  const 出 = { 動かないと言ったのに動く: [], 動くと言ったのに動かない: [] };
  for (const f of 動かない) if (押す(f) !== '#NAME?') 出.動かないと言ったのに動く.push(f + '→' + 押す(f));
  for (const f of 別名) if (押す(f) === '#NAME?') 出.動くと言ったのに動かない.push(f);
  return 出;
}

if (process.argv.includes('--self-test')) {
  console.log('\n[kansuu-tana --self-test] わざと壊して赤になるか');
  /* ★2026-09-05 に 実際に 通した 形★＝動く 物を 動かない 棚に 置いた */
  T('★★動かない と 言ったのに 動く物を 見つける（YEN の 事故）★★', () => {
    const r = 棚の嘘(['LAMBDA', 'YEN'], [], (f) => (f === 'YEN' ? '¥1,235' : '#NAME?'));
    if (r.動かないと言ったのに動く.length !== 1) throw new Error('見つけていない');
  });
  T('★逆向きの 嘘（動くと 言ったのに 動かない）も 見つける', () => {
    const r = 棚の嘘([], ['YEN'], () => '#NAME?');
    if (r.動くと言ったのに動かない.length !== 1) throw new Error('見つけていない');
  });
  T('★正しければ 何も 出さない', () => {
    const r = 棚の嘘(['LAMBDA'], ['YEN'], (f) => (f === 'YEN' ? '¥1,235' : '#NAME?'));
    if (r.動かないと言ったのに動く.length || r.動くと言ったのに動かない.length) throw new Error('誤検知');
  });
  console.log('\n' + pass + ' passed, ' + fail + ' failed');
  process.exit(fail ? 1 : 0);
}

/* ══ 本番＝本当に 押す ═══════════════════════════════════════ */
const HFns = require_(path.join(ROOT, 'hyperformula.full.min.js'));
const EF = require_(path.join(ROOT, 'exally-formula.js'));
const 台帳 = require_(path.join(ROOT, 'lib/formula-extra.js')).数える();

console.log('\n[kansuu-tana] 台帳の 棚を engine に 押して 確かめる');

/* ★プラグインを 積む★＝積み忘れると 素の engine が 答えて ★嘘の 緑★に なる
   （2026-08-29 の 実測＝907本を「合わない」と 誤報告した 家） */
const 積めた = EF.registerExallyFunctions(HFns) === true;
/* ★★lib の プラグインも 全部 積む★★（2026-09-07）
   ⇒ 前は ★exally-formula しか 積んでいなかった★
   ⇒★棚に 載っている 物が「本当は 動く」かを ★本番と 同じ 積み方★で 見ないと 分からない★
   ⇒ CONVERT が「動かない」棚に 載ったまま 間違った 数を 返していたのと ★同じ 家★ */
const HF0 = HFns.HyperFormula;
const H2 = Object.assign(Object.create(HF0), HFns,
  { registerFunctionPlugin: HF0.registerFunctionPlugin.bind(HF0) });
const 積んだlib = [];
for (const n of ['extra', 'nokori', 'kane', 'yosoku']) {
  積んだlib.push(require_(path.join(ROOT, 'lib/formula-' + n + '-plug.js'))
    .つなぐ(H2, require_(path.join(ROOT, 'lib/formula-' + n + '.js')),
      n === 'yosoku' ? function () { return { シート数: 1, 版: 'Exally', 台: 'win', OS: '', 左上: '$A$1' }; } : undefined));
}
/* ★外へ 出る 物は「必ず 断る 作り物」で 積む★＝ここでは 外へ 1回も 出さない */
積んだlib.push(require_(path.join(ROOT, 'lib/formula-soto-plug.js'))
  .つなぐ(H2, require_(path.join(ROOT, 'lib/formula-soto.js')), {
    取る: async () => { throw new Error('ここでは 外へ 出ません'); },
    聞く: async () => { throw new Error('ここでは AI に 聞きません'); },
    再計算: () => {},
  }));
/* ★XML から 取り出す★＝XPath は jsdom の 物を 入れる（無ければ 入れない） */
let XML部品 = null;
try {
  const { JSDOM } = require_('jsdom');
  const w = new JSDOM('').window;
  XML部品 = { DOMParser: w.DOMParser, XPathResult: w.XPathResult };
} catch (e) { XML部品 = null; }
積んだlib.push(require_(path.join(ROOT, 'lib/formula-filterxml-plug.js'))
  .つなぐ(H2, require_(path.join(ROOT, 'lib/formula-filterxml.js')), () => XML部品));
/* ★CELL★＝見た目は 画面が 入れる（ここでは 入れない＝既定の 見た目で 答える） */
積んだlib.push(require_(path.join(ROOT, 'lib/formula-cell-plug.js'))
  .つなぐ(H2, require_(path.join(ROOT, 'lib/formula-cell.js')), null));
const hf = HFns.HyperFormula.buildEmpty({ licenseKey: 'gpl-v3' });
/* ★本番と 同じ★＝book.html は hf を 作った後 initExallyFormula(hf) を 1回 呼ぶ */
EF.initExallyFormula(hf);
const SID = hf.getSheetId(hf.addSheet('S'));

/* ★本番と 同じ 道★＝convertFormula を 通してから engine に 渡す
   （素で 渡すと 本番と 違う物を 測る＝今日の YEN は それで 見落とした） */
/* ★★引数の 形を 手で 決め打たない★★（2026-09-06 実測で 踏んだ）
   ★1つの 形しか 試さないと、LAMBDA を 取る 関数（MAP/REDUCE/SCAN/MAKEARRAY）は
     ★普通の 引数では 一度も 当たらない★
   ⇒★動いているのに「動かない」に 落ちる★＝AIが 客に 嘘を 言う
   ⇒★当たるまで 形を 増やす★／★JS層も 通す★（本番は 3段） */
const 形たち = ['(1)', '()', '(A1:A2,1)', '(1,1)', '(A1:A2)', '(1,1,1)',
  '(A1:A2,LAMBDA(v,v*2))', '(0,A1:A2,LAMBDA(a,b,a+b))', '(2,2,LAMBDA(r,c,r*c))',
  '(x,2,x*3)', '(x,x+1)'];
/* ★台を 建てる★（★読めなければ null＝今まで通り 2段で 押す／黙って 緑に しない★） */
let 板 = null;
try {
  const SH = require_(path.join(ROOT, 'lib/shiki-hyou.js'));
  板 = SH.表();
  板.打つ('A1', 1); 板.打つ('A2', 2); 板.打つ('A3', 3);
} catch (e) { 板 = null; }

function 押す(名前, 引数) {
  const 形 = (引数 === undefined) ? 形たち : ['(' + 引数 + ')'];
  let 最後 = '#NAME?';
  for (const a of 形) {
    const 式 = '=' + 名前 + a;
    /* ①JS層（本番の 1段目）*/
    if (typeof EF._jsComputeFormula === 'function') {
      let j = null;
      try { j = EF._jsComputeFormula(0, 式); } catch (e) { return 'JS層が 投げた'; }
      if (j !== null) return String(j);
    }
    /* ══ ★★2026-09-19 ... ★台の 段を 足しました★★★ ══
         ★訳★ ... 本番の 道は ★3段★です
             ①JS層（`_jsComputeFormula`） → ②★台★（`lib/shiki-hyou.js`） → ③借り物
           ＝`book.html:recalcSheet` の 実物
         ★この 門は ①と ③しか 押して いませんでした★
         ⇒`docs/measured/honban-no-michi.mjs` の 断り
           「★この 段を 飛ばすと 『うちが 負けて いる』向きの 偽の 答えが 出ます★」
         ⇒★実際に 起きて いました★ ... BYROW / BYCOL / LAMBDA / PHONETIC を
           ★「動かない」と 読んで いた★（★本番では 動いて います★）
         ⇒★★門を 甘く したのでは ありません＝★見る 範囲★を 本番に 揃えました★★ */
    if (板) {
      try {
        板.打つ('ZZ9990', 式);
        const _v = 板.値('ZZ9990');
        if (_v && _v.溢れ === true) {
          const 頭 = (_v.並び[0] || [])[0];
          return String(頭 && 頭.値 !== undefined ? 頭.値 : 頭);
        }
        if (_v && !(_v.型 === '誤' && String(_v.値) === '#NAME?')) return String(板.字('ZZ9990'));
      } catch (e) { /* ★台が 投げたら 下の 借り物へ★（本番と 同じ） */ }
    }
    /* ③engine（借り物） */
    let 後; try { 後 = EF.convertFormula(式); } catch (e) { continue; }
    try {
      hf.setSheetContent(SID, [[1], [2], [3], [後]]);
      const v = hf.getCellValue({ sheet: SID, row: 3, col: 0 });
      if (v && v.type === 'NAME') { 最後 = '#NAME?'; continue; }
      if (v && v.type) return '#' + v.type;
      return String(v);
    } catch (e) { return 'engine が 投げた'; }
  }
  return 最後;
}

const 動かない = Object.keys(台帳.足さない).sort();
const 別名 = Object.keys(台帳.別名で動く || {}).sort();
const 嘘 = 棚の嘘(動かない, 別名, 押す);

T('★プラグインを 積めた（積み忘れると 素の engine が 答えて 嘘の 緑に なる）', () => {
  if (!積めた) throw new Error('registerExallyFunctions が true を 返さない');
});

T('★★「動かない」棚は 本当に 動かない（AIが 嘘を 言わない）★★', () => {
  if (嘘.動かないと言ったのに動く.length) {
    throw new Error('★動くのに「動かない」と 台帳に 書いてあります★: '
      + 嘘.動かないと言ったのに動く.join(' / ')
      + '\n   → ★AIが お客さんに「使えません」と 言います★'
      + '\n   → 台帳(lib/formula-extra.js)の ★別名で動く★ か ★足す★ へ 移してください');
  }
});

T('★★「別名で動く」棚は 本当に 動く★★', () => {
  if (嘘.動くと言ったのに動かない.length) {
    throw new Error('★動かないのに「動く」と 台帳に 書いてあります★: '
      + 嘘.動くと言ったのに動かない.join(' / ')
      + '\n   → ★AIが 動かない 式を 勧めます★');
  }
});

T('★この 検査が 空振りしていない（棚が 空でない）', () => {
  /* ★★数の 線を 引かない★★（2026-09-06 指示役の 差し戻し）
     ★私は 15→9 に 減った時「10未満なら赤」を「5未満なら赤」に 下げた★
     ⇒★次に 4個に なったら また 下げる★＝★線を 下げれば 通る＝いつも真＝死ぬ★
     ⇒★★名前で 見る★★
       ①棚が 空では ない（★1個以上＝これは 下げようが ない★）
       ②★必ず 動かない 物★が 居る
         ★★2026-09-07 に 2個 外した（★決めが 在って 外した★）★★
           司さん「★全部やって★」
           ⇒ WEBSERVICE／STOCKHISTORY を ★作った★（うちの サーバを 通す 形）
           ⇒★『通らないから 下げた』では なく『★決めが 出たから 外した★』★
           ⇒ 作った 中身 … docs/measured/soto/TSUKUTTA.md
         ★残る 2個★ … IMAGE／RTD
           IMAGE … ★絵を セルの 上に 置く 仕組みが まだ 無い★
           RTD  … ★向こうから 押してくる 相手（配信の 元）が うちに 無い★
           ⇒★どちらも「外の 契約」では なく「うちの 作り」が 足りない★
           ⇒★作ったら その時 外す★ */
  if (!動かない.length) throw new Error('★動かない棚が 空★＝この 検査は 何も 見ていない');
  const 必ず居る = ['IMAGE', 'RTD'];
  const 居ない = 必ず居る.filter((f) => 動かない.indexOf(f) < 0);
  if (居ない.length) {
    throw new Error('★外へ 聞きに 行く 物が 棚から 消えています★: ' + 居ない.join(' / ')
      + '\n   → 司さんが「外へ 出てよい」と 決めたのなら ★この 4個の 名前を ここから 外す★'
      + '\n   → 決めが 無いのに 消えたのなら ★台帳が 壊れています★');
  }
  if (!別名.length) throw new Error('別名で動く棚が 空です（YEN が 居るはず）');
});

console.log('\n── 実測（本番と 同じ道＝convertFormula → engine） ──');
/* ★ここは「名前が 通るか」だけを 見る★＝#NAME? でなければ 通っている。
   引数の 数は 名前ごとに 違うので #N/A などは 出る（それは 名前の 話では ない）。
   ★引数まで 込みの 実測は tests/betsumei-zenbu.test.mjs が やる★ */
for (const f of 別名) console.log('  ' + f + ' … 名前は 通る（' + 押す(f, '1234.5') + '）');
console.log('  「動かない」棚 ' + 動かない.length + '個 … 全部 #NAME?（' + 動かない.join(' ') + '）');
console.log('\n' + pass + ' passed, ' + fail + ' failed');
hf.destroy();
process.exit(fail ? 1 : 0);
