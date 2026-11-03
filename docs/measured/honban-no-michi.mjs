/* honban-no-michi.mjs — ★★測り道具が 通る「本番の 道」を 1本に する★★（2026-09-15）
 *
 *  ★★なぜ 作ったか（数えて から 作りました）★★
 *    2026-09-15 に 数えたら ★同じ 建て方が 6本の 道具に 写されて いました★
 *      osu-linest-hyou ／ osu-marume-mae-ato ／ osu-okane-4kansuu
 *      osu-wakeru ／ osu-xirr-sakaime ／ osu-yosoku-3kansuu
 *    ★つなぐ行 のべ 35★／★建てて いるのは どれも 1回★
 *    （`buildEmpty` の 字が 4-1-4-1-4-4 と ばらけて 見えたのは ★覚書きの 中の 引用★）
 *
 *  ★★写しが 危ない 訳（前に 5回 転んで います）★★
 *    2026-09-09〜10 に ★1日で 5回★ 転びました（`tests/hakaridai-mon.test.mjs` の 頭に 全部 在る）
 *      ・プラグインを 3本しか 積んで いなかった（本番は 8本）
 *      ・`smartRounding` が 既定（true）＝★803.6538461538445 が 803.65384615 に なる★
 *      ・板ごと 入れて いた（本番は 1マスずつ）… ★「本番が 壊れて いる」と 報告する 一歩 手前★
 *    ⇒★写しが 1本でも 本番と ずれたら 嘘の 数字が 出ます★
 *
 *  ★★これから 起きる 事（今 寄せる 一番の 訳）★★
 *    ★借り物（HyperFormula）を 外す★＝★自前の 関数 86個の 付け先を 変える★
 *      （足す 65個 ＋ 包む 21個／★司さんの 実物は この 86個を 1つも 使いません★）
 *    ⇒★建て方が 6本 在ると「本番は 動くのに 測り道具は 古い 付け方」が 起きます★
 *    ⇒★2026-09-15 に `板か()` の 写しで 実際に 起きた 事の 大きい版★
 *
 *  ★★この 道は 1本／違いは 引数★★（指示役1 の 決め・2026-09-15）
 *    ★別の 建て方を 作らない★＝★軽くしたい 等が 在れば ここの 引数に する★
 *
 *  ★★見て いない 事（★書かない 見張りは「全部 守った」と 読まれる★）★★
 *    ・★ここは node の 台です★＝★画面の 数では ありません★
 *      （`setCell(r,c,v)` が 画面側／`setCellFormula` は エンジンに 入れるだけ）
 *      ⇒★画面の 事を 言いたい なら ブラウザで 押す★
 *    ・★外へ 出る 関数は 出させません★（`取る`／`聞く` が 投げます）
 *      ⇒★WEBSERVICE 等の「本当に 取れるか」は ここでは 測れません★
 *    ・★本番と 同じ 本数・同じ 建て方かは 見ます★が ★答えが 合うかは 見ません★
 *
 *  使い方:
 *    const 道 = await 建てる();          // { HFns, HF0, H, EF, hf, SID, 積んだ }
 *    const 後 = 道.EF.convertFormula(式);   // ★★必ず 通す★★（下）
 *    道.hf.setSheetContent(道.SID, 表);     // ★★材料を 先に 入れる★★（下）
 *    道.EF._jsComputeFormula(0, 式) ?? 道.hf.getCellValue({ sheet: 道.SID, row, col });
 *
 *  ★★本番の 道は 3段＝1段でも 飛ばすと『うちが 負けて いる』向きの 偽の 答えが 出ます★★
 *    ①★`convertFormula` を 通す★（★飛ばすと 一番 痛い★）
 *       ここに 在る 物 … JIS→DBCS ／ YEN→DOLLAR ／ LET・LAMBDA の 展開 ／
 *                        ISREF・AREAS の 書き換え ／
 *                        ★★裸の `TRUE` `FALSE` → `TRUE()` `FALSE()`★★
 *       ★実測（2026-09-15）★
 *         `=IF(TRUE,1,2)`  飛ばす ★#NAME?★ ／ 通す ★1★
 *         `=AND(TRUE,TRUE)` 飛ばす ★#NAME?★ ／ 通す ★TRUE★
 *         `=NOT(FALSE)`     飛ばす ★#NAME?★ ／ 通す ★TRUE★
 *         `=VDB(10000,0,5,3,5,2,TRUE)` 飛ばす ★#NAME?★ ／ 通す ★1382.3999999999996★
 *       ★★`=SUM(A1:A3)` は 飛ばしても 同じ（6）★★
 *         ⇒★★だから 前の 見本（`=SUM(A1:A3)`）が 一番 危なかった★★
 *         ＝★飛ばして いるのに 見本の 上では 何も 起きない★
 *       ★2026-09-15 経営者1 が これを 踏んで ★11/700★ と 出した（正しくは ★14/700★）
 *         ⇒★危うく「裸の TRUE が 通らない＝客に 出る 大欠陥」と 出す 所だった★
 *    ②★材料を 板に 先に 入れる★
 *       ★飛ばすと JS層（`_jsSet` 27個）が ★空の 板★を 読みます★
 *       ★実測★ `=DSUM(A1:B5,1,A1:A2)` … 入れる前 ★0★ ／ 入れた後 ★2★（実Excel 2）
 *    ③★JS層 → エンジン の 順★（JS層が `null` を 返した 時だけ エンジン）
 *  ⇒★★この 3段を 自分で 書かず `kansuu46-no-dodai.mjs` の `押す()` を 使うのが 安全★★
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ここ = path.dirname(fileURLToPath(import.meta.url));
export const ROOT = path.join(ここ, '..', '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));

/** ★本番（book.html）が 読む プラグインの 数★＝★書き込まない★ */
export function 本番のプラグイン数() {
  const html = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf-8');
  return new Set([...html.matchAll(/lib\/(formula-[a-z]+)-plug\.js/g)].map((m) => m[1])).size;
}

/** ★本番の 建て方★＝★book.html から 読む★（★数を 書き込まない★） */
export function 本番の建て方() {
  const html = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf-8');
  const m = /buildEmpty\(\{([\s\S]{0,300}?)\}\)/.exec(html);
  if (!m) throw new Error('★book.html の buildEmpty が 読めない★');
  return m[1].replace(/\s+/g, '');
}

/**
 * ★★台に 1マス 聞く★★（★`book.html` の `_台に聞く` と ★同じ 混ぜ方★）
 *   ★使わない 時★
 *     ・`#NAME?` ... 台が 知らない（★借り物に 落とす★）
 *     ・溢れ     ... 1つの 式が 何マスにも 広がる（★並べるのは 借り物の 側★）
 *   ⇒★どちらも `null` を 返します★＝★半分 合う 答えを 出さない★
 *   ★★`book.html` と ここで 2か所に 同じ 決まりが 在ります★★
 *     ⇒★片方を 直したら もう 片方も 直して ください★
 *     ⇒`tests/tsunagu-mon.test.mjs` が ★両方に 落とす道が 在るか★を 見ます
 * @returns {string|null} 台の 答え（字）／答えられなければ null
 */
export function 台に聞く(板, 式) {
  if (!板) return null;
  if (!式 || String(式).charAt(0) !== '=') return null;
  try {
    板.打つ('ZZ9999', String(式));
    const v = 板.値('ZZ9999');
    if (!v) return null;
    /* ══ ★★2026-09-19 ... ★溢れを null に して いた のを やめました★★★ ══
         ★前★ ... `if (v.溢れ === true) return null;`
                ＝★溢れは 借り物へ★ という ★昔の `book.html` の 写し★
         ★今の 本番（`book.html:recalcSheet`）★
           ①JS層が 答えたら それ
           ②台が 答えたら それ
           ③★台が 溢れ かつ 借り物が 誤り（`#` で 始まる）なら ★台の 溢れを 使う★★
           ④でなければ 借り物の 溢れ
         ⇒★★道具が 本番より 1段 古く なって いました★★
         ⇒★何が 起きたか★
           `=MAP(A1:A5,LAMBDA(x,x*2))` … ★台は 2 を 出して います★
           でも ここで null を 返す ので 借り物へ 落ち、
           ★借り物は MAP を 知らない ⇒ `#NAME?`★
           ⇒`tests/ramuda-honban.test.mjs` が ★CI で 赤★（2026-09-19）
         ⇒★★「台が 悪い」のでは なく ★測り道具が 本番と 違った★★★
         ⇒記憶「★台が 違うと 言う 前に 測り道具が その式を 押せるかを 見る★」
         ★★左上だけ 返します★★（この 道具は 1マスの 字を 返す 約束）
           ＝溢れた先は `docs/measured/osu-kobore-zenmasu.mjs` が 見ます */
    if (v.溢れ === true) {
      const 頭 = (v.並び && v.並び[0]) ? v.並び[0][0] : null;
      if (頭 === null || 頭 === undefined) return null;
      if (頭 && 頭.型 === '誤' && String(頭.値) === '#NAME?') return null;
      return (頭 && 頭.値 !== undefined) ? String(頭.値) : String(頭);
    }
    if (v.型 === '誤' && String(v.値) === '#NAME?') return null;
    return 板.字('ZZ9999');
  } catch (e) { return null; }
}

/**
 * ★★本番の 道を 建てる★★
 * @param {{XML?:boolean, 外へ出す?:object, EFの道?:string, 板の名?:string}} 注文
 *    XML     … `jsdom` が 在れば FILTERXML に 渡す（既定 true／無ければ null の まま）
 *    EFの道  … `exally-formula.js` の 道（既定＝本番の 物／★写しを 押す 道具が 使う★）
 *    板の名  … 板の 名前（既定 `S`／★`kansuu46` の 紙は `Sheet1` で 取られて います★）
 *    外へ出す … ★既定は「出さない」★（`取る`／`聞く` が 投げる）。
 *               ★本当に 外へ 出す 物を 渡すのは ★お金と 秘密と 相手の 迷惑★が 掛かります★
 */
export async function 建てる(注文) {
  const 注 = 注文 || {};
  const HFns = require_(path.join(ROOT, 'hyperformula.full.min.js'));
  /* ★どの `exally-formula.js` を 読むか★（既定＝本番の 物）
     ＝★丸めの 前後を 比べる 道具★が 仮置きの 写しを 渡します
     ★別の 建て方を 作らず 口の 引数に する★（指示役1 の 決め・2026-09-15） */
  const EF = require_(注.EFの道 || path.join(ROOT, 'exally-formula.js'));
  EF.registerExallyFunctions(HFns);
  const HF0 = HFns.HyperFormula;
  const H = Object.assign(Object.create(HF0), HFns,
    { registerFunctionPlugin: HF0.registerFunctionPlugin.bind(HF0) });

  /* ══ ★プラグインを 本番と 同じ 並びで つなぐ★ ══ */
  let 積んだ = 0;
  /* ★1本ごとの 戻り（いくつ 足したか）★＝★呼ぶ側が 出す 紙に 使う★ */
  const 積 = [];
  for (const n of ['extra', 'nokori', 'kane']) {
    積.push(require_(path.join(ROOT, 'lib/formula-' + n + '-plug.js'))
      .つなぐ(H, require_(path.join(ROOT, 'lib/formula-' + n + '.js'))));
    積んだ++;
  }
  /* ★予測（TREND / GROWTH / LOGEST）★ */
  積.push(require_(path.join(ROOT, 'lib/formula-yosoku-plug.js'))
    .つなぐ(H, require_(path.join(ROOT, 'lib/formula-yosoku.js')),
      () => ({ シート数: 1, 版: 'Exally', 台: 'win', OS: '', 左上: '$A$1' })));
  積んだ++;
  /* ★網の 外へ 出る 物は ★出させない★（司さんの 決め）★
     ＝★試験で 外へ 出すと お金・秘密・相手の 迷惑の 3つとも 掛かります★ */
  積.push(require_(path.join(ROOT, 'lib/formula-soto-plug.js'))
    .つなぐ(H, require_(path.join(ROOT, 'lib/formula-soto.js')), 注.外へ出す || {
      取る: async () => { throw new Error('外へ 出ません'); },
      聞く: async () => { throw new Error('AI に 聞きません'); },
      再計算: () => {},
    }));
  積んだ++;
  /* ★FILTERXML（`jsdom` が 無ければ null の まま）★ */
  let XML部品 = null;
  if (注.XML !== false) {
    try {
      const { JSDOM } = require_('jsdom');
      const w = new JSDOM('').window;
      XML部品 = { DOMParser: w.DOMParser, XPathResult: w.XPathResult };
    } catch (e) { XML部品 = null; }
  }
  積.push(require_(path.join(ROOT, 'lib/formula-filterxml-plug.js'))
    .つなぐ(H, require_(path.join(ROOT, 'lib/formula-filterxml.js')), () => XML部品));
  積んだ++;
  積.push(require_(path.join(ROOT, 'lib/formula-cell-plug.js'))
    .つなぐ(H, require_(path.join(ROOT, 'lib/formula-cell.js')), null));
  積んだ++;
  積.push(require_(path.join(ROOT, 'lib/formula-complex-plug.js'))
    .つなぐ(H, require_(path.join(ROOT, 'lib/formula-complex.js'))));
  積んだ++;

  /* ★★本番と 同じ 本数か（★落としたら ここで 止まる★）★★ */
  const 要る = 本番のプラグイン数();
  if (要る !== 積んだ) {
    throw new Error('★book.html は ' + 要る + '本 読むのに ここでは ' + 積んだ + '本しか つないで いない★');
  }

  /* ★★本番と 同じ 建て方★★（2026-09-09 に 直した）
     ★`smartRounding` が 既定（true）だと エンジンが 答えを 勝手に 丸めます★
       803.6538461538445 → ★803.65384615★＝★桁が 落ちる★
     ⇒★本番に 無い 丸めを 測り台が 足す＝嘘の 数字★ */
  const hf = HF0.buildEmpty({
    licenseKey: 'gpl-v3', useArrayArithmetic: true, smartRounding: false,
    maxRows: 1048576, maxColumns: 18278,
  });
  { /* ★本番の 建て方と 食い違ったら 止める★（★数は book.html から 読む★） */
    const 本 = 本番の建て方();
    for (const 要 of ['useArrayArithmetic:true', 'smartRounding:false']) {
      if (本.indexOf(要) < 0) {
        throw new Error('★本番の buildEmpty に ' + 要 + ' が 無い＝建て方を 見直して ください★');
      }
    }
  }

  /* ★板の 名前も 口の 引数に する★（2026-09-15）
     ＝★`CELL("filename")` の ような 物は ★板の 名前で 答えが 変わる★
     ＝★違いは 別の 道では なく 口の 引数に する★（指示役1 の 決め） */
  const SID = hf.getSheetId(hf.addSheet(注.板の名 || 'S'));
  EF.initExallyFormula(hf);
  /* ★★台（自前の 計算）も 建てます★★（2026-09-18・㋑⑶の 後）
       ★本番の 道は ★3段★に なりました★
         ①JS層（`_jsComputeFormula`） → ②★台★（`lib/shiki-hyou.js`） → ③借り物
       ＝`book.html:setCellFormula` ／ `book.html:recalcSheet` の 2か所（実物）
       ★★この 段を 飛ばすと 「うちが 負けて いる」向きの 偽の 答えが 出ます★★
         ＝この 紙の 頭に 書いて ある 事が ★そのまま 起きました★
         ＝`osu-ramuda-honban.mjs` が `#ERROR!` を 出し、実物は `#VALUE!` でした
       ★台を 使うかは 呼ぶ側が 決めます★＝★古い 道具を 黙って 変えない★ */
  let 板 = null;
  try {
    const SH = require_(path.join(ROOT, 'lib/shiki-hyou.js'));
    板 = SH.表();
  } catch (e) { 板 = null; }
  return { HFns, HF0, H, EF, hf, SID, 積んだ, 積, 板 };
}
