/* zairyou.mjs — ★紙の 材料を 読んで 表に 置く★（部品・2026-09-14）
 *
 *  ★★なぜ 1か所に するか★★
 *    2026-09-14 の 1日で ★同じ型の 誤りを 4回★ 踏みました。全部 元は 同じ：
 *      ★紙の 中身を 機械が 読めない★
 *        ①`mae-ato` の 判定の 字を 実測と 読みかけた
 *        ②`marume-mae-ato` の 3列目（★うちの 古い 答え★）を 実Excel と 読みかけた
 *        ③`golden-oufuku` は ★式の 次も 式★（答えは 5列目）
 *        ④★`golden-shikaku-kansuu` の 材料が 文章でしか 無い★
 *          ⇒ 道具が ★空の 表★で 押し、`=PRODUCT(A1:A5)` が 0 に なり
 *            ★「土台が 壊れている」と 報告される★所だった（★材料を 置けば 2 で 合う★）
 *    ⇒★材料の 読み方・置き方を 道具ごとに 書かない★
 *
 *  ★★紙の 形★★  `#材料<タブ>マス<タブ>値<タブ>型`
 *    ★型の 列（4列目）★ … `数` `字` `真偽` `式` `空`
 *      書かない 時は こう 見ます … ★3列目が 空なら 空／数に なるなら 数／それ以外は 字★
 *    例）
 *      `#材料	A1	1`              … 数の 1
 *      `#材料	A2	2	字`          … ★字の "2"★（★数の 2 では ない★）
 *      `#材料	A3	TRUE	真偽`     … 真偽
 *      `#材料	A4`                  … 空
 *      `#材料	A6	=1/0	式`       … 式（誤りを 置く 時にも 使う）
 *
 *    ★★なぜ 頭の 印（'2 や =1/0）では なく 型の 列に したか★★
 *      指示役1 2026-09-14「★推し量る 形は いつか 本物の 字と ぶつかる★」
 *      ＝★本当に `=` や `'` で 始まる 字★が 材料に 出た 日に 壊れる。
 *      ★型の 列なら ぶつかりません★。
 *      （この 決めを した 時点で `#材料` が 在る 紙は 4本・中身は 全部 数＝★揃え直しは 0★）
 *
 *  ★★使う 側は 必ず「材料が 在るか」を 見る事★★
 *    材料が 無い 紙で マスを 指す 式を 押すと ★空の 表の 答え★に なります。
 *    ⇒ それを 実Excel と 比べると ★偽の 負け★に なる（上の ④）。
 *    ⇒ `押してよいか()` を 使う。
 */
import fs from 'node:fs';

const 型たち = ['数', '字', '真偽', '式', '空'];

/** 紙から 材料を 読む … { A1:{値:'1',型:'数'}, … } */
export function 材料を読む(紙の道) {
  const 出 = {};
  let 中身 = '';
  try { 中身 = fs.readFileSync(紙の道, 'utf8'); } catch (e) { return 出; }
  for (const l of 中身.split(/\r?\n/)) {
    if (!l.startsWith('#材料')) continue;
    const c = l.split('\t');
    if (c.length < 2) continue;
    const マス = c[1].trim();
    if (!マス) continue;
    const 生 = c.length >= 3 ? c[2] : '';
    let 型 = c.length >= 4 ? (c[3] || '').trim() : '';
    if (型 && 型たち.indexOf(型) < 0) {
      throw new Error('★知らない 型★ `' + 型 + '`（' + マス + '）／使えるのは ' + 型たち.join(' '));
    }
    if (!型) {
      /* ★型を 書かない 時の 見方★＝空／数／字 */
      if (生 === '') 型 = '空';
      else if (生.trim() !== '' && isFinite(Number(生))) 型 = '数';
      else 型 = '字';
    }
    出[マス] = { 値: 生, 型: 型 };
  }
  return 出;
}

/** 表に 置く（★型の 通りに★） */
export function 表に置く(表, 材料) {
  for (const マス of Object.keys(材料)) {
    const { 値, 型 } = 材料[マス];
    if (型 === '式') { 表.打つ(マス, 値); continue; }
    if (型 === '空') { 表.打つ(マス, ''); continue; }
    if (型 === '数') { 表.打つ(マス, 値); continue; }
    /* ★型つきの 口★（土台⑲・2026-09-15）
       ★前は `表.中身[マス].値` を 直に 書き換えて いました★
       ＝★その マスを 見て いる 式に 報せが 行きません★
         （★材料を 先に 置く 間は 当たって いませんが、
           順を 変えた 途端に 黙って 古い 答えに なります★）
       ★`置く` は 字を 字のまま 置き、見て いる 式を 直します★ */
    if (型 === '真偽') {
      表.置く(マス, String(値).trim().toUpperCase() === 'TRUE', String(値));
      continue;
    }
    /* ★字★＝`打つ('A2','2')` は ★数の 2★に なって しまうので `置く` で 字として */
    表.置く(マス, String(値), String(値));
  }
  return 表;
}

/* ★★材料が 要らない 関数★★（2026-09-14）
   ＝★マスを 指して いても 中身を 見ない★ 関数。材料が 無くても 答えが 変わりません。
   ★「マス参照が 在る＝材料が 要る」と 決め打ちすると 数を 大きく 見誤ります★
     （2026-09-14 実測＝参照つき 1,216行 のうち ★少なくとも 496行は 要らない★）
   ★裏づけ★
     CELL   … ★材料なしで 108/108 一致★（`kurabe-cell-zenbu` ／ 既定の 見た目で 答える）
     AREAS  … 「★いくつの 四角か★」を 数えるだけ＝中身を 見ない
     ISREF  … 「参照か どうか」だけ
     ROW/COLUMN/ROWS/COLUMNS … 番地・大きさだけ
     ISBLANK は ★入れません★＝★空か どうかは 中身の 話★
   ★足す 時は 訳を 1行 書く事★（★書けない なら まだ 分かって いない★） */
const 材料が要らない関数 = [
  'AREAS',       /* いくつの 四角かを 数えるだけ */
  'ISREF',       /* 参照か どうかだけ */
  'ROW', 'COLUMN', 'ROWS', 'COLUMNS',   /* 番地・大きさだけ */
  'ADDRESS',     /* 番地の 字を 作るだけ */
  'SHEET', 'SHEETS',                    /* 板の 番号・数だけ */
];

/* ★★CELL は 関数の 名前だけでは 決まりません★★（2026-09-14・指示役1 が 止めた）
   ★第1引数（何を 聞くか）で 変わります★：
     ★中身を 見る★ … `contents`（中身そのもの）／`type`（空 b・字 l・数 v）／
                      `prefix`（字の 寄せの 印＝★中身が 字でないと 出ない★）
     ★中身を 見ない★ … parentheses color format width row col address protect filename
   ★私の 誤り★＝「材料なしで 108/108 一致 ⇒ CELL は 材料が 要らない」と 書いた。
     ⇒★空のマスで 押して、実Excel も うちも 同じ 答えに なった だけ★かもしれない
       （＝経営者1 が 同じ日に 取り下げた `=PRODUCT(B1:B3)`（両方 空で 偶然 一致）と ★同じ型★）
     ⇒★材料を 置いて 押し直すまで「要らない」とは 言えません★
   ★知らない 引数が 来たら「要る」側★＝★多めに 倒す★ */
const CELLの材料が要らない第1引数 = [
  'parentheses', 'color', 'format', 'width', 'row', 'col', 'address', 'protect', 'filename',
];

/** `=CELL("row",B7)` の `row` を 取る（取れなければ ''） */
export function CELLの第1引数(式) {
  const m = /^=\s*CELL\s*\(\s*"([^"]*)"/i.exec(String(式).trim());
  return m ? m[1].trim().toLowerCase() : '';
}

/** その 式の 一番 外の 関数（`=CELL("row",B7)` → `CELL`） */
export function 外の関数(式) {
  const m = /^=\s*([A-Z][A-Z0-9._]*)\s*\(/.exec(String(式).trim());
  return m ? m[1] : '';
}

/** その 式が マスを 指して いるか（A1／A1:B2／$A$1 …）
 *  ★後読み（(?<= や (?<!）は 使いません★＝★WebKit（iPhone の Safari）が 読めない★
 *    `tests/no-lookbehind.test.mjs` が 見張って います（2026-09-14 に 1つ 足して 赤に した）。
 *  ⇒ ★手前の 1文字を 組で 取る★（`(^|[^…])`）＝同じ事が 出来ます。 */
export function マスを指すか(式) {
  return /(^|[^A-Z0-9_."])\$?[A-Z]{1,3}\$?[0-9]{1,5}(?![0-9(])/.test(String(式));
}

/**
 * ★押してよいか を 決める★
 *   材料が 無い 紙で マスを 指す 式は ★押さない★（★偽の 負けを 作らない★）
 *   @returns {{よい:boolean, なぜ:string}}
 */
export function 押してよいか(式, 材料) {
  if (!マスを指すか(式)) return { よい: true, なぜ: '' };
  if (Object.keys(材料).length) return { よい: true, なぜ: '' };
  const 外 = 外の関数(式);
  if (外 === 'CELL') {
    const 何 = CELLの第1引数(式);
    if (何 && CELLの材料が要らない第1引数.indexOf(何) >= 0) {
      return { よい: true, なぜ: '★CELL("' + 何 + '") は 中身を 見ない★' };
    }
    /* contents / type / prefix ／ 読めない 引数 … ★要る★（多めに 倒す） */
    return { よい: false, なぜ: '★CELL("' + (何 || '？') + '") は 中身を 見る★（材料が 要る）' };
  }
  if (外 && 材料が要らない関数.indexOf(外) >= 0) {
    return { よい: true, なぜ: '★' + 外 + ' は 中身を 見ない★（材料が 無くても 答えは 同じ）' };
  }
  return { よい: false, なぜ: '★紙に 材料が 無い★（マスを 指す 式なので 押さない）' };
}

export const 材料が要らない = 材料が要らない関数.slice();

export const 使える型 = 型たち.slice();

/* ══ ★紙の「答えは 何列目か」を 決める★（2026-09-14） ══
   ★当て推量で 決めない★＝2026-09-14 に ★4本の 紙★で「式の 次」が 答えでは 無かった
     golden-afure（式1・答3）／golden-excel-kara（式2・答5）／
     golden-sashikomi（式1・答5）／golden-tsunagi（式1・答3）
     ＋ golden-oufuku（★式の 次も 式★・答えは 5列目）
     ＋ golden-marume-mae-ato（式の 次は ★うちの 古い 答え★）
   ★決め方（上から 順に）★
     ①★見出し行に「答」の 列が 在れば その 列★（「実Excel」を 含む 物を 先に）… 一番 確か
     ②型の 列（String/Double/Boolean/error値…）の ★1つ 手前★
     ③式の 次（★当て推量★） */
const 型らしい = /^(String|Double|Number|Boolean|error値|Empty|Date|DateTime|Currency|Long)$/;

/** 紙の 生の 行から 見出しの「答」の 列を 探す（無ければ -1） */
export function 見出しの答えの列(生) {
  let 見 = null;
  for (let i = 0; i < 生.length; i++) {
    const l = 生[i];
    if (!l) continue;
    if (l.startsWith('#')) { if (l.includes('\t')) 見 = l.replace(/^#\s*/, ''); continue; }
    if (i === 0 && l.includes('\t') && !/^=/.test(l.split('\t')[0])) 見 = 見 || l;
    break;
  }
  if (!見) return -1;
  const h = 見.split('\t').map((x) => x.trim());
  const 実 = h.findIndex((x) => /答/.test(x) && /実Excel/.test(x));
  return 実 >= 0 ? 実 : h.findIndex((x) => /答/.test(x));
}

/** その 行の 実Excel の 答えと 型を 取る */
export function 行の答え(列たち, i式, 見出しの答え) {
  let i型 = -1;
  for (let i = i式 + 1; i < 列たち.length; i++) {
    if (型らしい.test((列たち[i] || '').trim())) { i型 = i; break; }
  }
  const i答 = 見出しの答え >= 0 ? 見出しの答え : (i型 > i式 ? i型 - 1 : i式 + 1);
  return {
    答: (列たち[i答] === undefined ? '' : String(列たち[i答]).trim()),
    型: i型 > i式 ? (列たち[i型] || '').trim() : '',
    決め方: 見出しの答え >= 0 ? '見出し' : (i型 > i式 ? '型の手前' : '式の次'),
  };
}

/** ★紙が「判じられない」と 言って いる 行か★（＝真値では ない）
 *  紙は `—` や `★判じられない★` `保留` と 書きます。
 *  ★これを 突き合わせに 使うと「両方 負け」が 水増しされます★
 *  （2026-09-14 実測＝`golden-hoyuu-27` の `CELL("width",D1)` は 紙が `—`。
 *    うちも 借り物も 8 で 合って いるのに ★両方 負け★に 数えて いた） */
export function 真値か(答) {
  const a = String(答 === undefined || 答 === null ? '' : 答).trim();
  if (a === '' || a === '—' || a === '-' || a === '―') return false;
  return !/判じられない|保留|測れない|土台が 違う/.test(a);
}
