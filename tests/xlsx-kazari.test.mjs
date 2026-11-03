/* xlsx-kazari.test.mjs — ★マスの 飾りを 自前で 読めて いるか★（2026-09-21）
 *
 *  ★★なぜ 要るか★★
 *    実Excel が 作った 飾り付きの ファイルを お客さんの 道で 開いて 数えたら
 *    ★届いた 7 / 13★ でした（`docs/measured/golden-kazari-gamen-made-2026-09-21.tsv`）。
 *    ★太字・字の色・罫線2つ・塗り★ が ★画面に 1つも 出て いませんでした★。
 *    因 ＝ ★借り物（SheetJS 0.20.3）が 飾りを 捨てる★（同じ日 実測）
 *          `wb.Styles.Borders` は ★10個 とも `{}`★／マスの `cellXf` の 番号も 捨てる
 *    ⇒`lib/xlsx-kazari.js` で ★生の `xl/styles.xml` を 自前で 読みます★
 *
 *  ★★物差し★★
 *    `tests/fixtures/kazari-hiraku3.xlsx`（11639B＝10-10 に 作者の 欄・10-11 に 保存した PC の 道を 空に した・元 11725B）
 *    sha256 affd42e9f5c6de9d5565c8ffd623ffd2e37a7ba5ca8820fdd5c971b15ff93053
 *    ★実Excel（16.0 build 20326）が COM で 作った 物★
 *    作り方 `docs/measured/tsukuru-tameshi-hiraku3-kazari-to-kobore3.ps1`
 *    ★入って いるのは 作り物の 数だけ★（3 / 1 / 0.25 / abc / kazari no tame no musubi）
 *      ＝★司さんの 商売の 中身は 1文字も 在りません★
 *
 *  ★★待つ 物★★（上の ps1 の 92〜101行目 が そのまま 元）
 *    `A1:C3` の 周り ＝ `BorderAround(1, 3)`（LineStyle 1 / Weight 3 ＝ medium）
 *    `B2` の 下       ＝ LineStyle 1 / Weight 4（＝ thick）
 *    `B1` の 塗り     ＝ Interior.Color 65535（＝黄 `#FFFF00`）
 *    `A1` の 字       ＝ Bold ／ Color 255（＝赤 `#FF0000`）
 *
 *  ★★包みは 自前で ほどきます★★
 *    借り物に 読ませると「借り物が どう 読んだか」しか 分かりません。
 *
 *  使い方: node tests/xlsx-kazari.test.mjs
 *          node tests/xlsx-kazari.test.mjs --self-test
 */
import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';
import crypto from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const K = require_(path.join(ROOT, 'lib/xlsx-kazari.js'));

let pass = 0, fail = 0;
const NL = String.fromCharCode(10);
const T = (n, f) => {
  try { f(); pass++; console.log('  ok   ' + n); }
  catch (e) { fail++; console.log('  NG   ' + n + NL + '       ' + e.message); }
};

/* ★包みを ほどく★（局所ヘッダを 順に 舐めて 生の字を 見る） */
function ほどく(buf) {
  const 出 = {};
  let p = 0;
  while (p + 30 <= buf.length) {
    if (buf.readUInt32LE(p) !== 0x04034b50) break;
    const 方 = buf.readUInt16LE(p + 8);
    const 圧 = buf.readUInt32LE(p + 18);
    const n = buf.readUInt16LE(p + 26);
    const x = buf.readUInt16LE(p + 28);
    const 名 = buf.toString('utf8', p + 30, p + 30 + n);
    const 頭 = p + 30 + n + x;
    const 体 = buf.slice(頭, 頭 + 圧);
    出[名] = 方 === 8 ? zlib.inflateRawSync(体).toString('utf8') : 体.toString('utf8');
    p = 頭 + 圧;
  }
  return 出;
}

const 材料道 = path.join(ROOT, 'tests/fixtures/kazari-hiraku3.xlsx');
const 中 = fs.existsSync(材料道) ? fs.readFileSync(材料道) : null;

console.log('');
console.log('[xlsx-kazari] ★マスの 飾りを 自前で 読めて いるか★');

T('★材料が 在る（空振りして いない）★', () => {
  if (!中) throw new Error('★材料が 無い★ ' + 材料道);
  if (中.length !== 11639) throw new Error('★大きさが 違う★ ' + 中.length + 'B（11639B の はず）');
  const h = crypto.createHash('sha256').update(中).digest('hex');
  if (h !== 'affd42e9f5c6de9d5565c8ffd623ffd2e37a7ba5ca8820fdd5c971b15ff93053') {
    throw new Error('★材料が 入れ替わって います★ sha256=' + h);
  }
});

const 部品 = 中 ? ほどく(中) : {};
T('★包みが ほどけた（styles と sheet が 在る）★', () => {
  if (!部品['xl/styles.xml']) throw new Error('★xl/styles.xml が 無い★');
  if (!部品['xl/worksheets/sheet1.xml']) throw new Error('★xl/worksheets/sheet1.xml が 無い★');
});

const 飾り = 中 ? K.飾りを読む(部品['xl/worksheets/sheet1.xml'], 部品['xl/styles.xml']) : {};
console.log('      ＝ 飾りの 在る マス ' + Object.keys(飾り).length + '個 ／ '
  + Object.keys(飾り).join(' '));

T('★★字（A1）が 太字で 赤★★', () => {
  const a = 飾り.A1 || {};
  if (a.bold !== true) throw new Error('★太字が 付いて いない★ ' + JSON.stringify(a));
  if (a.color !== '#FF0000') throw new Error('★色が ' + a.color + '★（#FF0000 の はず）');
});

T('★★塗り（B1）が 黄★★', () => {
  const b = 飾り.B1 || {};
  if (b.bgColor !== '#FFFF00') throw new Error('★塗りが ' + b.bgColor + '★（#FFFF00 の はず）');
});

/* ★★`BorderAround(A1:C3)` は ★周りの 9マスに 散ります★★
     ＝角の マスは 2辺、辺の マスは 1辺。★真ん中（B2）には 外枠が 付きません★
     ⇒★1マスだけ 見ると 「読めて いる」と 言えません★ */
T('★★A1:C3 の 周りの 罫線が 9マスに 正しく 散って いる★★', () => {
  const 待つ = {
    A1: { top: 2, left: 2 }, B1: { top: 2 }, C1: { top: 2, right: 2 },
    A2: { left: 2 }, C2: { right: 2 },
    A3: { bottom: 2, left: 2 }, B3: { bottom: 2 }, C3: { bottom: 2, right: 2 },
  };
  for (const ま of Object.keys(待つ)) {
    const 出 = (飾り[ま] || {}).border || {};
    const あ = JSON.stringify(待つ[ま]);
    const い = JSON.stringify(出);
    if (あ !== い) throw new Error('★' + ま + '★ 待つ ' + あ + ' ／ 出た ' + い);
  }
});

T('★★B2 の 下の 罫線（thick）が 在る★★（★真ん中の マス★）', () => {
  const b = (飾り.B2 || {}).border || {};
  if (b.bottom !== 2) throw new Error('★B2 の 下が ' + b.bottom + '★（2 の はず）／' + JSON.stringify(b));
  if (b.top || b.left || b.right) {
    throw new Error('★B2 に 外枠が 付いて います★ ' + JSON.stringify(b)
      + '／★BorderAround は 真ん中には 付きません★');
  }
});

T('★飾りの 無い マスは 入れない（空の 物を 増やさない）★', () => {
  for (const ま of ['D1', 'E1', 'A10', 'A5']) {
    if (飾り[ま]) throw new Error('★' + ま + ' に 飾りが 付いた★ ' + JSON.stringify(飾り[ま]));
  }
});

/* ══ ★★道具の 一つずつ★★ ══（★口を 名指しで 押す★） */
T('★色にする ＝ 8桁は 頭の 2桁を 落とす／6桁は そのまま／他は null★', () => {
  const 組 = [['FFFF0000', '#FF0000'], ['FFFF00', '#FFFF00'], ['00FFFF00', '#FFFF00'],
    ['', null], ['XYZ', null], ['FFGG0000', null], ['12345', null]];
  for (const [入, 待] of 組) {
    const 出 = K.色にする(入);
    if (出 !== 待) throw new Error('★' + JSON.stringify(入) + ' → ' + JSON.stringify(出)
      + '★（' + JSON.stringify(待) + ' の はず）');
  }
});

T('★太さ ＝ none と 空は 0／medium thick double は 2／thin hair は 1★', () => {
  const 組 = [['', 0], ['none', 0], ['medium', 2], ['thick', 2], ['double', 2],
    ['thin', 1], ['hair', 1], ['dotted', 1], ['mediumDashed', 2]];
  for (const [入, 待] of 組) {
    const 出 = K.太さ(入);
    if (出 !== 待) throw new Error('★' + 入 + ' → ' + 出 + '★（' + 待 + ' の はず）');
  }
});

T('★マスの型番 ＝ `<c r="A1" s="7">` から A1→7 を 取る★', () => {
  const 番 = K.マスの型番(部品['xl/worksheets/sheet1.xml']);
  if (番.A1 !== 7) throw new Error('★A1 の 型番が ' + 番.A1 + '★（7 の はず）');
  if (番.B1 !== 8) throw new Error('★B1 の 型番が ' + 番.B1 + '★（8 の はず）');
  if (番.D1 !== undefined) throw new Error('★D1 に `s=` は 無い はず★ ' + 番.D1);
});

/* ══ ★★テーマの 色★★ ══（2026-09-21）
     `.xlsx` の `styles.xml` は `<color theme="4"/>` ＝ ★番号しか 書いて いません★
     ⇒`xl/theme/theme1.xml` を 引かないと 色が 出ません。
     ★★番号と 紙の 並びは 同じでは ありません★★（経営者1 が 実Excel に 聞いた 物）
       0 ⇒ lt1（白）／ 1 ⇒ dk1（黒）／ 2 ⇒ lt2 ／ 3 ⇒ dk2
       ＝★0と1、2と3が 入れ替わって います★
     ⇒★紙を 上から 数えると 白黒が 逆に なります★＝★ここを 門で 押さえます★ */
const テーマ = K.テーマを読む(部品['xl/theme/theme1.xml']);
console.log('      ＝ テーマの 色 ' + Object.keys(テーマ).length + '個 ／ 0=' + テーマ[0]
  + ' 1=' + テーマ[1] + ' 4=' + テーマ[4]);

T('★★テーマの 色が 実Excel と 合う★★（★0と1が 入れ替わる 所★）', () => {
  const 待つ = { 0: '#FFFFFF', 1: '#000000', 2: '#E8E8E8', 3: '#0E2841',
    4: '#156082', 5: '#E97132', 6: '#196B24', 7: '#0F9ED5', 8: '#A02B93', 9: '#4EA72E' };
  for (const n of Object.keys(待つ)) {
    if (テーマ[n] !== 待つ[n]) {
      throw new Error('★番号 ' + n + ' が ' + テーマ[n] + '★（' + 待つ[n] + ' の はず）');
    }
  }
});

T('★★0番と 1番を 取り違えて いない★★（★紙の 並びは dk1 が 先★）', () => {
  const 字 = 部品['xl/theme/theme1.xml'];
  const i1 = 字.indexOf('<a:dk1>'), i2 = 字.indexOf('<a:lt1>');
  if (!(i1 >= 0 && i2 >= 0 && i1 < i2)) {
    throw new Error('★紙の 並びが 違います★ dk1=' + i1 + ' lt1=' + i2);
  }
  if (テーマ[0] !== '#FFFFFF') {
    throw new Error('★番号 0 が ' + テーマ[0] + '★／★紙の 1番目（lt1・白）の はず★'
      + '／★上から 数えると 黒に なります★');
  }
});

T('★テーマが 無ければ テーマの 色は 付けない（当て推量で 黒を 入れない）★', () => {
  const 無 = K.飾りを読む(部品['xl/worksheets/sheet1.xml'], 部品['xl/styles.xml']);
  const 有 = K.飾りを読む(部品['xl/worksheets/sheet1.xml'], 部品['xl/styles.xml'], 部品['xl/theme/theme1.xml']);
  if (JSON.stringify(無.A1) !== JSON.stringify(有.A1)) {
    throw new Error('★A1 が テーマの 有無で 変わりました★／A1 は 字で 書いた 赤の はず');
  }
});

/* ══ ★★濃さ（tint）★★ ══（2026-09-21 経営者1 が 実Excel に 聞きました）
     紙 `docs/measured/golden-jitsu-excel-iro-no-koide-2026-09-21.tsv`
     ★★9 / 9 ぴったり 合います★★（2026-09-21 実測・★差 0★）
       ⇒`lib/theme.js` の `濃淡()` を 呼んで います。
       ⇒2026-08-30 に ★実測 15/15 で 合わせて あった 台★ です。
     ★★2人とも 1度 「ぴったりでは ない」と 出しました★★
       ＝★どちらも 在る 台を 呼ばず 自分で 書き直した★ から
       私 ... HSL（0〜1 の 実数）で 四捨五入 ⇒ 9個 中 2個が 差 2
       経営者1 ... 同じ 書き直し ＋ ★変換の R と B が 入れ替わって いた★
       正 ... ★Windows の HLS（0〜240 の 整数）で 切り捨て★
       ⇒記憶「★作る前に 探せ★／★在るのに 呼んで いないを 探せ★」
     ★★門は 差 0 です★★（★緩い 門は 間違った 作りを 通します★）
       ＝前は 差 2で 通して いました。★それだと 私の 書き直しが 通ります★。 */
const 濃さの実測 = [
  [-0.9, '#020A0D'], [-0.75, '#05171F'], [-0.5, '#0B3040'], [-0.25, '#104861'],
  [0, '#156082'], [0.25, '#229ACE'], [0.5, '#64BEE6'], [0.75, '#B0DEF2'], [0.9, '#E0F2FA'],
];
const 色の差 = (a, b) => Math.max(
  ...[1, 3, 5].map((i) => Math.abs(parseInt(a.slice(i, i + 2), 16) - parseInt(b.slice(i, i + 2), 16))));

T('★★濃さ（tint）が 実Excel と ぴったり（差 0）★★（9 / 9）', () => {
  const 外 = [];
  for (const [t, 待] of 濃さの実測) {
    const 出 = K.濃さを掛ける('#156082', t);
    const d = 色の差(出, 待);
    if (d > 0) 外.push('t=' + t + ' 出 ' + 出 + ' 実 ' + 待 + ' 差 ' + d);
  }
  if (外.length) throw new Error('★' + 外.length + '個 外れ★' + NL + '       ' + 外.join(NL + '       '));
});

T('★濃さ 0 は 色を 変えない★', () => {
  if (K.濃さを掛ける('#156082', 0) !== '#156082') {
    throw new Error('★0 で 変わりました★ ' + K.濃さを掛ける('#156082', 0));
  }
});

T('★★濃さの 向きを 取り違えて いない★★（★正なら 明るく・負なら 暗く★）', () => {
  const 明 = K.濃さを掛ける('#156082', 0.5);
  const 暗 = K.濃さを掛ける('#156082', -0.5);
  const 明さ = (c) => [1, 3, 5].reduce((a, i) => a + parseInt(c.slice(i, i + 2), 16), 0);
  if (!(明さ(明) > 明さ('#156082') && 明さ(暗) < 明さ('#156082'))) {
    throw new Error('★向きが 逆です★ 明 ' + 明 + ' 元 #156082 暗 ' + 暗);
  }
});

T('★★既定の 字体（0番）の 色は 付けない★★（★黒と 灰が 混ざらない★）', () => {
  /* ★★2026-09-21 ここで 1回 踏みました★★
       テーマの 色を 読む ように したら `<font><color theme="1"/>`（既定の 黒）まで 拾い、
       ★`s=` の 在る マスだけ 黒／無い マスは Exally の 既定★ に なりました。
       ⇒★絵で 見て 気づきました★（数は 9 ⇒ 12 に 増えただけ＝★数では 分かりません★） */
  const 出 = K.飾りを読む(部品['xl/worksheets/sheet1.xml'], 部品['xl/styles.xml'],
    部品['xl/theme/theme1.xml']);
  const 黒 = Object.keys(出).filter((k) => 出[k].color === '#000000');
  if (黒.length) {
    throw new Error('★既定の 黒が ' + 黒.length + 'マスに 付いて います★ ' + 黒.join(' ')
      + '／★Exally が 既に 描いて いる 色なので 付けません★');
  }
  if (出.A1.color !== '#FF0000') {
    throw new Error('★A1 の 赤が 消えました★ ' + 出.A1.color
      + '／★既定を 落とす 時に 本物まで 落として います★');
  }
});

console.log('');
console.log('  ★見て いない 事★');
console.log('    ・★画面に 描いて いるかは ここでは 測って いません★');
console.log('      ＝`docs/measured/hakaru-kazari-ga-gamen-made-todoku-ka.mjs`（画素まで 測る）');
console.log('    ・★濃さ（tint）は `lib/theme.js` の `濃淡()` を 呼んで います★（9/9 差 0）');
console.log('      ＝★自分で 書き直すと 外れます★（2人とも 1度 外しました）');
console.log('    ・★accent1 以外の 色で 同じ 式に なるかは 未測定★（1色でしか 見て いません）');
console.log('      ＝★元の 色だけ 出します★（色味は 合う／明るさが ずれる）');
console.log('    ・番号の 色（`indexed`）は ★読んで いません★');
console.log('    ・テーマを ★別の テーマに 変えた 時★ は 未測定');
console.log('    ・斜めの 罫線は ★台に 持ち方が 有りません★');
console.log('    ・`.xlsb` は ★まだ★（包みの 中が 別物）');
console.log('xlsx-kazari: ' + pass + ' 緑 / ' + fail + ' 赤');

if (process.argv.includes('--self-test')) {
  console.log('');
  console.log('--self-test: ★わざと壊して 赤に なるか★（★lib は 1字も 触りません★）');
  const 板 = 部品['xl/worksheets/sheet1.xml'];
  const 型 = 部品['xl/styles.xml'];
  /* ★★壊し方の 表★★＝★元の 字を 書き換えて 押します★
       ★1つずつ 別の 所を 壊します★＝1か所で 全部 赤に なる 門は 弱い
       ★直すのは 字の 写しだけ★＝`lib` も 材料の ファイルも 触りません */
  const 壊し方 = [
    { 名: '`<b/>` を 抜く（太字が 消える はず）',
      型: (x) => x.replace(new RegExp('<b/>', 'g'), ''),
      見る: (k) => !(k.A1 || {}).bold },
    { 名: '赤の 字を 緑に する（色が 変わる はず）',
      型: (x) => x.replace(new RegExp('FFFF0000', 'g'), 'FF00FF00'),
      見る: (k) => (k.A1 || {}).color === '#00FF00' },
    { 名: 'medium を none に する（罫線が 消える はず）',
      型: (x) => x.replace(new RegExp('style="medium"', 'g'), 'style="none"'),
      見る: (k) => !(k.A1 || {}).border },
    { 名: 'solid を none に する（塗りが 消える はず）',
      型: (x) => x.replace(new RegExp('patternType="solid"', 'g'), 'patternType="none"'),
      見る: (k) => !(k.B1 || {}).bgColor },
    { 名: 'マスの 型番を 全部 0 に する（飾りが 0個に なる はず）',
      板: (x) => x.replace(new RegExp('s="[1-9]"', 'g'), 's="0"'),
      見る: (k) => Object.keys(k).length === 0 },
  ];
  let 悪 = 0;
  for (const こ of 壊し方) {
    const 板2 = こ.板 ? こ.板(板) : 板;
    const 型2 = こ.型 ? こ.型(型) : 型;
    if (板2 === 板 && 型2 === 型) {
      console.log('  NG   ★1字も 変わって いません★ ' + こ.名);
      悪++; continue;
    }
    if (こ.見る(K.飾りを読む(板2, 型2))) console.log('  ok   ' + こ.名);
    else { console.log('  NG   ★壊したのに 同じ 答え★ ' + こ.名); 悪++; }
  }
  process.exit(悪 ? 1 : 0);
}
process.exit(fail ? 1 : 0);
