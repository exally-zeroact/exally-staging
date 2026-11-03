/* view-extras.test.mjs — ★表示｜表示 の 組（ルーラー／ナビゲーション／セルにフォーカス）★ 2026-08-30
 *
 *  ★なぜ★
 *    2026-08-30 に 実Excel（16.0 build 20326）を 測り直したら
 *    ★「表示｜表示」という 組が 正本から まるごと 抜けていた★。
 *    抜けた 訳は ★自分の 絞り込み★（組の名前が タブ名と 同じ物を 捨てていた）。
 *    ⇒ ★目盛線・見出し・数式バー・ルーラー・ナビゲーション・セルにフォーカス★が 無かった。
 *
 *  ★見る物★
 *    ① ルーラーの 目盛り … ★1cm = 96/2.54 px★（CSSの 決まり）で 合っているか
 *    ② 倍率を 変えたら 目盛りの 間も 変わるか
 *    ③ ナビゲーションの 一覧 … シート／名前／テーブルを 落とさず 出すか
 *    ④ セルにフォーカスの 帯 … ★選んだ範囲では なく 今の 1つのセル★を 見るか
 *    ⑤ 正本に 7個 在り、6個 結んである（データ型アイコンだけ まだ）
 *
 *  走らせ方: node tests/view-extras.test.mjs  ／  --self-test
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
let ok = 0, ng = 0;
const 言う = (よい, 文, 添え) => {
  if (よい) { ok++; console.log('  ok   ' + 文); }
  else { ng++; console.log('  NG   ' + 文); if (添え) console.log('       ' + 添え); }
};
const 近い = (a, b, 幅) => Math.abs(a - b) <= (幅 === undefined ? 0.001 : 幅);

const VX = (await import('file://' + path.join(ROOT, 'lib/view-extras.js').replace(/\\/g, '/'))).default;

console.log('★表示｜表示 の 組★');

/* ── ① ルーラー ───────────────────────────── */
言う(近い(VX.CM, 96 / 2.54), '★1cm = 96/2.54 px（CSSの 決まり）★ … ' + VX.CM.toFixed(4));
const 目 = VX.目盛り(400, 1);
言う(目.length > 0, '★目盛りが 出る★（' + 目.length + '本）');
言う(目[0].x === 0 && 目[0].大きい === true && 目[0].札 === '0', '★0の 所から 始まる★');
const 大きい = 目.filter((t) => t.大きい);
言う(近い(大きい[1].x - 大きい[0].x, VX.CM, 0.01),
  '★1cmごとに 大きい 目盛り＝' + (大きい[1].x - 大きい[0].x).toFixed(2) + 'px★');
言う(目[1].大きい === false && 目[1].札 === '',
  '★0.5cm の 所は 小さい 目盛り（札なし）★');
言う(大きい[1].札 === '1' && 大きい[2].札 === '2', '★札は 1, 2, … と 増える★',
  大きい.slice(0, 3).map((t) => t.札).join(','));
言う(目[目.length - 1].x <= 400, '★はみ出さない★', String(目[目.length - 1].x));

/* ② 倍率 */
const 倍 = VX.目盛り(400, 2).filter((t) => t.大きい);
言う(近い(倍[1].x - 倍[0].x, VX.CM * 2, 0.01),
  '★2倍に したら 目盛りの 間も 2倍★＝' + (倍[1].x - 倍[0].x).toFixed(2) + 'px');
言う(VX.目盛り(0, 1).length === 0 && VX.目盛り(-5, 1).length === 0,
  '★幅0や 負でも 落ちない（空を 返す）★');

/* ── ③ ナビゲーション ─────────────────────── */
const 章 = VX.一覧({
  sheets: [{ name: 'Sheet1' }, { name: '売上' }],
  activeSheet: 1,
  names: [{ name: '税率', ref: 'Sheet1!$A$1' }],
  tables: [{ name: '表1', ref: '売上!$A$1:$C$9' }]
});
言う(章.length === 3, '★章は 3つ（シート／名前／テーブル）★', String(章.length));
言う(章[0].行.length === 2 && 章[0].行[1].印 === '今',
  '★今の シートに 印が 付く★', JSON.stringify(章[0].行));
言う(章[1].行[0].名 === '税率' && 章[1].行[0].印 === 'Sheet1!$A$1', '★名前も 出る★');
言う(章[2].行[0].名 === '表1', '★テーブルも 出る★');
言う(VX.一覧の数(章) === 4, '★数え方が 合う（2+1+1=4）★', String(VX.一覧の数(章)));
const 空章 = VX.一覧({ sheets: [], activeSheet: 0 });
言う(VX.一覧の数(空章) === 0, '★空でも 落ちない★');

/* ── ④ セルにフォーカス ───────────────────── */
言う(JSON.stringify(VX.帯({ r: 3, c: 5 })) === JSON.stringify({ 行: 3, 列: 5 }), '★今の セルの 行と 列を 返す★');
言う(VX.帯(null) === null && VX.帯({ r: -1, c: 0 }) === null, '★変な 値では 塗らない★');

/* ── ⑤ 正本と 結び ───────────────────────── */
const NL = String.fromCharCode(10), TAB = String.fromCharCode(9);
const 正本 = fs.readFileSync(path.join(ROOT, 'docs/excel-ribbon-flat.tsv'), 'utf8')
  .split(NL).map((l) => l.replace(String.fromCharCode(13), '')).filter((l) => l.trim())
  .map((l) => l.split(TAB));
const 表示表示 = 正本.filter((c) => c[0] === '表示' && c[1] === '表示');
言う(表示表示.length === 8,
  '★正本に 表示｜表示 が 8個 在る（7個＋↘）★', String(表示表示.length) + ' … '
    + 表示表示.map((c) => c[2]).join(' / '));
for (const 名 of ['ルーラー', '目盛線', '数式バー', '見出し', 'ナビゲーション', 'セルにフォーカス', 'データ型アイコン']) {
  言う(表示表示.some((c) => c[2] === 名), '  正本に 在る … ' + 名);
}

const spec = fs.readFileSync(path.join(ROOT, 'lib/ribbon-spec.js'), 'utf8');
const 結んだ = ['ルーラー', '目盛線', '数式バー', '見出し', 'ナビゲーション', 'セルにフォーカス']
  .filter((名) => new RegExp("t: '表示', g: '表示', p: '" + 名 + "', a: \\{").test(spec));
言う(結んだ.length === 6, '★6個 結んである★（' + 結んだ.join(' / ') + '）', String(結んだ.length));
/* ★データ型アイコンは まだ★＝Excelの「データの種類（株式・地理）」が 要る。
   その データの種類 自体が ★未着手★なので、ここも まだ。★黙って 消さない★。 */
言う(!/t: '表示', g: '表示', p: 'データ型アイコン', a: \{/.test(spec),
  '★データ型アイコンは まだ 結んでいない（データの種類が 未着手だから）★');

/* ── 画面の 側も 在るか（呼び先が 本当に 在る） ── */
const book = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
for (const f of ['ルーラーを出すか', 'ナビゲーションを開く', 'セルにフォーカスを切り替える']) {
  言う(new RegExp('function\\s+' + f + '\\s*\\(').test(book), '★画面に ' + f + ' が 在る★');
}
言う(/lib\/view-extras\.js/.test(book), '★book.html が view-extras.js を 読んでいる★');
const act = fs.readFileSync(path.join(ROOT, 'lib/ribbon-actions.js'), 'utf8');
for (const a of ['ルーラー', 'ナビゲーション', 'セルにフォーカス']) {
  言う(new RegExp(a + ':\\s*function').test(act), '★動作の層に ' + a + ' が 在る★');
}

/* ── わざと 壊して 赤に なるか ────────────────── */
if (process.argv.includes('--self-test')) {
  console.log('\n★わざと 壊して 赤に なるか★');
  const にせ = VX.目盛り(400, 1).map((t) => ({ ...t }));
  にせ[1].大きい = true;                       /* 0.5cm を 大きい に して しまう */
  const 大2 = にせ.filter((t) => t.大きい);
  言う(!近い(大2[1].x - 大2[0].x, VX.CM, 0.01),
    '★0.5cm を 大きいに したら 1cm の 検査が 落ちる★（' + (大2[1].x - 大2[0].x).toFixed(2) + 'px）');
  const 章に = VX.一覧({ sheets: [{ name: 'A' }], activeSheet: 9 });
  言う(章に[0].行[0].印 === '', '★今の シートが 居なければ 印は 付かない★');
}

console.log('\nview-extras: ' + ok + '/' + (ok + ng) + ' passed');
process.exit(ng ? 1 : 0);
