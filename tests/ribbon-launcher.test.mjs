/* ribbon-launcher.test.mjs — ★組の 右下の ↘（起動ツール）★ 2026-08-31
 *
 *  ★★2026-09-03 に 建て替えた＝3種類を 1つの箱に 混ぜていた★★
 *    前は「↘ は 26個（その他のオプション20 ＋ 開く6）」と 数えていた。★誤り★。
 *    ★名前で 数えた★のが 原因。★位置（組の 下端）で 数えるのが 正しい★。
 *
 *  ★実Excel の 実測（座標つき）★ docs/excel-ribbon-remeasure-2026-08-30.tsv
 *      ★↘（起動ツール）★ … 組の 下端（★上=321★）の Button ＝ ★8個★
 *      ★▾（その他のオプション）★ … 組の 中の メニュー（MenuItem）＝ ★↘では ない★
 *      ★▼（コンボの 開く）★     … 入力欄の 右の Button       ＝ ★↘では ない★
 *      （上=309 の 11個＝全タブ共通「リボンの表示オプション」＝どれでも ない）
 *
 *  ★消していない★＝▾ も ▼ も 実Excel の 画面に 在る物。★印を 分けて 残す★。
 *
 *  ★この 見張りが 見る物★
 *    ①★3つに 分けた 数★ … ↘ 8 ／ ▾ 15 ／ ▼ 1（★1本の数字に 混ぜない★）
 *    ②★呼ぶ先が 全部 生きている★（押しても 何も 起きない 印を 出さない）
 *    ③★出さない 物に 理由が 書いてある★
 *    ④ 画面（リボン）が この 表を 使って 描いている
 *    ⑤★狭い画面では 出さない★（組の 名前と ぶつかる）
 *    ⑥★↘ そのものは 組の 中に 出さない★（同じ字が 2つ 並ばない）
 *
 *  ★本物の 確かめは 実ブラウザ★
 *    2026-08-31 に 11タブ 全部 回って ★22個 描画・落ち0★ を 実測した。
 *
 *  走らせ方: node tests/ribbon-launcher.test.mjs  ／  --self-test
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(import.meta.url);
const NL = String.fromCharCode(10);
let ok = 0, ng = 0;
const 言う = (よい, 文, 添え) => {
  if (よい) { ok++; console.log('  ok   ' + 文); }
  else { ng++; console.log('  NG   ' + 文); if (添え) console.log('       ' + 添え); }
};

const 起動 = require_(path.join(ROOT, 'lib/ribbon-launcher.js'));
const 働き = require_(path.join(ROOT, 'lib/ribbon-actions.js'));
const 品 = require_(path.join(ROOT, 'lib/ribbon-spec.js'));

console.log('★組の 右下の ↘（起動ツール）★');

/* ── ① 数 ── */
const 数 = 起動.数える(働き);
/* ★3つに 分けて 数える（★二度と 1本の数字に 混ぜない★）★ */
言う(数.起動ツール === 8, '★↘（起動ツール）は 8組★（今 ' + 数.起動ツール + '組）');
言う(数.その他のオプション === 15, '★▾（その他のオプション）は 15組★（今 ' + 数.その他のオプション + '組）');
言う(数.コンボ === 1, '★▼（コンボの 開く）は 1組★（今 ' + 数.コンボ + '組）');
言う(数.表の全部 === 24, '★8 ＋ 15 ＋ 1 ＝ 24組★（今 ' + 数.表の全部 + '組）');
言う(数.実Excelの起動ツール === 8,
  '★実Excel の ↘ は 8個（組の 下端＝上=321）★（今 ' + 数.実Excelの起動ツール + '個）');
/* ★↘ の 8組が 表と 一致しているか（★片方だけ 直すのを 防ぐ★） */
const 起動の鍵 = Object.keys(起動.表).filter((k) => 起動.表[k].印 === '↘').sort();
const 実の鍵 = 起動.実の起動ツール.map((x) => x.t + '|' + x.g).sort();
言う(起動の鍵.join(',') === 実の鍵.join(','),
  '★表の ↘ と 実測の ↘ が 同じ 8組★', 起動の鍵.join(' / ') + NL + '       vs ' + 実の鍵.join(' / '));

/* ── ② 呼ぶ先が 生きているか ── */
言う(数.死んだ.length === 0, '★呼ぶ先は 全部 生きている（偽の 印を 出さない）★',
  数.死んだ.join(NL + '       '));
言う(数.出せる === 24, '★24組 全部 出せる★（今 ' + 数.出せる + '組）');

/* ── ③ 出さない 物に 理由 ── */
const 理由なし = Object.keys(起動.出さない).filter((k) => !String(起動.出さない[k]).trim());
言う(理由なし.length === 0, '★出さない 物には 理由が 書いてある★', 理由なし.join(' / '));
言う(!!起動.出さない['数式|Python (プレビュー)'],
  '★Python の ↘ は 出さない（Microsoft の クラウドが 要る）★');

/* ── ④ 組の 名前が 実物と 合っているか（★表が 空振りしていないか★） ── */
const 実の組 = {};
for (const it of 品.ITEMS) 実の組[it.t + '|' + it.g] = 1;
const 外れ = Object.keys(起動.表).filter((k) => !実の組[k]);
言う(外れ.length === 0, '★表の 組は 全部 実Excel の 並びに 在る★', 外れ.join(' / '));
const 外れ2 = Object.keys(起動.出さない).filter((k) => !実の組[k]);
言う(外れ2.length === 0, '★出さない 組も 実Excel の 並びに 在る★', 外れ2.join(' / '));

/* ── ⑤ 画面 ── */
const rb = fs.readFileSync(path.join(ROOT, 'lib/ribbon.js'), 'utf8');
const { 注記を外す } = await import(
  pathToFileURL(path.join(ROOT, 'scripts/lib/chuki.mjs')).href);
const rb素 = 注記を外す(rb);
言う(/ribbon-launcher/.test(rb素), '★リボンが 表を 読み込んでいる★');
言う(/起動\.引く\(/.test(rb素), '★リボンは 表から 引いている（画面に 直に 書いていない）★');
{
  言う(/rb-launch/.test(rb素), '★↘ の 印を 描いている★');
}
言う(/'\.rb-item, \.rb-q, \.rb-launch'/.test(rb素),
  '★↘ も 押し込みの 一覧に 入っている★',
  '★入れないと 描いてあるのに 押せない★');
const book = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
言う(/ribbon-launcher\.js/.test(book), '★画面が 部品を 読み込んでいる★');
const css = fs.readFileSync(path.join(ROOT, 'lib/ribbon.css'), 'utf8');
言う(/\.rb-launch/.test(css), '★↘ の 見た目が 在る★');
言う(/max-width:\s*560px[\s\S]{0,200}\.rb-launch[\s\S]{0,60}display:\s*none/.test(css),
  '★狭い画面では 出さない（組の 名前と ぶつかる）★');
言う(/\.rb-gname\s*\{[^}]*position:\s*relative/.test(css),
  '★↘ を 置く 土台が 在る（rb-gname が relative）★');

/* ── わざと 壊して 赤に なるか ── */
if (process.argv.includes('--self-test')) {
  console.log('\n★わざと 壊して 赤に なるか★');
  言う(起動.引く('ホーム', 'フォント', {}) === null,
    '★働きが 無ければ ↘ を 出さない★');
  言う(起動.引く('ホーム', '在るわけない組', 働き) === null,
    '★表に 無い 組には 出さない★');
  const v = 起動.引く('ホーム', 'フォント', 働き);
  言う(!!v && v.先 === 'フォントの設定', '★在る 組は ちゃんと 引ける★');
  言う(起動.数える({}).出せる === 0, '★働きが 空なら 0組（偽の ↘ を 出さない）★');
  言う(起動.数える(働き).出せる === 24, '★本物は 壊していない★');
  /* ★印を 混ぜたら 赤に なるか（今回の 事故そのもの）★ */
  const 印なし = Object.keys(起動.表).filter((k) => !/^[↘▾▼]$/.test(String(起動.表[k].印)));
  言う(印なし.length === 0, '★24組 とも 印（↘/▾/▼）が 付いている★', 印なし.join(' / '));
  言う(起動.起動の品か('ページ レイアウト', 'シートのオプション', 'シートのオプション') === true
    && 起動.起動の品か('ページ レイアウト', 'シートのオプション', '表示') === false,
    '★↘ そのものだけを 組の 中から 外す（他は 外さない）★');
}

console.log('\nribbon-launcher: ' + ok + '/' + (ok + ng) + ' passed');
process.exit(ng ? 1 : 0);
