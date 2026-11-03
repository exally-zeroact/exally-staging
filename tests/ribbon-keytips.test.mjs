/* ribbon-keytips.test.mjs — ★Alt の キーが 実Excel と 同じ順で 効く★ 2026-08-31
 *
 *  ★なぜ★（司さん 2026-08-30）
 *    「Excel を 細胞分解レベルまで 網羅して 把握した上で 持ち込み パクる」
 *    ★毎日 Excel を 使う人ほど Alt の 順で 打つ★（Alt,H,1＝太字／Alt,H,A,L＝左揃え）
 *    実測＝実Excel 462個 vs うち ★0個★ だった。
 *
 *  ★この 見張りが 見る物★
 *    ① キーは ★実Excel を 測った tsv から 起こした 物★（手で 書いていない）
 *    ② ★Alt,H,1＝太字★／★Alt,H,A,L＝左揃え★ が 実際に その 動作へ 届く
 *    ③ ★つないでいない 物には キーを 付けない★（押しても 何も 起きないボタンを 作らない）
 *    ④ ★同じ キーを 違う名前が 名乗ったら 付けない★
 *       （実Excel 自体が 描画の ペン10個に 同じ `Alt, J I, G` を 付けている＝画廊のキー。
 *         1つ 選ぶと ★黒のペンの つもりが 消しゴム★に なる）
 *    ⑤ ★数が 減っていない★
 *    ⑥ ★画面側の 繋ぎ★＝capture で 受けている（＝セルに H と 入らない）
 *
 *  走らせ方: node tests/ribbon-keytips.test.mjs  ／  --self-test
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

const 鍵 = require_(path.join(ROOT, 'lib/ribbon-keys.js'));
const 木部 = require_(path.join(ROOT, 'lib/ribbon-keytips.js'));
const 品 = require_(path.join(ROOT, 'lib/ribbon-spec.js'));
const 範囲 = require_(path.join(ROOT, 'lib/ribbon-scope.js'));
const 動作 = require_(path.join(ROOT, 'lib/ribbon-actions.js'));

console.log('★Alt の キー（実Excel の 順）★');

/* ── ① 手で 書いていない（tsv から 起こした 物） ── */
const 生 = path.join(ROOT, 'docs/excel-keys-2026-08-30.tsv');
言う(fs.existsSync(生), '★元の 実測が 在る（docs/excel-keys-2026-08-30.tsv）★');
const 字 = fs.readFileSync(path.join(ROOT, 'lib/ribbon-keys.js'), 'utf8');
言う(/自動生成/.test(字) && /excel-keys-2026-08-30\.tsv/.test(字),
  '★キーの 表は 自動生成（元が 書いてある）★',
  '★手で 書くと 実Excel と ずれる★');
言う(Object.keys(鍵.タブ).length === 11,
  '★タブの 入口 11個★（今 ' + Object.keys(鍵.タブ).length + '個）');

/* ── ② 実Excel と 同じ順で 届くか ── */
const 木 = 木部.作る(鍵, 品.ITEMS, 範囲);
function 打つ(順) {
  let 節 = 木, 最後 = null;
  for (const c of 順) { const r = 木部.進む(節, c); if (!r.節) return null; 節 = r.節; 最後 = r; }
  return 最後;
}
const 覚えている手 = [
  { 順: ['H'], どうする: 'タブ', 先: 'ホーム' },
  { 順: ['H', '1'], どうする: '押す', 先: '太字' },
  { 順: ['H', '2'], どうする: '押す', 先: '斜体' },
  { 順: ['H', 'A', 'L'], どうする: '押す', 先: '左揃え' },
  { 順: ['H', 'C'], どうする: '押す', 先: 'コピー' },
  { 順: ['H', 'X'], どうする: '押す', 先: '切り取り' },
  { 順: ['N'], どうする: 'タブ', 先: '挿入' },
  { 順: ['W'], どうする: 'タブ', 先: '表示' },
];
for (const t of 覚えている手) {
  const r = 打つ(t.順);
  const 着いた = r && (r.節.動作 || r.節.タブ);
  言う(!!r && r.どうする === t.どうする && 着いた === t.先,
    '★Alt,' + t.順.join(',') + ' → ' + t.先 + '★',
    '実際＝' + (r ? r.どうする + ' / ' + 着いた : '外れ'));
}

/* ── ③ つないでいない 物に キーを 付けていないか ── */
const 死んだ = [];
(function 潜る(n, 道) {
  if (!n) return;
  if (n.動作 && typeof 動作[n.動作] !== 'function') 死んだ.push(道 + ' → ' + n.動作);
  for (const c in n.子) if (Object.prototype.hasOwnProperty.call(n.子, c)) 潜る(n.子[c], 道 + ',' + c);
}(木, ''));
言う(死んだ.length === 0, '★キーの 先は 全部 生きている（働きが 在る）★',
  死んだ.slice(0, 8).join(NL + '       '));

/* ── ④ 同じ キーを 違う名前が 名乗ったら 付けない ── */
const 木2 = 木部.作る(鍵, 品.ITEMS, 範囲);
言う(Array.isArray(木2.曖昧), '★決まらない キーを 数えている★');
const 画廊 = (木2.曖昧 || []).find((v) => v.鍵 === 'J,I,G');
言う(!!画廊 && 画廊.数 >= 5,
  '★描画の ペン（実Excel 自体が 同じ キー）は 付けていない★',
  '★1つ 選ぶと 黒のペンの つもりが 消しゴムに なる★');

/* ── ⑤ 数が 減っていない ── */
const 数 = 木部.数える(木);
const 下限 = 205;   /* ★2026-08-31 の 実測＝209★ */
言う(数.動作 >= 下限,
  '★キーで 押せる 部品 ' + 下限 + '個 以上（今 ' + 数.動作 + '個）★',
  '★減った＝どこかで 繋ぎが 切れた★');
言う(数.タブ === 11, '★キーで 行ける タブ 11個（今 ' + 数.タブ + '個）★');

/* ── ⑥ 画面側の 繋ぎ ── */
const book = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
const { 注記を外す } = await import(
  pathToFileURL(path.join(ROOT, 'scripts/lib/chuki.mjs')).href);
const 素 = 注記を外す(book);
言う(/ribbon-keys\.js/.test(素) && /ribbon-keytips\.js/.test(素),
  '★画面が キーの 部品を 読み込んでいる★');
言う(/キーを受ける\([^)]*\)/.test(素), '★画面が キーを 受けている★');
言う(/addEventListener\('keydown', _リボンキー, true\)/.test(素),
  '★一番 外側（capture＝true）で 受けている★',
  '★でないと 打った 字が セルに 入る（Alt,H の H が 入る）★');
言う(/使った[\s\S]{0,80}?preventDefault\(\)[\s\S]{0,40}?stopPropagation\(\)/.test(素),
  '★こちらで 使った 時だけ 止めている★',
  '★何でも 止めると Alt＋＝や Ctrl+B が 効かなくなる★');
const css = fs.readFileSync(path.join(ROOT, 'lib/ribbon.css'), 'utf8');
言う(/\.rb-tip/.test(css) && /pointer-events:\s*none/.test(css),
  '★札の 見た目が 在る／札は 押せない（下の ボタンに 当たる）★');

/* ── ★★★★Alt は 「離した 時」に 出す★★★★ ──
   ★これを 忘れると 何が 壊れるか（2026-08-31 ★自分で 壊した★）★
     押した 瞬間に 札を 出すと、続けて 来る 「=」を キーとして 食べてしまい
     ★Alt＋＝（オートSUM）が 効かなくなった★。実ブラウザで 見つけた。 */
const 偷見 = { 札: 0, keying: false };
function にせの画面() {
  return {
    classList: { add() { 偷見.keying = true; }, remove() { 偷見.keying = false; } },
    querySelectorAll() { return []; },
  };
}
function キーを打つ(el, e) { return 画面.キーを受ける(el, 品, e); }
const 画面 = require_(path.join(ROOT, 'lib/ribbon.js'));
{
  const el = にせの画面();
  画面.キーをやめる(el);
  const a = キーを打つ(el, { type: 'keydown', key: 'Alt' });
  言う(a === false && !画面.キー中か(),
    '★Alt を 押した だけでは 札を 出さない（止めも しない）★',
    '★出すと Alt＋＝の「＝」を 食べる★');
  const b = キーを打つ(el, { type: 'keydown', key: '=' });      /* Alt を 押しながら = */
  言う(b === false, '★Alt＋＝は 下の層へ 通す（オートSUM）★');
  const c = キーを打つ(el, { type: 'keyup', key: 'Alt' });
  言う(c === false && !画面.キー中か(),
    '★途中で 別の キーを 打ったら 札は 出ない★');
  /* 今度は Alt だけ */
  キーを打つ(el, { type: 'keydown', key: 'Alt' });
  const d = キーを打つ(el, { type: 'keyup', key: 'Alt' });
  言う(d === true && 画面.キー中か(), '★Alt を 離したら 札が 出る★');
  /* 打って 進む */
  言う(キーを打つ(el, { type: 'keydown', key: 'h' }) === true, '★H を 打つと こちらで 使う（セルに 入れない）★');
  言う(画面.状態.tab === 'ホーム', '★H で ホームに 変わる★（今 ' + 画面.状態.tab + '）');
  言う(キーを打つ(el, { type: 'keydown', key: 'Escape' }) === true && !画面.キー中か(),
    '★Esc で やめられる★');
  言う(キーを打つ(el, { type: 'keydown', key: 'a' }) === false,
    '★打っていない 時は 何も しない（普通の 入力を 邪魔しない）★');
}
言う(/e\.type === 'keyup'/.test(fs.readFileSync(path.join(ROOT, 'lib/ribbon.js'), 'utf8')),
  '★離した 時を 見ている（keyup）★');
言う(/addEventListener\('keyup', _リボンキー, true\)/.test(素)
  && /addEventListener\('keydown', _リボンキー, true\)/.test(素),
  '★画面は 押した時と 離した時の 両方を 同じ 口で 受けている★');
言う(/ev\.ctrlKey \|\| ev\.metaKey/.test(素),
  '★Ctrl／⌘ は 別の層に 渡している（Ctrl+B を 壊さない）★');

/* ── わざと 壊して 赤に なるか ── */
if (process.argv.includes('--self-test')) {
  console.log('\n★わざと 壊して 赤に なるか★');
  const 空 = 木部.作る(鍵, [], 範囲);
  言う(木部.数える(空).動作 === 0, '★部品を 空に したら 0個に なる（空振りを 見つけられる）★');
  const にせ品 = 品.ITEMS.map((v) => (v.p === '太字' ? { ...v, a: null } : v));
  const 木3 = 木部.作る(鍵, にせ品, 範囲);
  let n = 木3, 着 = null;
  for (const c of ['H', '1']) { const r = 木部.進む(n, c); if (!r.節) { n = null; break; } n = r.節; 着 = r; }
  言う(!着 || 着.どうする !== '押す' || 着.節.動作 !== '太字',
    '★太字の 繋ぎを 外したら Alt,H,1 が 届かなくなる★');
  言う(打つ(['H', '1']).節.動作 === '太字', '★本物は 壊していない★');
  const にせ木 = 木部.作る({ タブ: {}, 部品の鍵: () => null }, 品.ITEMS, 範囲);
  言う(木部.数える(にせ木).動作 === 0, '★キーの 表が 空なら 0個（勝手に 作っていない）★');
}

console.log('\nribbon-keytips: ' + ok + '/' + (ok + ng) + ' passed');
process.exit(ng ? 1 : 0);
