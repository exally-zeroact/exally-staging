/* theme.test.mjs — ★テーマ（配色・フォント・効果）と 背景・ふりがな★ 2026-08-30
 *
 *  ★真値（実Excel 16.0 で 実測）★ … tools/measure-theme.ps1
 *    12色 … dk1 #000000／lt1 #FFFFFF／dk2 #0E2841／lt2 #E8E8E8／
 *           accent1 #156082／accent2 #E97132／accent3 #196B24／accent4 #0F9ED5／
 *           accent5 #A02B93／accent6 #4EA72E／hlink #467886／folHlink #96607D
 *    フォント … 見出し Aptos Display／本文 Aptos Narrow／セルは 游ゴシック 11
 *    ★濃淡（TintAndShade）★ … 実測 15通り（下の 表）。★0〜240 の HLS で 切り捨て★
 *    背景 … 画面だけ・紙には 刷られない
 *    ふりがな … 大きさ 6／行の 高さ 26.7／★中身から 打った 字には 入らない（Count=0）★
 *
 *  走らせ方: node tests/theme.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const 壊す = process.argv.includes('--self-test');
let 緑 = 0, 赤 = 0;
const ok = (名, 条件, 添え) => {
  if (条件) { 緑++; console.log('  ok   ' + 名); }
  else { 赤++; console.log('  ★NG★ ' + 名 + (添え !== undefined ? '  … ' + 添え : '')); }
};
const book = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
const T = require_(path.join(ROOT, 'lib/theme.js'));

function 抜く(名) {
  const i = book.indexOf('function ' + 名 + '(');
  if (i < 0) return null;
  let d = 0; const j = book.indexOf('{', i);
  for (let k = j; k < book.length; k++) {
    if (book[k] === '{') d++;
    else if (book[k] === '}') { d--; if (d === 0) return book.slice(i, k + 1); }
  }
  return null;
}

/* ★実Excel で 測った 濃淡 15通り★（tools/measure-theme.ps1 と 追加の 実測） */
const 濃淡の実測 = [
  ['#156082', 0.9, '#E0F2FA'], ['#156082', 0.8, '#C0E6F5'], ['#156082', 0.6, '#83CCEB'],
  ['#156082', 0.4, '#44B3E1'], ['#156082', 0.2, '#1F8EBE'], ['#156082', 0.05, '#176C91'],
  ['#156082', -0.05, '#145B7A'], ['#156082', -0.1, '#135673'], ['#156082', -0.25, '#104861'],
  ['#156082', -0.5, '#0B3040'], ['#156082', -0.75, '#05171F'], ['#156082', -0.9, '#020A0D'],
  ['#E97132', 0.4, '#F1A983'], ['#E97132', -0.25, '#BE5014'], ['#E97132', 0.8, '#FBE2D5'],
];

console.log('\n[① 測った 道具が 残っている]');
ok('tools/measure-theme.ps1 が 在る', fs.existsSync(path.join(ROOT, 'tools/measure-theme.ps1')));

console.log('\n[② ★色の 濃淡が 実Excel と 数まで 同じか★（15通り）]');
{
  let 合 = 0;
  const 外 = [];
  for (const [c, t, 期] of 濃淡の実測) {
    const 出 = T.濃淡(c, t);
    if (出 === 期) 合++; else 外.push(c + ' ' + t + ': ' + 出 + ' ≠ ' + 期);
  }
  ok('★15通り 全部 ぴったり★', 合 === 濃淡の実測.length, 合 + '/' + 濃淡の実測.length + '  ' + 外.join(' / '));
  ok('★0 は そのまま★', T.濃淡('#156082', 0) === '#156082', T.濃淡('#156082', 0));
  ok('★+1 で 真っ白★', T.濃淡('#156082', 1) === '#FFFFFF', T.濃淡('#156082', 1));
  ok('★-1 で 真っ黒★', T.濃淡('#156082', -1) === '#000000', T.濃淡('#156082', -1));
  ok('★1 より 大きくても 1 で 止める★', T.濃淡('#156082', 5) === T.濃淡('#156082', 1));
  ok('★-1 より 小さくても -1 で 止める★', T.濃淡('#156082', -5) === T.濃淡('#156082', -1));
  ok('★灰色（彩度0）でも 落ちない★', /^#[0-9A-F]{6}$/.test(T.濃淡('#808080', 0.4)), T.濃淡('#808080', 0.4));
  ok('★白を 濃くできる★', T.濃淡('#FFFFFF', -0.5) !== '#FFFFFF', T.濃淡('#FFFFFF', -0.5));
  ok('★黒を 薄くできる★', T.濃淡('#000000', 0.5) !== '#000000', T.濃淡('#000000', 0.5));
  /* 四捨五入だと 合わない事を 見張る＝★切り捨てで ある事★の 証拠 */
  ok('★切り捨てで ある（四捨五入なら 違う 数に なる 所）★',
    T.濃淡('#156082', 0.4) === '#44B3E1' && T.濃淡('#156082', 0.4) !== '#46B3E1');
  ok('★部品の 中で 切り捨てを 使っている★',
    /Math\.floor\(L \* \(1 - t\) \+ HLSMAX \* t\)/.test(fs.readFileSync(path.join(ROOT, 'lib/theme.js'), 'utf8')));
}

console.log('\n[③ 12の 役割（並びも 実Excel と 同じ）]');
{
  ok('★12個★', T.役割.length === 12, String(T.役割.length));
  ok('★並びは dk1,lt1,dk2,lt2,accent1..6,hlink,folHlink★',
    T.役割.join(',') === 'dk1,lt1,dk2,lt2,accent1,accent2,accent3,accent4,accent5,accent6,hlink,folHlink',
    T.役割.join(','));
  ok('★実Excel の 既定 12色を 書き残している★',
    T.実Excelの既定.length === 12 && T.実Excelの既定[4] === '#156082' && T.実Excelの既定[11] === '#96607D');
  /* ★色は 写さない★＝実Excel の 色が うちの 配色に 入っていない事 */
  const うちの色 = T.配色たち.flatMap(c => c.色.map(v => v.toUpperCase()));
  const 写した = T.実Excelの既定.filter(v => v !== '#000000' && v !== '#FFFFFF')
    .filter(v => うちの色.indexOf(v.toUpperCase()) >= 0);
  ok('★Excel の 色を 写していない（白黒 以外）★', 写した.length === 0, 写した.join(' '));
  for (const c of T.配色たち) {
    ok('  配色「' + c.名 + '」は 12色', c.色.length === 12, String(c.色.length));
    ok('  「' + c.名 + '」は 全部 #RRGGBB', c.色.every(v => /^#[0-9A-Fa-f]{6}$/.test(v)));
  }
  ok('★役割の 名前で 色を 引ける★', T.役割の色(T.配色を引く('みどり（既定）'), 'accent1') === '#3D9E72');
  ok('★無い 役割は null★', T.役割の色(T.配色を引く('みどり（既定）'), 'ないやつ') === null);
  ok('★無い 配色は 1つ目に なる★', T.配色を引く('ない') === T.配色たち[0]);
  ok('★見本の段は 6色★', T.見本の段(T.配色を引く('みどり（既定）'), 'accent1').length === 6);
  ok('★見本の段は 薄い→濃いの 順★', (() => {
    const 段 = T.見本の段(T.配色を引く('みどり（既定）'), 'accent1');
    const 明 = 段.map(v => parseInt(v.slice(1, 3), 16) + parseInt(v.slice(3, 5), 16) + parseInt(v.slice(5, 7), 16));
    for (let i = 1; i < 明.length; i++) if (明[i] > 明[i - 1]) return false;
    return true;
  })(), T.見本の段(T.配色を引く('みどり（既定）'), 'accent1').join(' '));
}

console.log('\n[④ 画面から 押せる]');
for (const n of ['テーマを当てる', 'テーマの窓を開く', 'テーマを選ぶ', '配色の窓を開く',
  'この色で塗る', 'テーマのフォントの窓を開く', 'テーマのフォントを選ぶ',
  'テーマの効果の窓を開く', 'テーマの効果を選ぶ',
  '背景を選ぶ', '背景を読み込む', '背景を消す', '背景の窓を開く',
  'ふりがなを切り替え', 'ふりがなを入れる', 'ふりがなを決める']) {
  ok(n + ' が 在る', !!抜く(n));
}
ok('★塗る 出口は applyFormat 1つ★', /applyFormat\('bgColor', 色\);/.test(抜く('この色で塗る') || ''));
ok('★背景を 選ぶ 入り口が 在る★', /id="bgFileInput"/.test(book));
ok('★背景は 紙に 刷られない事を 画面に 書いてある★', /紙には 刷られません/.test(book));
/* ★実測の 6 は 点（pt）＝画面では 8px★（6px だと 読めなかった＝実ブラウザで 見た） */
ok('★ふりがなは 8px＝実測の 6点★', /'8px "Noto Sans JP",sans-serif'/.test(book));
ok('  点→画面の 直し方を 書き残している', /6×96\/72 = ★8px★/.test(book));
ok('★読み方の 在る 行だけ 伸ばす（実測 18.75→26.7＝1.424倍）★',
  /Math\.round\(ROW_H \* 1\.424\)/.test(book));
ok('★人が 高さを 決めた 行は 触らない★', /&& !sh\.rowH\[r\]\) \{/.test(book));
ok('★ふりがなは 人が 入れる事を 画面に 書いてある★', /ふりがなが 入りません/.test(book));
ok('★ふりがなは シートごとに 持つ★', /sheets\[activeSheet\]\.ruby/.test(book));
ok('★出していない 時は 描かない★', /if \(ふりがなを出す\) \{/.test(book));
ok('★字が 無い セルは 断る★', /字が 入っている セルを 選んでください/.test(book));
ok('★alert / prompt / confirm を 使っていない★', !/\balert\(|\bprompt\(|\bconfirm\(/.test(book));

console.log('\n[⑤ 副題を 決めていない 窓が 増えていないか]');
{
  const 行 = book.split(String.fromCharCode(10));
  const 抜け = [];
  for (let i = 0; i < 行.length; i++) {
    if (!/getElementById\('funcTitle'\)\.textContent =/.test(行[i])) continue;
    if (!/窓の副題\(/.test(行.slice(i, i + 5).join(String.fromCharCode(10)))) 抜け.push(i + 1);
  }
  ok('★副題を 決めていない 窓は 0個★', 抜け.length === 0, 抜け.join(' / '));
}

console.log('\n[⑥ リボンから 押せる]');
{
  const ACT = require_(path.join(ROOT, 'lib/ribbon-actions.js'));
  const g = globalThis, 前w = g.window;
  const 試す = (ボタン, 呼ぶ名) => {
    let 受け = 'よばれていない';
    const w = {}; w[呼ぶ名] = function () { 受け = 'ok'; };
    g.window = w; ACT[ボタン](); g.window = 前w;
    ok('「' + ボタン + '」→ ' + 呼ぶ名, 受け === 'ok', String(受け));
  };
  試す('テーマ', 'テーマの窓を開く');
  試す('配色', '配色の窓を開く');
  試す('テーマのフォント', 'テーマのフォントの窓を開く');
  試す('テーマの効果', 'テーマの効果の窓を開く');
  試す('背景', '背景の窓を開く');
  試す('ふりがな', 'ふりがなを入れる');
}

console.log('\ntheme: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝この 見張りが 見ている物を 直に 確かめる（★壊して 赤を 見るのは tools/break-check.mjs★）★');
  let 素通り = 0;
  if (T.濃淡('#156082', 0.4) !== '#44B3E1') { 素通り++; console.log('  ★素通り★ 濃淡が 実測と 違う'); }
  else console.log('  ok   濃淡が 実測と 同じ');
  /* 四捨五入に 変えた 物を 作り、★赤に なる事★を 見る */
  const にせ = (hex, t) => {
    const a = T.RGBをHLSへ(parseInt(hex.slice(1, 3), 16), parseInt(hex.slice(3, 5), 16), parseInt(hex.slice(5, 7), 16));
    const L = t > 0 ? Math.round(a[1] * (1 - t) + 240 * t) : Math.round(a[1] * (1 + t));
    const o = T.HLSをRGBへ(a[0], L, a[2]);
    return '#' + o.map(v => (v < 16 ? '0' : '') + v.toString(16).toUpperCase()).join('');
  };
  const 合 = 濃淡の実測.filter(([c, t, e]) => にせ(c, t) === e).length;
  if (合 === 濃淡の実測.length) { 素通り++; console.log('  ★素通り★ 四捨五入でも 全部 合ってしまう＝試験が 甘い'); }
  else console.log('  ok   四捨五入だと ' + 合 + '/' + 濃淡の実測.length + ' しか 合わない（★切り捨てが 正しい★）');
  const うちの色 = T.配色たち.flatMap(c => c.色.map(v => v.toUpperCase()));
  if (うちの色.indexOf('#156082') >= 0) { 素通り++; console.log('  ★素通り★ Excel の 色を 写している'); }
  else console.log('  ok   Excel の 色を 写していない');
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
process.exit(赤 ? 1 : 0);
