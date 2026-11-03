/* view-review-tab.test.mjs — ★表示（ズーム）と 校閲（保護・数）★ 2026-08-30
 *
 *  ★真値（実Excel 16.0 で 実測）★
 *    tools/measure-view-tab.ps1
 *      ・View の 既定 = 1（標準）／2=改ページプレビュー（★Zoom が 60 に なる★）／3=ページレイアウト
 *      ・Zoom は ★10〜400 が 入る★（★5 も 401 も 入らない★）／既定 100
 *      ・A1:E10 に 合わせたら Zoom=258（★窓の 大きさで 変わる数★＝真似ない）
 *      ・分割 … C5 で Split=True → SplitRow=4 / SplitColumn=2
 *    tools/measure-review-tab.ps1
 *      ・ブックの保護 … 既定 構造=False。★守ると シートが 足せない★・外すと 足せる
 *      ・メモ … 作者つき（作者＝Excelを 使っている人の名）・★既定は 隠れている★・形 93 × 54.8
 *      ・範囲の編集を許可する … 名前＋範囲を 持つ（既定 0個）
 *      ・統計 … シート数／中身の在るセル／式／字／数 を 数えられる
 *
 *  走らせ方: node tests/view-review-tab.test.mjs [--self-test]
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

console.log('\n[① 測った 道具が 残っている]');
ok('tools/measure-view-tab.ps1 が 在る', fs.existsSync(path.join(ROOT, 'tools/measure-view-tab.ps1')));
ok('tools/measure-review-tab.ps1 が 在る', fs.existsSync(path.join(ROOT, 'tools/measure-review-tab.ps1')));
/* ★人の 名前は ここに 書かない★（締め出す 字を 書くと 公開 repo に 名前を もう一度 出す・10-11）
   ⇒ 測った 中（book.html・tools/measure-review-tab.ps1 を 含む 全部の ファイル）に 名前が 戻らないかは、
     CI の 段「実在の字が戻っていないか（一覧は secret）」＝tests/kinshi-ji.test.mjs が 見る（一覧は repo の 外） */

console.log('\n[② ズーム（実測＝10〜400）]');
for (const n of ['ズームの窓を開く', 'ズームを決めて閉じる', '選択範囲に合わせる']) ok(n + ' が 在る', !!抜く(n));
ok('窓が 在る', /id="zoomOverlay"/.test(book));
ok('★下は 10★（前は 25 だった＝実測で 直した）', /var ズームの下 = 10/.test(book));
ok('★上は 400★', /ズームの上 = 400/.test(book));
ok('★決める所も 0.10 まで 下げた★', /Math\.max\(0\.10, Math\.min\(4, 倍\)\)/.test(book));
ok('★外れた数は 入れずに 理由を 出す★', /％から ' \+ ズームの上 \+ '％までです/.test(book));
{
  /* 範囲に 合わせる 計算（★数は 窓の 大きさ次第＝考え方だけ 見る★） */
  const f = new Function('selR1', 'selR2', 'selC1', 'selC2', 'sheets', 'activeSheet',
    'COL_W', 'ROW_H', 'wrapW', 'wrapH', 'HDR_W', 'HDR_H', 'ズームの下', 'ズームの上',
    'ズームを決める', 'ズームの窓を閉じる', 'notify',
    抜く('選択範囲に合わせる') + '\nreturn 選択範囲に合わせる;');
  let 決めた = null;
  const 走らす = (r2, c2, W, H) => {
    決めた = null;
    return f(0, r2, 0, c2, [{ colW: {}, rowH: {} }], 0, 80, 22, W, H, 46, 22, 10, 400,
      (v) => { 決めた = v; }, () => {}, () => {})();
  };
  const 大 = 走らす(9, 4, 1200, 700);      /* 5列×10行＝400x220 が 1146x656 に 入る */
  ok('★小さい所を 選ぶと 大きく なる★', 大 > 100, String(大));
  const 小 = 走らす(99, 49, 400, 300);     /* 50列×100行＝広い */
  ok('★広い所を 選ぶと 小さく なる★', 小 < 100, String(小));
  ok('★下限（10）より 下には しない★', 小 >= 10, String(小));
  const 上 = 走らす(0, 0, 4000, 4000);
  ok('★上限（400）より 上には しない★', 上 <= 400, String(上));
}

console.log('\n[③ ブックの保護（実測＝守ると シートが 足せない）]');
for (const n of ['ブックの保護を開く', 'ブックの保護を決める', 'ブックの守りに引っかかる']) ok(n + ' が 在る', !!抜く(n));
ok('窓が 在る', /id="bookProtOverlay"/.test(book));
ok('★はじめは 守っていない（実測 構造=False）★', /var ブックを守っている = false;/.test(book));
ok('★シートを 足す前に 門を 通る★',
  /function addSheet\(\)\{[\s\S]{0,200}ブックの守りに引っかかる\('シートを 足す事'\)/.test(book));
ok('★止めた時は 理由を 出す（黙って 素通りさせない）★', /ブックを 守っています＝/.test(book));
ok('★合い言葉を 出さない（守れているふりを しない）★', /合い言葉（パスワード）は 付けられません/.test(book));
{
  const f = new Function('notify', 'ブックを守っている',
    抜く('ブックの守りに引っかかる') + '\nreturn ブックの守りに引っかかる;');
  const 言った = [];
  ok('守っていなければ 通す', f((m) => 言った.push(m), false)('シートを 足す事') === false);
  ok('★守っていたら 止める★', f((m) => 言った.push(m), true)('シートを 足す事') === true);
  ok('  その時 理由を 言う', 言った.length === 1 && /できません/.test(言った[0]), JSON.stringify(言った));
}

console.log('\n[④ 範囲の編集を許可する（実測＝名前＋範囲・既定0個）]');
for (const n of ['編集を許す箱', '範囲の編集を開く', '範囲の編集を足す', '範囲の編集を消す',
  '範囲の編集の一覧', '編集を許された所か']) ok(n + ' が 在る', !!抜く(n));
ok('窓が 在る', /id="editRangeOverlay"/.test(book));
ok('★守っていても 許した所は 書ける★',
  /function ここに書けるか[\s\S]{0,400}編集を許された所か\(r, c\)\) return true;/.test(book));
{
  const f = new Function('sheets', 'activeSheet', 抜く('編集を許された所か') + '\nreturn 編集を許された所か;');
  const s0 = { allowEdit: [{ 名: 'は1', r1: 1, c1: 1, r2: 3, c2: 3 }] };
  const 中か = f([s0], 0);
  ok('中は true', 中か(2, 2) === true);
  ok('へりも true', 中か(1, 1) === true && 中か(3, 3) === true);
  ok('外は false', 中か(0, 0) === false && 中か(4, 4) === false);
  ok('1つも 無ければ false', f([{}], 0)(0, 0) === false);
}
ok('★同じ 名前は 足せない★', /その 名前は もう 在ります/.test(book));

console.log('\n[⑤ ブックの数（数えた物だけ 出す）]');
for (const n of ['ブックの数を数える', 'ブックの数を開く', 'ブックの数を閉じる']) ok(n + ' が 在る', !!抜く(n));
ok('窓が 在る', /id="statsOverlay"/.test(book));
{
  const f = new Function('sheets', 抜く('ブックの数を数える') + '\nreturn ブックの数を数える;');
  const 数 = f([
    { data: { '0,0': { v: 'あいう' }, '0,1': { v: 123 }, '1,0': { f: '=A1', d: '5' } },
      comments: { '0,0': {} }, links: {} },
    { data: { '0,0': { v: '' }, '0,1': { v: 'x' } } },
  ])();
  ok('シートは 2', 数.シート === 2, String(数.シート));
  ok('★中身の 在るセルだけ 数える（空は 数えない）★', 数.セル === 4, JSON.stringify(数));
  ok('式は 1', 数.式 === 1, String(数.式));
  ok('数は 1', 数.数 === 1, String(数.数));
  ok('字は 2', 数.字 === 2, String(数.字));
  ok('字の 数は 3+3+1+1=8', 数.文字数 === 8, String(数.文字数));
  ok('付箋は 1', 数.付箋 === 1, String(数.付箋));
}
ok('★数えた物だけ 出すと 書いてある★', /数えたのは 中身の 在るセルだけ/.test(book));
ok('★alert / prompt / confirm を 使っていない★', !/\balert\(|\bprompt\(|\bconfirm\(/.test(book));

console.log('\n[⑥ リボンから 押せる]');
{
  const ACT = require_(path.join(ROOT, 'lib/ribbon-actions.js'));
  const g = globalThis, 前w = g.window;
  for (const [ボタン, 呼ぶ名] of [
    ['ズーム', 'ズームの窓を開く'], ['選択範囲に合わせて拡大縮小', '選択範囲に合わせる'],
    ['ブックの保護', 'ブックの保護を開く'], ['範囲の編集を許可する', '範囲の編集を開く'],
    ['ブックの統計情報', 'ブックの数を開く'],
  ]) {
    let 受け = null;
    const w = {}; w[呼ぶ名] = function () { 受け = 'ok'; };
    g.window = w;
    ACT[ボタン]();
    g.window = 前w;
    ok('「' + ボタン + '」→ ' + 呼ぶ名, 受け === 'ok', String(受け));
  }
}

console.log('\n[⑦ ページ レイアウト（余白・用紙・印刷範囲・シートのオプション）]');
{
  for (const n of ['ページ設定の初期化', 'ページの詳しい設定を開く', 'ページの詳しい設定を決める',
    '余白を標準に', '印刷範囲を今の選択に', '印刷範囲を消す']) ok(n + ' が 在る', !!抜く(n));
  ok('窓が 在る', /id="pageMoreOverlay"/.test(book));
  /* ★実測の 値そのもの★ */
  ok('★余白の 標準＝上下 1.905cm・左右 1.778cm（54pt / 50.4pt）★',
    /var 余白の標準 = \{ 上: 1\.905, 下: 1\.905, 左: 1\.778, 右: 1\.778 \}/.test(book));
  ok('★1cm = 28.3464566929134pt（実測）★', /28\.3464566929134/.test(book));
  ok('★枠線・見出しの 表示は はじめ true（実測）★',
    /ページ設定\.枠線を表示 = true/.test(book) && /ページ設定\.見出しを表示 = true/.test(book));
  ok('★用紙の はじめは A4（実測）★', /ページ設定\.用紙 = 'A4'/.test(book));
  ok('★印刷範囲の はじめは 空（実測）★', /ページ設定\.印刷範囲 = ''/.test(book));
  ok('★測れなかった「広い／狭い」を 並べていない★',
    /「広い／狭い」の 数は 外から 読めませんでした/.test(book)
    && !/広い<\/option>/.test(book) && !/狭い<\/option>/.test(book));
  ok('★余白は 0〜10cm の 外を 断る★', /余白は 0〜10cm で 入れてください/.test(book));
  ok('★印刷範囲の 形が おかしければ 断る★', /印刷範囲は A1:C10 の 形で/.test(book));
  /* 刷る所に 本当に 渡しているか */
  ok('★刷る時に 余白と 用紙を 渡している★',
    /余白: ページ設定\.余白, 用紙: ページ設定\.用紙/.test(book));
  ok('★印刷範囲を 決めてあれば そこを 刷る★',
    /if\(!範囲 && typeof ページ設定 !== 'undefined' && ページ設定\.印刷範囲\)/.test(book));
  ok('★選んだ所が 在れば そちらが 勝つ★', /選んだ所が 在れば そちらが 勝つ/.test(book));
  /* 部品の 側 */
  const 刷 = fs.readFileSync(path.join(ROOT, 'lib/grid-print.js'), 'utf8');
  ok('部品が 余白を 受け取る', /var 余 = o\.余白 \|\| \{ 上: 1\.905/.test(刷));
  ok('部品が 用紙を 受け取る（A4/A3/A5/B5/Letter）',
    /A4: 'A4', A3: 'A3', A5: 'A5', B5: 'B5', Letter: 'letter'/.test(刷));
  ok('★4方向べつべつに 書く★', /余\.上 \+ 'cm ' \+ 余\.右 \+ 'cm ' \+ 余\.下 \+ 'cm ' \+ 余\.左/.test(刷));
  /* 決めた ✓ が その場で 画面に 効くか */
  ok('★決めたら その場で 画面に 効かせる★',
    /枠線を出す = ページ設定\.枠線を表示/.test(book) && /window\.見出しを出す = ページ設定\.見出しを表示/.test(book));

  const ACT = require_(path.join(ROOT, 'lib/ribbon-actions.js'));
  const g = globalThis, 前w = g.window;
  for (const [ボタン, 期待] of [
    ['余白', '余白'], ['サイズ', '用紙'], ['印刷範囲', '印刷範囲'],
    ['シートのオプション表示', '表示'], ['シートのオプション印刷', '印刷'],
  ]) {
    let 受け = null;
    g.window = { ページの詳しい設定を開く: function (a) { 受け = a; } };
    ACT[ボタン]();
    g.window = 前w;
    ok('「' + ボタン + '」→ ページの詳しい設定を開く（' + 期待 + '）', 受け === 期待, String(受け));
  }
}

console.log('\nview-review-tab: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝この 見張りが 見ている物を 直に 確かめる（★壊して 赤を 見るのは tools/break-check.mjs★）★');
  let 素通り = 0;
  /* 壊し① ズームの 下が 25 に 戻っていないか（実測は 10） */
  if (/Math\.max\(0\.25, Math\.min\(4, 倍\)\)/.test(book)) {
    素通り++; console.log('  ★素通り★ ズームの 下が 25 に 戻っている（実測は 10）');
  } else console.log('  ok   ズームの 下は 10');
  /* 壊し② 守りの 門が 外れていないか */
  if (!/ブックの守りに引っかかる\('シートを 足す事'\)/.test(book)) {
    素通り++; console.log('  ★素通り★ シートを 足す所の 門が 無い');
  } else console.log('  ok   シートを 足す所に 門が 在る');
  /* 壊し③ 許可範囲を 見ずに 保護だけで 決めていないか */
  if (!/編集を許された所か\(r, c\)\) return true;/.test(book)) {
    素通り++; console.log('  ★素通り★ 許した範囲を 見ていない');
  } else console.log('  ok   許した範囲を 見ている');
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
process.exit(赤 ? 1 : 0);
