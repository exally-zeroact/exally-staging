/* ribbon.test.mjs — ★リボン（Excelと同じ配置・見せ方は うちの物）★ 2026-08-29
 *
 *  ★司さんの指示★
 *    「リボンは前から言うてるけど ★配置なども真似しろ★」
 *    「★訴えられんような見せ方で 同じように★」
 *    「★Excel全機能全能力が Exallyに 入って Excelの最上級に なる★」
 *
 *  ★ここで 見る事★
 *    ①並びが ★正本（実Excelから 機械で取った物）と 1文字も 違わない★
 *    ②★出来ていない物の ボタンを 出していない★（押せない物を 見せない）
 *    ③★結び先の関数が 本当に 在る★（押しても 何も起きない ボタンを 作らない）
 *    ④★絵を 写していない★＝印は 字だけ（Excelの絵は 1つも 使わない）
 *    ⑤★対象外は 理由が 書いてある物だけ★（黙って 分母から 引かない）
 *    ⑥★実際に 描いて 数える★（作っただけで 緑に しない）
 *
 *  走らせ方: node tests/ribbon.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';
import { JSDOM } from 'jsdom';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const 壊す = process.argv.includes('--self-test');
let 緑 = 0, 赤 = 0;
const ok = (名, 条件, 添え) => {
  if (条件) { 緑++; console.log('  ok   ' + 名); }
  else { 赤++; console.log('  ★NG★ ' + 名 + (添え !== undefined ? '  … ' + 添え : '')); }
};

const SPEC = require_(path.join(ROOT, 'lib/ribbon-spec.js'));
const RB = require_(path.join(ROOT, 'lib/ribbon.js'));
const book = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');

/* ── 正本 ─────────────────────────────────── */
const NL = String.fromCharCode(10), TAB = String.fromCharCode(9), CR = String.fromCharCode(13);
const 正本 = Array.from(new Set(
  fs.readFileSync(path.join(ROOT, 'docs/excel-ribbon-flat.tsv'), 'utf8')
    .split(NL).map((l) => l.replace(CR, '')).filter((l) => l.trim())
)).map((l) => l.split(TAB)).filter((c) => c.length >= 3);

console.log('\n[① 並びが 正本と 1文字も 違わない]');
ok('★正本が 空でない（試験が 空振りしていない）', 正本.length >= 100, 正本.length + '個');
ok('部品の数が 同じ', SPEC.ITEMS.length === 正本.length, SPEC.ITEMS.length + ' / ' + 正本.length);
let ずれ = -1;
for (let i = 0; i < Math.min(SPEC.ITEMS.length, 正本.length); i++) {
  const a = SPEC.ITEMS[i], b = 正本[i];
  if (a.t !== b[0] || a.g !== b[1] || a.p !== b[2]) { ずれ = i; break; }
}
ok('★並び順まで 同じ★', ずれ < 0, ずれ < 0 ? '' : (ずれ + 1) + '番目が 違う');
const tabs = SPEC.ツリー();
/* ★Sheet1 を 正本から 落とした★（2026-08-30 測り直し＝下の シート見出しの 写り込み） */
ok('タブ 11個（Sheet1は 幻）', tabs.length === 11, String(tabs.length));
ok('グループ 65個', tabs.reduce((a, t) => a + t.groups.length, 0) === 65,
  String(tabs.reduce((a, t) => a + t.groups.length, 0)));

/* ── ③ ★結び先を 実際に 押す★ ───────────────────────
 *  ★「関数が 在る」だけ 見ていて 危なかった★（2026-08-29）:
 *    applyFormat は 引数が 2つ・openColorPalette は event を 取る のに
 *    1つの文字で 呼ぶ結びを 書いていた ⇒ ★押しても 何も起きないボタン★に なる所だった。
 *  ⇒ ★動作の層を 1つずつ 実際に 呼び、画面の関数に 正しく 届いたかを 数える★ */
console.log('\n[③ ★結び先を 実際に 押す★]');
const ACT = require_(path.join(ROOT, 'lib/ribbon-actions.js'));
const 結んだ = SPEC.ITEMS.filter((x) => x.a);
ok('★1つ以上 結んである（試験が 空振りしていない）', 結んだ.length >= 20, 結んだ.length + '個');
const 名無し = 結んだ.filter((x) => typeof ACT[x.a.act] !== 'function');
/* ★引き取る物も 名前を 持つ★＝上の 検査に かかる（名前だけは 要る） */
ok('★働きが 無い結びが 0件★', 名無し.length === 0, 名無し.map((x) => x.p + '→' + x.a.act).join(', '));

/* 画面の関数を にせ物に 差し替えて、★どれが 何を どう 呼んだか★ を 記録する */
{
  const 記録 = [];
  const g = globalThis;
  const 前window = g.window, 前document = g.document;
  const にせ画面 = {};
  const 拾う名 = ['ctxPaste','ctxCut','ctxCopy','toggleFormat','applyBorderAll','applyFormat',
    'openColorPalette','openFmtModal','applyMergeCells','applyIndent','toggleDecimal',
    'openCondFormat','ctxInsertRow','ctxDeleteRow','ctxInsertCol','ctxDeleteCol','findNextCell',
    'doUndo','doRedo','openFnPalette','sortRange','filterByValue','clearFilter','openValid',
    'freezePanes','unfreezePanes','printSheet',
    'autoSum','fillFromEdge','中身を消す','書式を消す','すべて消す',
    '数式の表示を切り替える','すべて再計算','このシートを再計算','openShindan','重複を削除','ズーム100','openMacro','名前の窓を開く','区切り位置を開く','グラフを作る',
    'コメントの窓を開く','コメントを消す','前のコメントへ','次のコメントへ','コメントの一覧',
    'ページ設定を開く','向きを切り替える','枠線を刷るか','見出しを刷るか',
    'シートの保護を切り替える','選んだ所のロックを切り替える','枠線の表示を切り替える',
    '小数を増やす','小数を減らす','書式をコピー','書式を貼る','セルを下へ挿入','セルを上へ削除','行の高さを開く','列の幅を開く','セルのスタイルを開く','並べ替えの窓を開く','見出しを出すか','数式バーを出すか','ctxUnhideRow','ctxUnhideCol','openRireki','リンクの窓を開く','おすすめグラフ','スパークラインの窓を開く','表の窓を開く','記号の窓を開く','フォームの窓を開く','計算方法の窓を開く','関数の窓を開く','数式で使用の窓を開く','範囲から名前の窓を開く','参照元のトレース','参照先のトレース','トレース矢印の削除','数式の検証を開く','ウォッチを開く','ズームの窓を開く','選択範囲に合わせる','ブックの保護を開く','範囲の編集を開く','ブックの数を開く','フラッシュフィル','絞りを再適用','ゴールシークを開く','ページの詳しい設定を開く','フォントの窓を開く','文字の向きの窓を開く','クリップボードを開く','ピボットの窓を開く','ピボットグラフを作る','統合の窓を開く','アウトラインの窓を開く','スライサーを開く','予測の窓を開く','重さを調べる','サブテーブルを作る','リンクを全部消す','空の行を数える','テーブルをJSONにする','表からピボット','データの取得を開く','CSVの窓を開く','openRireki','画像の窓を開く','図形の窓を開く','テキストを置く','前面へ移動','背面へ移動','物の一覧を開く','描画を始める','選ぶにする','なげなわにする','消しゴムにする','ペンにする','ペンを足す','インクを再生','インクのヘルプ','拡大縮小を開く','印刷倍率を足す','ページ設定を開く','配置の窓を開く','グループの窓を開く','回転の窓を開く',
  'シートビューの窓を開く','シートビューを保持','シートビューを終了','シートビューを新規',
  'シートビューのオプション','標準の表示にする','改ページプレビューにする','ページレイアウト表示にする',
  'ブックのビューの窓を開く','新しいウィンドウを開く','整列の窓を開く','分割する','窓を表示しない',
  '窓を再表示','並べて比較','同時にスクロールを切り替え','窓の位置を元に戻す','窓の切り替えを開く',
  '改ページの窓を開く','テーマの窓を開く','配色の窓を開く','テーマのフォントの窓を開く',
  'テーマの効果の窓を開く','背景の窓を開く','ふりがなを入れる',
  '図解の窓を開く','アイコンの窓を開く','数式の窓を開く','チェックボックスを入れる','画面を撮って貼る',
  '詳細設定を開く','表から接続を作る','すべて更新','接続の窓を開く','接続のプロパティ','ブックのリンクを見る',
  'スペルを見る','類義語を聞く','見やすさを見る','翻訳を聞く','変更内容を見る','共有を解除','インクを隠す切り替え',
  'XMLの窓を開く','XMLの決めを見る','XML拡張パック','XMLを更新','XMLを書き出す',
  'コントロールを挿入','コントロールのプロパティ',
  'インクを図形に','インクを数式に','アクションペン','マップを作る','タイムラインを作る',
  '三次元の窓を開く','Webから読む','データモデルを開く','データ分析',
  /* ★表示｜表示★（2026-08-30 実Excelを 測り直して 見つけた 組） */
  'ルーラーを出すか','ナビゲーションを開く','セルにフォーカスを切り替える',
  /* ★[開発]→[Visual Basic]＝マクロの 字を そのまま 見る（読むだけ・動かさない）★ 2026-09-03 */
  'VBAの字を見る'];
  for (const n of 拾う名) にせ画面[n] = (function (nm) {
    return function () { 記録.push({ fn: nm, args: Array.prototype.slice.call(arguments) }); };
  }(n));
  const dom2 = new JSDOM('<input id="fmt-font-size" value="12">');
  g.window = Object.assign(にせ画面, { RibbonActions: ACT });
  g.document = dom2.window.document;
  let 届かない = [];
  for (const it of 結んだ) {
    記録.length = 0;
    try { ACT[it.a.act](); } catch (e) { 届かない.push(it.a.act + '（落ちた: ' + e.message.slice(0, 40) + '）'); continue; }
    /* ★引き取る物（元から在る 入力・色の見本）は 押す物では ない★＝除く。
       押す働きは 元の物が 持っている（配線ごと リボンへ 引っ越すだけ）。 */
    if (!記録.length && it.a.act !== '字を大きく' && it.a.act !== '字を小さく' && !it.a.取り込む) {
      届かない.push(it.a.act + '（画面の関数を 1つも 呼んでいない）');
    }
  }
  g.window = 前window; g.document = 前document;
  ok('★押したら 画面の関数に 届く（' + 結んだ.length + '個 全部 押した）★',
    届かない.length === 0, 届かない.join(' / '));
}

/* ★関数ライブラリの 分類★＝引数が 効いているか（2026-08-29 に 足した）
 *   知らない分類（財務など・うちに まだ 無い物）は ★すべて★に なる＝黙って 落ちない。 */
{
  const g = globalThis, 前w = g.window;
  const 受け = [];
  g.window = { openFnPalette: function (c) { 受け.push(c); } };
  ACT['論理の関数'](); ACT['日付の関数'](); ACT['関数の挿入']();
  g.window = 前w;
  ok('★分類を 渡している★', JSON.stringify(受け) === JSON.stringify(['論理', '日付', undefined]), JSON.stringify(受け));
}

/* ★引数の 取り違えを 名指しで 弾く★（前に やらかした形） */
{
  const 記録 = [];
  const g = globalThis;
  const 前window = g.window, 前document = g.document;
  g.window = {
    applyFormat: function (k, v) { 記録.push(['applyFormat', k, v]); },
    toggleFormat: function (k) { 記録.push(['toggleFormat', k]); },
    openColorPalette: function (e, t) { 記録.push(['openColorPalette', e, t]); },
    applyIndent: function (d) { 記録.push(['applyIndent', d]); },
    sortRange: function (d) { 記録.push(['sortRange', d]); },
  };
  g.document = new JSDOM('<div></div>').window.document;
  ACT['左揃え'](); ACT['太字'](); ACT['字下げを増やす'](); ACT['昇順']();
  g.window = 前window; g.document = 前document;
  const 見つけ = (fn) => 記録.find((r) => r[0] === fn);
  ok('applyFormat は ★2つ★の引数で 呼ぶ',
    JSON.stringify(見つけ('applyFormat')) === JSON.stringify(['applyFormat', 'align', 'left']),
    JSON.stringify(見つけ('applyFormat')));
  ok('toggleFormat は key 1つ', JSON.stringify(見つけ('toggleFormat')) === JSON.stringify(['toggleFormat', 'bold']));
  ok('applyIndent は ★数★で 呼ぶ', JSON.stringify(見つけ('applyIndent')) === JSON.stringify(['applyIndent', 1]));
  ok('sortRange は 向きの文字', JSON.stringify(見つけ('sortRange')) === JSON.stringify(['sortRange', 'asc']));
}

/* ── ④ 絵を 写していない ───────────────────── */
console.log('\n[④ 絵を 写していない（印は 字だけ）]');
const 絵っぽい = 結んだ.filter((x) => x.a.icon && /<svg|<img|url\(|\.png|\.svg|\.ico/i.test(x.a.icon));
ok('★印に 画像を 使っていない★', 絵っぽい.length === 0, 絵っぽい.map((x) => x.p).join(', '));
const css = fs.readFileSync(path.join(ROOT, 'lib/ribbon.css'), 'utf8');
ok('★見た目に 外の画像を 使っていない★', !/url\(\s*['"]?(https?:|\/\/)/.test(css));
ok('★うちの緑を 使っている★', /#3D9E72|#2E7D54|#C8ECD8|#F0FAF4/i.test(css));

/* ── ⑤ 対象外は 理由つき ───────────────────── */
console.log('\n[⑤ 対象外は 理由が 書いてある物だけ]');
/* ★対象外の 正本は lib/ribbon-scope.js に 移した★（2026-08-30）
   前は scripts/ribbon-bind.mjs の 中だけに 在り、★画面からは 見えなかった★。
   そのせいで ★作らないと 決めた物にも「これから」と 出ていた★（監査役が 絵で 見つけた）。 */
const bind = fs.readFileSync(path.join(ROOT, 'scripts/ribbon-bind.mjs'), 'utf8');
const scope = fs.readFileSync(path.join(ROOT, 'lib/ribbon-scope.js'), 'utf8');
const 対象外の数 = (scope.match(/'[^']+\|[^']+\|[^']+':/g) || []).length;
ok('★対象外に 理由が 書いてある★', 対象外の数 >= 10, String(対象外の数));
/* ★\b は 日本語の 後ろでは 効かない★（単語の切れ目の 決まりが [A-Za-z0-9_] 前提）。
   実測 2026-08-29＝/export const 対象外\b/ は 在るのに 当たらなかった。 */
ok('★対象外の一覧が repo に 在る（口約束に しない）', /var 対象外\s*=/.test(scope));
ok('★対象外のタブの一覧も 在る', /var 対象外タブ\s*=/.test(scope));
ok('★「出さない」の一覧も 在る（押す物では ない 組）', /var 出さない\s*=/.test(scope));
ok('★bind は その1本を 見ている（二重に 持たない）★',
  /ribbon-scope\.js/.test(bind));

/* ── ⑥⑦ 実際に 描く ───────────────────────── */
console.log('\n[⑥ 実際に 描いて 数える]');
const dom = new JSDOM('<div id="ribbon"></div>');
const el = dom.window.document.getElementById('ribbon');
const r = RB.描く(el, SPEC);
ok('★描けた★', !!r && r.tab === 'ホーム', JSON.stringify(r));
const 描いたタブ = el.querySelectorAll('.rb-tab').length;
/* ★Sheet1 は 描かない★（2026-08-30 実Excelを 測り直して 確かめた）
   Sheet1 を 選んだ時に UIAが 返したのは
   「シートの追加／標準／改ページ プレビュー／右スクロール」＝
   ★下の シート見出しと 状態バー★であって リボンの タブでは ない。
   ★空の タブを 見せない★ので 11個が 正しい。
   理由は lib/ribbon-scope.js の 出さない[] に 書いてある。 */
ok('タブを 11個 描いた（Sheet1は 幻なので 出さない）', 描いたタブ === 11, String(描いたタブ));
const 描いた群 = el.querySelectorAll('.rb-group').length;
ok('★ホームの グループを 9個 描いた★', 描いた群 === 9, String(描いた群));
const ボタン = [...el.querySelectorAll('.rb-item')];
ok('★押せるボタンを 1つ以上 描いた★', ボタン.length >= 10, ボタン.length + '個');

console.log('\n[② 出来ていない物の ボタンを 出していない]');
/* ★その組の ↘ そのものは 組の 中に 出さない★ 2026-09-03
   ＝実Excel では ★組の 下端の Button（上=321）＝右下に ↘ として 1つ 描く物★。
   中にも 出すと ★同じ字が 2つ 並ぶ★（ページ レイアウトで「ページ設定」が 3つに 見えた）。
   ★線は 緩めていません★＝右下に 出ている事は ★下の ⑧で 別に 数える★。
   正本：lib/ribbon-launcher.js の 実の起動ツール（実測 8個） */
const LAUNCH = require_(path.join(ROOT, 'lib/ribbon-launcher.js'));
const 結び済み名 = new Set(SPEC.ITEMS
  .filter((x) => x.t === 'ホーム' && x.a && !x.a.取り込む)
  .filter((x) => !LAUNCH.起動の品か(x.t, x.g, x.p))
  .map((x) => RB._名を短く(x.p)));
const 出た名 = ボタン.map((b) => b.getAttribute('title'));
const 余計 = 出た名.filter((n) => !結び済み名.has(n));
ok('★結んでいない物が 画面に 出ていない★', 余計.length === 0, 余計.join(', '));
ok('★ボタンの数＝結んだ数★', ボタン.length === 結び済み名.size, ボタン.length + ' / ' + 結び済み名.size);
const 空箱 = [...el.querySelectorAll('.rb-group[data-empty="1"]')];
ok('★中身が 無い箱は「これから」と 出す（偽のボタンを 出さない）',
  空箱.every((g) => g.querySelectorAll('.rb-item').length === 0 && /これから/.test(g.textContent)),
  空箱.length + '箱');

/* ═══ ★★中身が 空の 組を 赤にする★★ 2026-09-03 ═══
 *  ★なぜ 要るか★＝2026-09-03 に ★名前だけの 空っぽの 箱★を 出していた
 *    （ページ レイアウト｜シートのオプション）。★数字には 1つも 出なかった★
 *    ＝札の切れ 0・↘ 8/8・▾ 15/15 とも 緑のまま。★絵を 開いて 初めて 分かった★。
 *  ★★数え方を 1つ 強くした（理由つき）★★
 *    指示役の 注文は「★↘ だけの 組も 中身1 と 数える★」でした。
 *    ★その形だと この見張りは 一度も 赤に なれません★＝実際に 壊して 試したら 緑のまま でした
 *    （↘ が 中身に 数えられるので、今回の 事故そのものが 通ってしまう）。
 *    ⇒ ★押せる札（rb-item）が 0個の 組を 赤にする★＝★今回の 事故を 捕まえられる形★に した。
 *    ★↘ は 中身に 数えない★（↘ だけの 組は ★空っぽに 見える★＝それが 今回 起きた事）。
 *  ★「これから」「付けません」と 書いてある 箱は 別★＝理由を 出しているので 空では ない。
 *
 *  ★★正直に 書いておく（この見張りの 限界）★★
 *    ここは JSDOM です。★JSDOM では ↘ の 表（RibbonLauncher）が 画面に 積まれない★ので、
 *    ↘ だけの 組は 「これから」に なって ★この見張りを すり抜けます★。
 *    ⇒ 2026-09-03 に ★本物の 抜け（結び目が 4つ 消えた）を 通してしまった★。
 *    ⇒ ★本当に 効いたのは 下の「結び目が 減っていない」★と ★実ブラウザで 数える事★。
 *    ⇒ ★この 見張りだけを 頼りに しない事★。 */
{
  const 空っぽ = [];
  let 見た組 = 0;
  /* ★↘ が 描かれる 状態で 測る★＝働きが 無いと ↘ が 出ず、
     組が「これから」に なって ★見張りが 空振りする★（2026-09-03 実際に 空振りした）。 */
  const 前の働き = globalThis.window && globalThis.window.RibbonActions;
  if (globalThis.window) globalThis.window.RibbonActions = ACT;
  /* ★↘ の 表も 積む★＝積まないと ↘ が 出ず、組が「これから」に なって ★この見張りが 空振りする★
     （2026-09-03 実際に 空振りした＝本物の 抜け（結び目が 4つ 消えた）を 通してしまった） */
  const 前の起動 = globalThis.window && globalThis.window.RibbonLauncher;
  if (globalThis.window) globalThis.window.RibbonLauncher = LAUNCH;
  for (const t of SPEC.ツリー()) {
    RB.状態.tab = t.name;
    RB.描く(el, SPEC);
    for (const g of el.querySelectorAll('.rb-group')) {
      見た組++;
      const 数 = g.querySelectorAll('.rb-item, [data-take]').length;   /* ★↘ は 数えない★ */
      const 訳あり = /これから|付けません/.test(g.textContent);
      if (数 === 0 && !訳あり) 空っぽ.push(t.name + '|' + g.dataset.group);
    }
  }
  if (globalThis.window) globalThis.window.RibbonActions = 前の働き;
  if (globalThis.window) globalThis.window.RibbonLauncher = 前の起動;
  RB.状態.tab = 'ホーム'; RB.描く(el, SPEC);
  ok('★名前だけの 空っぽな 組が 0個★（' + 見た組 + '組 を 全部 見た）',
    空っぽ.length === 0, 空っぽ.join(' / '));
  /* ★この 見張り自身が 空振りしていないか★ 2026-09-03
     ＝★名前だけの 箱★を わざと 1つ 作って、上の 数え方が それを 拾えるか 見る。
     （★画面側を 壊す形では 赤に できなかった＝「これから」に なって 逃げる★ので、
       ★数え方そのもの★を 試す。★これが 出来ないと 見張りは 嘘をつく★） */
  {
    const 仮 = new JSDOM('<div><div class="rb-group" data-group="わざと空"><div class="rb-items"></div>'
      + '<div class="rb-gname">わざと空</div></div></div>').window.document;
    const g2 = 仮.querySelector('.rb-group');
    const 数2 = g2.querySelectorAll('.rb-item, [data-take]').length;
    const 訳2 = /これから|付けません/.test(g2.textContent);
    ok('★この 見張りは 名前だけの 箱を 拾える（空振りしていない）★', 数2 === 0 && !訳2,
      '数=' + 数2 + ' 訳=' + 訳2);
  }
  /* ═══ ★★結び目が 黙って 減っていないか★★ 2026-09-03 ═══
   *  ★2026-09-03 に 実際に やらかした★＝正本を 作り直す 途中で
   *  ★4つの 結び目（シートのオプションの 枠線/見出し）が `a: null` に 戻った★のに
   *  ★どの 見張りも 赤に ならなかった★（「まだ」の 数が 57→61 に 増えただけ）。
   *  ⇒ ★結んだ数の 下限を 決めて、下回ったら 赤★。
   *  ★わざと 減らす時は この数字を 一緒に 直す★＝★黙って 減らせない★。 */
  {
    const 下限 = 253;
    const 今 = SPEC.ITEMS.filter((x) => x.a).length;
    ok('★結び目が 減っていない（' + 下限 + '個 以上）★', 今 >= 下限,
      '今 ' + 今 + '個（下限 ' + 下限 + '）★減らすなら この数字も 一緒に 直す★');
  }
  /* ═══ ★★同じ組の 中に「同じ働き」の 札が 2つ 出ていないか★★ 2026-09-03 ═══
   *  ★今日 3回とも「同じ字が 並ぶ」で 差し戻された★
   *  （表示×4／印刷×3／ページ設定×3 → 名前の定義×2）。★人の目で 探すのを やめて 機械に 移す★。
   *  ★字だけで 見ると 足りない★＝「オート SUM」と「合計」は ★字は 違うが 働きは 同じ★だった。
   *  ⇒ ★働き（act）で 数える★。★同じ組に 同じ働きが 2つ 在れば 赤★。
   *  ★働きが 違うなら 2つ 在ってよい★＝実Excel の 分かれボタンの 既定動作を
   *    ★1押しで 出せる★形（罫線／下罫線 など）。★消すと 1押しが 2段に 落ちる★。 */
  {
    const 束 = {};
    for (const x of SPEC.ITEMS) {
      if (!x.a || !x.a.act) continue;
      /* ★その組の ↘ そのものは 組の 中に 描かない★＝画面に 2つ 並ばない。
         （↘ が 同じ 窓を 開くのは 当たり前＝ホーム｜数値の「セルの書式設定」） */
      if (LAUNCH.起動の品か(x.t, x.g, x.p)) continue;
      const k = x.t + '|' + x.g + '|' + x.a.act;
      (束[k] = 束[k] || []).push(x.p);
    }
    const 二重 = Object.keys(束).filter((k) => 束[k].length > 1)
      .map((k) => k + '（' + 束[k].join(' / ') + '）');
    ok('★同じ組に 同じ働きの 札が 2つ 無い★（' + Object.keys(束).length + '通り を 見た）',
      二重.length === 0, 二重.slice(0, 6).join('\n       '));
  }
}

console.log('\n[⑦ 画面に 差し込んである]');
ok('book.html に リボンの箱が 在る', /id="ribbon"/.test(book));
ok('book.html が 部品を 読み込んでいる', /lib\/ribbon-spec\.js/.test(book) && /lib\/ribbon\.js/.test(book));
ok('book.html が 見た目を 読み込んでいる', /lib\/ribbon\.css/.test(book));
ok('★起動時に 描いている★', /Ribbon\.描く/.test(book));

/* ── 数を 出す ─────────────────────────────── */
const n = SPEC.数える();
/* ═══ ★画面の 中で 「触ってから でないと 動かない」物を 探す★ 2026-08-30 ═══
 *  ★実ブラウザで 見つけた 事故★
 *    `onfocus` で 入れた 変数を `onblur` が そのまま 使っていた。
 *    リボンの「字を大きく」は ★触らずに blur を 呼ぶ★ので
 *    `window._fsCell[0]` が ★undefined[0]★ に なり ★落ちた★。
 *  ⇒ ★onblur / onchange の 中で 使う window の 覚え書きは
 *     「入っていなければ 今の 選び」に 落とす★ こと。
 */
{
  const 行 = book.split(String.fromCharCode(10));
  const あぶない = [];
  for (let i = 0; i < 行.length; i++) {
    const l = 行[i];
    if (!/onblur=|onchange=/.test(l)) continue;
    /* window.〈名〉[0] の ように 添え字を そのまま 取っている物 */
    const m = l.match(/window\.(_[A-Za-z0-9_]+)\[/g);
    if (!m) continue;
    for (const g of m) {
      const 名 = g.replace('window.', '').replace('[', '');
      /* 同じ 行で 「無ければ こうする」を 書いているか
         ★正規表現に すると 逃がしの `\\.` が 素の 字に なり、
         `|` だけの 空の 選択肢が 出来て ★何でも 通して しまう★。
         （08-30 実測＝これで 見張りが 1件も 拾えていなかった）
         ⇒ ★字を そのまま 探す★。 */
      const 逃 = l.replace(/\s+/g, '');
      if (逃.indexOf('window.' + 名 + '||') >= 0) continue;
      あぶない.push((i + 1) + ': ' + 名 + ' … ' + l.trim().slice(0, 70));
    }
  }
  ok('★触ってから でないと 動かない 覚え書きが 0個★', あぶない.length === 0, あぶない.join(' / '));
}

console.log('\n  ── 今の数 ──');
console.log('  Excelの部品 … ' + n.全 + '個 ／ ★うちに 在る … ' + n.有 + '個★ ／ まだ … ' + n.無 + '個');

console.log('\nribbon: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝わざと 壊して 赤に なるか★');
  let 素通り = 0;
  /* 壊し① 並びを 1つ 入れ替える */
  const 写し = SPEC.ITEMS.slice();
  const 入替 = 写し.slice(); const t0 = 入替[0]; 入替[0] = 入替[1]; 入替[1] = t0;
  let ず = -1;
  for (let i = 0; i < 入替.length; i++) {
    if (入替[i].t !== 正本[i][0] || 入替[i].g !== 正本[i][1] || 入替[i].p !== 正本[i][2]) { ず = i; break; }
  }
  if (ず < 0) { 素通り++; console.log('  ★素通り★ 並びを 入れ替えても 気づかない'); }
  /* 壊し② ★動作の層に 無い名前を 結ぶ★（前は「関数が在るか」だけ 見ていて 危なかった） */
  if (typeof ACT['この働きは絶対に無い___'] === 'function') { 素通り++; console.log('  ★素通り★ 無い働きを 在ると 言った'); }
  /* 壊し③ ★引数の 取り違え★＝1つで 呼ぶと 画面に 届かない事を 見せる */
  {
    const g = globalThis, 前w = g.window, 前d = g.document;
    let 受けた = null;
    g.window = { applyFormat: function (k, v) { 受けた = [k, v]; } };
    g.document = new JSDOM('<div></div>').window.document;
    /* わざと 1つで 呼ぶ（＝前に 私が 書いた 形） */
    g.window.applyFormat('align:left');
    g.window = 前w; g.document = 前d;
    if (受けた && 受けた[1] !== undefined) { 素通り++; console.log('  ★素通り★ 1つで 呼んでも 値が 届いた事に なっている'); }
  }
  /* 壊し③ 結んでいない物を 画面に 出す */
  const にせ = { t: 'ホーム', g: 'テスト', p: 'にせもの', a: null };
  if (RB._部品のHTML(にせ) !== '') { 素通り++; console.log('  ★素通り★ 結んでいない物の ボタンを 作った'); }
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
process.exit(赤 ? 1 : 0);
