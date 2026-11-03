/* range-gap.mjs — ★Range の 命令 110個を 1つずつ うちと 突き合わせる★ 2026-08-30
 *
 *  ★司さんの 方針「細胞分解レベルまで 網羅」★
 *  Range は ★セルそのもの★＝Excel の 一番 太い 幹。
 *
 *  ★数え方★
 *    ・COM の 呼び出し用（AddRef/Release/QueryInterface/GetIDsOfNames/
 *      GetTypeInfo/GetTypeInfoCount/Invoke）と ★_ や __ で 始まる 影★は 除く
 *    ・うちに 在るか は ★手で 決めた 対応表★（名前の 一致では 測れないので）
 *      ⇒ ★推測では なく「その働きが 画面に 在るか」で 決める★
 *
 *  走らせ方: node tools/range-gap.mjs
 */
import fs from 'node:fs';
import path from 'node:path';

const ROOT = path.join(path.dirname(new URL(import.meta.url).pathname.replace(/^\/([A-Za-z]:)/, '$1')), '..');
const NL = String.fromCharCode(10), TAB = String.fromCharCode(9);
const 行 = fs.readFileSync(path.join(ROOT, 'docs/excel-objectmodel-2026-08-30.tsv'), 'utf8')
  .split(NL).filter((l) => l.trim() && !l.startsWith('#')).map((l) => l.split(TAB));

const COM = ['AddRef', 'Release', 'QueryInterface', 'GetIDsOfNames', 'GetTypeInfo',
  'GetTypeInfoCount', 'Invoke'];
const 命令 = [...new Set(行.filter((c) => c[1] === 'Range' && c[2] === 'Method').map((c) => c[3]))]
  .filter((n) => !COM.includes(n) && !n.startsWith('_')).sort();

/* ★うちに 在る★＝画面に その働きが 在ると ★実際に 押して 確かめた★物だけ */
const 在る = {
  Activate: 'セルを選ぶ', Calculate: '再計算', Clear: 'すべて消す',
  ClearContents: '中身を消す', ClearFormats: '書式を消す', Copy: 'コピー', Cut: '切り取り',
  Delete: '削除', Find: '検索', FindNext: '次を検索', Insert: '挿入',
  Merge: 'セルを結合', UnMerge: '結合を解除', PasteSpecial: '値のみ貼り付け',
  Replace: '置換', Select: '範囲を選ぶ', Sort: '並べ替え', AutoFilter: 'フィルター',
  AutoFit: '列幅の自動調整', RemoveDuplicates: '重複を削除', TextToColumns: '区切り位置',
  GoalSeek: 'ゴールシーク', PrintOut: '印刷', PrintPreview: '印刷プレビュー',
  ShowPrecedents: '参照元のトレース', ShowDependents: '参照先のトレース',
  ClearHyperlinks: 'ハイパーリンクを削除', AddComment: 'コメントの挿入',
  ClearComments: 'コメントを消す', FillDown: '下へコピー', FillRight: '右へコピー',
  BorderAround: '罫線', Group: 'グループ化', Ungroup: 'グループ解除',
  SpecialCells: '条件を選んで選択', FlashFill: 'フラッシュフィル',
  Subtotal: '小計', AdvancedFilter: '詳細設定（フィルター）',
};

const 無い = 命令.filter((n) => !在る[n]);
const 有 = 命令.filter((n) => 在る[n]);

console.log('★Range の 命令（COM の 呼び出し用と 影を 除く）= ' + 命令.length + '個★');
console.log('★うちに 在る = ' + 有.length + '個（' + (有.length / 命令.length * 100).toFixed(1) + '%）★');
console.log('★無い       = ' + 無い.length + '個★');
console.log('');
console.log('── 在る ──');
有.forEach((n) => console.log('  ○ ' + n.padEnd(26) + 在る[n]));
console.log('');
console.log('── ★無い★ ──');
無い.forEach((n) => console.log('  ✕ ' + n));
