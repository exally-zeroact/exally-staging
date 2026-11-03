/* ribbon-label.test.mjs — ★リボンの 札が 箱に 収まり、同じに 見えない★ 2026-08-30
 *
 *  ★なぜ★（監査役の 差し戻し・2026-08-30）
 *    絵を 開いて 見たら ★「ウィンドウ…」が 3つ 並んでいた★＝
 *      ウィンドウ枠の固定／ウィンドウの位置を元に戻す／ウィンドウの切り替え
 *    ★お客さんには どれが どれか 分かりません★。
 *    ★数字では 出ませんでした（絵を 見たから 分かった）★
 *    ⇒ ★機械で 数える形に して、また 伸びたら 赤に する★
 *
 *  ★見る物（3つ）★
 *    ① 札の 幅が 上限（全角6字）を 超えていない … 超えると CSS が 切る＝頭だけ 残る
 *    ② ★同じ組の 中で 同じ札が 2つ 出ていない★（リボン ★全部★ で 見る）
 *    ③ 元の 名前を 消していない（title に 残っている）
 *
 *  ★空振りを 止める★
 *    ・数えた 部品が 200個 未満なら 赤（読めていない）
 *    ・--self-test で ★わざと 長い札を 入れて 赤に なるか★ 見る
 *
 *  走らせ方: node tests/ribbon-label.test.mjs  ／  --self-test
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { JSDOM } from 'jsdom';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
let ok = 0, ng = 0;
const 言う = (よい, 文, 添え) => {
  if (よい) { ok++; console.log('  ok   ' + 文); }
  else { ng++; console.log('  NG   ' + 文); if (添え) console.log('       ' + 添え); }
};

/* ── 読む ─────────────────────────────────────── */
const ラベル = (await import('file://' + path.join(ROOT, 'lib/ribbon-label.js').replace(/\\/g, '/'))).default;
const specSrc = fs.readFileSync(path.join(ROOT, 'lib/ribbon-spec.js'), 'utf8');

const 部品 = [];
for (const 行 of specSrc.split(String.fromCharCode(10))) {
  /* ★s（親の字）が 付く 行が 在る★ 2026-09-03
     ＝1つの 箱の 中で 親の字の 下に 並ぶ 形（シートのオプション の 枠線／見出し）。
     ★s を 数えないと この4行が 丸ごと 抜け落ちて 見張りが 空振りする★ */
  const m = /\{ t: '([^']*)', g: '([^']*)', p: '([^']*)', (?:s: '([^']*)', )?a: (.*)\},?\s*$/.exec(行.trim());
  if (!m) continue;
  部品.push({ t: m[1], g: m[2], p: m[3], s: m[4] || null, 結んだ: m[5].trim() !== 'null' });
}
const 結んだ = 部品.filter((v) => v.結んだ);

console.log('★リボンの 札★');
言う(結んだ.length >= 200, '★空振りしていない★（結んだ部品 ' + 結んだ.length + '個 を 読めた）',
  結んだ.length < 200 ? '200個 未満＝読めていない' : '');

/* ── ① 幅 ───────────────────────────────────── */
function 札を出す(v, n) { return ラベル.札(v.p, v.t, v.g, n); }

const 何個目 = {};
const 一覧 = 結んだ.map((v) => {
  const k = v.t + '|' + v.g + '|' + ラベル._尻を落とす(v.p);
  何個目[k] = (何個目[k] || 0) + 1;
  return { ...v, 札: 札を出す(v, 何個目[k]) };
});

const 長すぎ = 一覧.filter((v) => ラベル.幅(v.札) > ラベル.上限);
言う(長すぎ.length === 0,
  '★札が 箱に 収まっている（全角' + ラベル.上限 + '字まで）＝はみ出し 0個★',
  長すぎ.length ?長すぎ.slice(0, 8).map((v) => v.t + '|' + v.g + '|' + v.札
    + '（幅' + ラベル.幅(v.札) + '）').join(' / ') + (長すぎ.length > 8 ? ' …他' + (長すぎ.length - 8) + '個' : '') : '');

/* ── ② 同じ組の 中で 同じ札が 出ていないか（リボン全部） ── */
const 組ごと = {};
for (const v of 一覧) {
  /* ★親の字（s）が 在る物は 列ごとに 見る★ 2026-09-03
     ＝実Excel の シートのオプションは 1つの 箱の 中に
       ★枠線[表示][印刷]／見出し[表示][印刷]★ が 2列で 入る。
     ★「表示」が 2つ 在っても 親の字が 違えば 見分けが つく★（それが Excel の 形）。
     ★親の字が 同じ 中で 重なったら 今までどおり 赤★＝線は 緩めていない。 */
  const k = v.t + '|' + v.g + (v.s ? '|' + v.s : '');
  (組ごと[k] = 組ごと[k] || []).push(v);
}
const ぶつかり = [], 許した = [];
for (const k of Object.keys(組ごと)) {
  const 見え = {};
  for (const v of 組ごと[k]) (見え[v.札] = 見え[v.札] || []).push(v.p);
  for (const 札 of Object.keys(見え)) {
    if (見え[札].length <= 1) continue;
    /* ★理由を 書いた物だけ 許す★（実Excelの 分かれボタン＝UIAが 同じ場所で 2回 返す） */
    const 訳 = ラベル.重なってよい[k + '|' + 札];
    if (訳) { 許した.push(k + ' … 「' + 札 + '」：' + 訳); continue; }
    ぶつかり.push(k + ' … 「' + 札 + '」が ' + 見え[札].length + '個（' + 見え[札].join(' / ') + '）');
  }
}
言う(ぶつかり.length === 0,
  '★同じ組の 中で 同じに 見える札 = 0件★（組 ' + Object.keys(組ごと).length + '個 を 全部 見た）',
  ぶつかり.slice(0, 6).join('\n       '));

/* ★黙って 見逃さない★＝許した物は 必ず 名前を 出す */
console.log('       …… ★理由を 書いて 許した 重なり = ' + 許した.length + '件★');
許した.forEach((v) => console.log('          ' + v));
言う(許した.length === Object.keys(ラベル.重なってよい).length,
  '★許可の 行が 全部 使われている（死んだ許可が 残っていない）＝'
    + Object.keys(ラベル.重なってよい).length + '行★',
  '使っていない 許可を 残すと ★次の 本物を 見逃す★');

/* ── ③ 元の 名前を 消していない ───────────────── */
const ribbonSrc = fs.readFileSync(path.join(ROOT, 'lib/ribbon.js'), 'utf8');
言う(/title="'\s*\+\s*esc\(元\)/.test(ribbonSrc),
  '★元の 名前は title に 残している（短くした札で 上書きしていない）★',
  '当てると 出る 字まで 短くすると 探せなくなる');

/* ── ★実際に 描いて title を 見る★（監査役の 宿題・2026-08-30）──
 *  ★札を 短く したので、★実Excelの 名前は title に 必ず 残っている★事を
 *  ★ソースを 読むのでは なく 描いて 確かめる★。
 *  （1個でも 出なければ それは ★探せない 札★） */
{
  const dom = new JSDOM('<div id="rb"></div>');
  global.self = dom.window; global.window = dom.window; global.document = dom.window.document;
  const Spec = (await import(pathToFileURL(path.join(ROOT, 'lib/ribbon-spec.js')).href)).default;
  dom.window.RibbonActions = new Proxy({}, { get: () => function () {} });
  const Ribbon = (await import(pathToFileURL(path.join(ROOT, 'lib/ribbon.js')).href)).default;
  const el = dom.window.document.getElementById('rb');
  const タブ = Spec.ツリー().map((t) => t.name);
  let 見た = 0, 空 = 0;
  const 違う = [];
  for (const n of タブ) {
    Ribbon.状態.tab = n;
    Ribbon.描く(el, Spec);
    for (const g of el.querySelectorAll('.rb-group')) {
      /* ★並び順では 突き合わせない★＝JSDOM では 引き取り枠（サイズの入力）が
         「元の物が 無い」ので 消える＝★番号が 1つ ずれる★（2026-08-30 実際に 踏んだ）。
         ⇒ ★「この組の 正本の 名前の どれかに 当たるか」で 見る★。 */
      const 期待 = new Set(一覧.filter((v) => v.t === n && v.g === g.dataset.group)
        .map((v) => ラベル._尻を落とす(v.p)));
      for (const b of g.querySelectorAll('.rb-item, [data-take]')) {
        const ti = (b.getAttribute('title') || '').trim();
        見た++;
        if (!ti) { 空++; continue; }
        if (!期待.has(ti)) 違う.push(n + '|' + g.dataset.group + '：title「' + ti + '」が 正本に 無い');
      }
    }
  }
  言う(見た >= 200, '★描いて title を 見た（' + 見た + '個）★');
  言う(空 === 0, '★title が 空の ボタン = 0個★', String(空));
  言う(違う.length === 0, '★title は 実Excelの 名前の まま = 違い 0個★', 違う.slice(0, 5).join(' / '));
}

/* ── ④ 短い札の 表そのものに 重なりが 無いか ──── */
const 表 = ラベル._短い;
const 逆 = {};
for (const k of Object.keys(表)) (逆[表[k]] = 逆[表[k]] || []).push(k);
const 表の重なり = Object.keys(逆).filter((k) => 逆[k].length > 1);
言う(表の重なり.length === 0,
  '★短い札の 表の 中で 同じ字を 使い回していない = ' + Object.keys(表).length + '個★',
  表の重なり.slice(0, 5).map((k) => '「' + k + '」← ' + 逆[k].join(' / ')).join('\n       '));

/* ── わざと 壊して 赤に なるか ────────────────── */
if (process.argv.includes('--self-test')) {
  console.log('\n★わざと 壊して 赤に なるか★');
  const 前 = 表['ウィンドウ枠の固定'];
  表['ウィンドウ枠の固定'] = 'ウィンドウ枠の固定';      /* ★長い物に 戻す★ */
  const 幅NG = ラベル.幅(ラベル.札('ウィンドウ枠の固定')) > ラベル.上限;
  表['ウィンドウ枠の固定'] = '窓の切替';                 /* ★別の物と 同じに する★ */
  const 札A = ラベル.札('ウィンドウ枠の固定');
  const 札B = ラベル.札('ウィンドウの切り替え');
  const 同じNG = 札A === 札B;
  表['ウィンドウ枠の固定'] = 前;                         /* 戻す */
  言う(幅NG, '★長い札に 戻したら 幅で 引っかかる★');
  言う(同じNG, '★別の物と 同じ札に したら ぶつかりで 引っかかる★（' + 札A + ' = ' + 札B + '）');
  言う(ラベル.札('ウィンドウ枠の固定') === 前, '★戻した（後に 残していない）★');
}

console.log('\nribbon-label: ' + ok + '/' + (ok + ng) + ' passed');
process.exit(ng ? 1 : 0);
