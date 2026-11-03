/* ai-hoshi-hakaru.mjs -- ★AI に渡す 頼み文の★を 外すと、答えの★と 答えの形が どう 変わるか★（2026-10-10）
 *
 *  ★お金を 使う 測り★（本番と 同じ 鍵＝客の AI と 同じ 財布）。司さんの 許し（10-10「残高を見てから呼ぶ」→ クレジット購入）。
 *  ★1回だけ 回す★＝答えは 全部 ai-hoshi-kotae-2026-10-10.json に 残し、読み返しは そこから（2回目は 回さない）。
 *
 *  当て方：
 *    ・問い＝ai-hoshi-toi-2026-10-10.tsv の 23本（走らせる 前に commit 済み）
 *    ・頼み文＝api/claude.js の buildPromptParts を そのまま 呼ぶ（版 latest＝Excel 365）。
 *      法定の 行は ★本番の 倉庫を 読まず★ tests/api-claude.test.mjs と 同じ 作り物の 行（数字を 言うな の 守りの 文が 出る 形）
 *    ・A＝今の 頼み文 ／ B＝★だけ 外した 頼み文（共通・版ごと とも）。問いごとに 交互（奇数は A→B・偶数は B→A）
 *    ・model / max_tokens / system の 置き方（1時間）は 本番（api/claude.js 743-776）と 同じ。temperature は 送らない（本番と 同じ）
 *    ・api/claude.js の handler は 通さない（倉庫・記録・回数の 止めに 触れない）
 *  ★上限★：使った 額＋次の 1回の 一番 高い 見積もり が 上限（既定 5 ドル）を 超えるなら 止める。
 *    値段（ドル／百万トークン）：入力 3・出力 15・1時間 置く 6・読み直し 0.3（Sonnet の 公表値＝変わったら ここを 直す）
 *  使い方: node docs/measured/ai-hoshi-hakaru.mjs --鍵=<ANTHROPIC_API_KEY= の 行が 在る ファイル> [--上限=5]
 *  ★鍵の 中身は 出さない／書かない★
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const 引数 = process.argv.slice(2);
const 取る = (名) => (引数.find((a) => a.startsWith('--' + 名 + '=')) || '').slice(名.length + 3);
const 上限 = Number(取る('上限') || 5);
const 鍵の置き場 = 取る('鍵');
const 出す先 = path.join(ROOT, 'docs', 'measured', 'ai-hoshi-kotae-2026-10-10.json');
const 値段 = { 入力: 3, 出力: 15, 置く1時間: 6, 読み直し: 0.3 };
const MODEL = 'claude-sonnet-4-6', MAX = 2000;

process.exitCode = 1;
if (fs.existsSync(出す先)) { console.log('★もう 測ってある★（' + path.relative(ROOT, 出す先) + '）＝2回目は 回さない'); process.exit(1); }
const 鍵行 = String(fs.readFileSync(鍵の置き場, 'utf8')).split(/\r?\n/).find((l) => /^ANTHROPIC_API_KEY=/.test(l));
const 鍵 = 鍵行 ? 鍵行.slice('ANTHROPIC_API_KEY='.length).replace(/^"|"$/g, '') : '';
if (!鍵) { console.log('★鍵が 読めない★（中身は 出さない）'); process.exit(1); }

const handler = require_(path.join(ROOT, 'api', 'claude.js'));
const 行 = [
  { kind: 'shakaihoken', year: 2026, data: { kenko_total: { tokyo: 0.0991 }, kosei_total: 0.183 } },
  { kind: 'koyo', year: 2026, data: { ippan: 0.0055 } },
  { kind: 'shouhizei', year: 2019, data: { hyojun: 0.1, keigen: 0.08 } },
  { kind: 'saitei_chingin', year: 2025, data: {} },
  { kind: 'rousai_ritsu', year: 2024, data: {} },
];
const 部品 = handler.__buildPromptParts({ name: 'Excel 365', group: 'latest' }, 行);
if (!/数字を 言っては いけない/.test(部品.共通)) { console.log('★守りの 文が 頼み文に 無い★'); process.exit(1); }
const 型 = {
  A: { 共通: 部品.共通, 版ごと: 部品.版ごと },
  B: { 共通: 部品.共通.replace(/★/g, ''), 版ごと: 部品.版ごと.replace(/★/g, '') },
};
const 数 = (s, re) => (String(s).match(re) || []).length;
console.log('頼み文 A：字 ' + (型.A.共通.length + 型.A.版ごと.length) + '・★ ' + 数(型.A.共通 + 型.A.版ごと, /★/g)
  + ' ／ B：字 ' + (型.B.共通.length + 型.B.版ごと.length) + '・★ ' + 数(型.B.共通 + 型.B.版ごと, /★/g));

const 問い = fs.readFileSync(path.join(ROOT, 'docs', 'measured', 'ai-hoshi-toi-2026-10-10.tsv'), 'utf8')
  .split(/\r?\n/).filter((l) => /^\d+\t/.test(l)).map((l) => { const [番号, 材料, 問] = l.split('\t'); return { 番号: +番号, 材料, 問 }; });
if (問い.length !== 23) { console.log('★問いが 23本で ない★（' + 問い.length + '）'); process.exit(1); }

const Anthropic = require_('@anthropic-ai/sdk');
const client = new (Anthropic.default || Anthropic)({ apiKey: 鍵 });
const 額 = (u) => ((u.input_tokens || 0) * 値段.入力 + (u.output_tokens || 0) * 値段.出力
  + (u.cache_creation_input_tokens || 0) * 値段.置く1時間 + (u.cache_read_input_tokens || 0) * 値段.読み直し) / 1e6;
const 一番高い次 = (12000 * 値段.置く1時間 + MAX * 値段.出力) / 1e6;

const 記録 = [];
let 使った = 0, 止めた = '';
outer:
for (const q of 問い) {
  for (const t of (q.番号 % 2 ? ['A', 'B'] : ['B', 'A'])) {
    if (使った + 一番高い次 > 上限) { 止めた = '上限 ' + 上限 + ' ドルの 手前で 止めた（使った ' + 使った.toFixed(4) + '）'; break outer; }
    let r;
    try {
      r = await client.messages.create({
        model: MODEL, max_tokens: MAX,
        system: [
          { type: 'text', text: 型[t].共通, cache_control: { type: 'ephemeral', ttl: '1h' } },
          { type: 'text', text: 型[t].版ごと, cache_control: { type: 'ephemeral', ttl: '1h' } },
        ],
        messages: [{ role: 'user', content: q.問 }],
      });
    } catch (e) {
      止めた = q.番号 + t + ' で 転んだ：' + String((e && e.status) || '') + ' ' + String((e && e.message) || e).split('\n')[0].slice(0, 120);
      break outer;
    }
    const 答え = (r.content || []).filter((c) => c.type === 'text').map((c) => c.text).join('');
    const u = r.usage || {};
    使った += 額(u);
    記録.push({
      番号: q.番号, 材料: q.材料, 型: t, 答え, stop_reason: r.stop_reason, usage: u,
      星: 数(答え, /★/g), 字: 答え.length, コード: Math.floor(数(答え, /^```/gm) / 2), 等号行: 数(答え, /^\s*=/gm),
      率や額: 数(答え, /\d+(?:\.\d+)?\s*(?:%|％|パーセント)|\d{2,4}\s*円/g),
    });
    console.log(String(q.番号).padStart(2) + t + ' ★' + 記録[記録.length - 1].星 + ' 字' + 答え.length + ' 出力' + (u.output_tokens || 0) + ' 累計 $' + 使った.toFixed(4));
  }
}
fs.writeFileSync(出す先, JSON.stringify({ 測った日: new Date().toISOString(), model: MODEL, max_tokens: MAX, 上限, 使ったドル: +使った.toFixed(4), 止めた: 止めた || null, 値段, 記録 }, null, 1) + '\n');

const まとめ = (t, 絞り) => {
  const x = 記録.filter((r) => r.型 === t && 絞り(r));
  const 和 = (k) => x.reduce((a, r) => a + r[k], 0);
  return '★入り ' + x.filter((r) => r.星 > 0).length + '/' + x.length + '・★計 ' + 和('星') + '・字の平均 ' + Math.round(和('字') / (x.length || 1))
    + '・コード ' + 和('コード') + '・= 行 ' + 和('等号行') + '・率や額 ' + 和('率や額') + '・切れた ' + x.filter((r) => r.stop_reason === 'max_tokens').length;
};
const 普通 = (r) => r.材料 !== 'okane', お金 = (r) => r.材料 === 'okane';
console.log('\n[20本] A ' + まとめ('A', 普通) + '\n[20本] B ' + まとめ('B', 普通));
console.log('[お金3本] A ' + まとめ('A', お金) + '\n[お金3本] B ' + まとめ('B', お金));
const 和u = (k) => 記録.reduce((a, r) => a + (r.usage[k] || 0), 0);
console.log('使った：入力 ' + 和u('input_tokens') + '・置いた ' + 和u('cache_creation_input_tokens') + '・読み直し ' + 和u('cache_read_input_tokens')
  + '・出力 ' + 和u('output_tokens') + ' トークン ＝ $' + 使った.toFixed(4) + '（' + 記録.length + '/46 回）');
if (止めた) { console.log('★途中で 止まった★ ' + 止めた); } else if (記録.length === 46) process.exitCode = 0;
