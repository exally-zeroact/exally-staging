/* betsumei-zenbu.test.mjs — ★「打てば 動く 別名」を ★機械で 全部 出して★ 台帳と 突き合わせる★
 *                            （2026-09-06・指示役の 指摘）
 *
 *  ★★きっかけ（2026-09-05〜06）★★
 *    `YEN` を ★「動かない」の 棚★に 置いていた ⇒ 実際は ★¥1,235 が 出る★
 *    ⇒ AIが 客に「YEN は 動きません」と 言う 所だった。
 *    ★指示役★「★2人目が 居ます★＝`JIS` は YEN と 同じ 家です」
 *              「★そして 残りを 手で 挙げないで ください★
 *                ⇒ `convertFormula` の 直している 表を ★機械で 全部 出す★
 *                ⇒ ★数を 先に 言わない★＝★出してから 数える★」
 *
 *  ★★この 見張りが する 事★★
 *    ①`exally-formula.js` の `convertFormula` から
 *      ★お客さんが 打つ 名前★を ★機械で 拾う★（手で 並べない）
 *    ②1つずつ ★本番と 同じ道（convertFormula → engine）で 押す★
 *    ③★答えが 出る 物★は 台帳に 載っていないと ★赤★
 *       ・★「動かない」棚に 居たら 赤★（＝AIが 客に 嘘を 言う）
 *       ・★どの 棚にも 無くても 赤★（＝AIは その 名前を 知らない）
 *    ④★答えが 出ない 物★は 逆に「別名で動く」に 居たら ★赤★（逆向きの 嘘）
 *
 *  ★★手で 挙げない事が この 見張りの 芯★★
 *    ⇒ ★新しい 直しを 足したら、その日に ここが 赤に なる★
 *    ⇒ 私が 気づかなくても ★機械が 気づく★
 *
 *  ★重ならない（作る前に 探した）★
 *    `tests/kansuu-tana.test.mjs` … ★台帳の 棚★を engine に 押す（棚 → 実物）
 *    `tests/xlsx-harness/alias.test.mjs` … ★直しその物★（入口と 出口）
 *    ここ … ★実物（コードの 直し）→ 台帳★＝★向きが 逆★＝抜けを 見つける
 *
 *  使い方: node tests/betsumei-zenbu.test.mjs
 *          node tests/betsumei-zenbu.test.mjs --self-test
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(import.meta.url);
let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };

/** ★純関数①★＝`convertFormula` の 中身から ★直している 名前★を 機械で 拾う
 *  ★手で 並べない★＝ここが 芯。拾う 形は `/\bNAME\s*\(/` （＝関数の 呼び出しを 直す 形）。 */
export function 直している名前(中身) {
  const 出 = [];
  const re = /\/\\b([A-Z][A-Z0-9._]*)\\s\*\\\(/g;
  let m;
  while ((m = re.exec(中身))) if (出.indexOf(m[1]) < 0) 出.push(m[1]);
  return 出.sort();
}

/** ★純関数②★＝拾った 名前・押した 答え・台帳 から ★食い違い★を 返す */
export function 台帳の抜け(名たち, 押す, 棚を引く) {
  const 出 = { 動くのに動かない棚: [], 動くのにどの棚にも無い: [], 動かないのに動く棚: [] };
  for (const f of 名たち) {
    const 動く = 押す(f);
    const 棚 = 棚を引く(f);
    if (動く) {
      if (棚 === '動かない') 出.動くのに動かない棚.push(f);
      else if (!棚) 出.動くのにどの棚にも無い.push(f);
    } else if (棚 === '別名で動く') {
      出.動かないのに動く棚.push(f);
    }
  }
  return 出;
}

if (process.argv.includes('--self-test')) {
  console.log('\n[betsumei-zenbu --self-test] わざと壊して赤になるか');
  T('★★直している 名前を 手で 並べずに 拾える★★', () => {
    const 中 = "f = f.replace(/\\bJIS\\s*\\(/gi, 'DBCS(');\n"
      + "f = f.replace(/\\bYEN\\s*\\(/gi, 'DOLLAR(');\n"
      + "f = f.replace(/\\bTRUE\\b(?!\\s*\\()/g, 'TRUE()');";
    const r = 直している名前(中);
    if (r.join(',') !== 'JIS,YEN') throw new Error('拾えていない: ' + JSON.stringify(r));
  });
  T('★1つも 拾えなければ それが 分かる（空を 緑に しない）', () => {
    if (直している名前('なにも ない').length !== 0) throw new Error('拾いすぎ');
  });
  /* ★2026-09-05 に 実際に 起きた 形★ */
  T('★★動くのに「動かない」棚に 居たら 赤（YEN の 事故）★★', () => {
    const r = 台帳の抜け(['YEN'], () => true, () => '動かない');
    if (r.動くのに動かない棚.length !== 1) throw new Error('見つけていない');
  });
  /* ★2026-09-06 に 指示役が 見つけた 形★ */
  T('★★動くのに どの 棚にも 無かったら 赤（JIS の 抜け）★★', () => {
    const r = 台帳の抜け(['JIS'], () => true, () => null);
    if (r.動くのにどの棚にも無い.length !== 1) throw new Error('見つけていない');
  });
  T('★逆向き＝動かないのに「別名で動く」に 居たら 赤', () => {
    const r = 台帳の抜け(['X'], () => false, () => '別名で動く');
    if (r.動かないのに動く棚.length !== 1) throw new Error('見つけていない');
  });
  T('★正しければ 何も 出さない', () => {
    const r = 台帳の抜け(['YEN'], () => true, () => '別名で動く');
    if (r.動くのに動かない棚.length || r.動くのにどの棚にも無い.length || r.動かないのに動く棚.length) {
      throw new Error('誤検知: ' + JSON.stringify(r));
    }
  });
  console.log('\n' + pass + ' passed, ' + fail + ' failed');
  process.exit(fail ? 1 : 0);
}

/* ══ 本番＝拾って 押して 突き合わせる ═══════════════════════ */
const src = fs.readFileSync(path.join(ROOT, 'exally-formula.js'), 'utf8');
const 本体 = src.split('function convertFormula')[1].split('\nfunction ')[0];
const 名たち = 直している名前(本体);

const HFns = require_(path.join(ROOT, 'hyperformula.full.min.js'));
const EF = require_(path.join(ROOT, 'exally-formula.js'));
const 積めた = EF.registerExallyFunctions(HFns) === true;
const 台帳 = require_(path.join(ROOT, 'lib/formula-extra.js')).数える();
const hf = HFns.HyperFormula.buildEmpty({ licenseKey: 'gpl-v3' });
const SID = hf.getSheetId(hf.addSheet('S'));

/* ★引数は 名前ごとに 要る数が 違う★＝足りないと 本当は 動く物が #NA? に なる
   ⇒★分からない 物は 1つずつ 増やして 試す★（最大 5つ）＝★数を 決め打たない★ */
const 押した答え = {};
function 押す(f) {
  const 例 = f === 'JIS' ? ['"あ"'] : [];
  const 候補 = 例.length ? 例 : ['1', '1,1', '0.5,1,2', '1,2,3,4', '1,2,0.5,1', 'A1'];
  for (const 引数 of 候補) {
    const 後 = EF.convertFormula('=' + f + '(' + 引数 + ')');
    hf.setSheetContent(SID, [[1, 2, 3], [後]]);
    const v = hf.getCellValue({ sheet: SID, row: 1, col: 0 });
    if (!(v && v.type)) { 押した答え[f] = { 引数, 後, 答: String(v) }; return true; }
    押した答え[f] = { 引数, 後, 答: '#' + v.type + '?' };
  }
  return false;
}
function 棚を引く(f) {
  if (台帳.足す.indexOf(f) >= 0) return '足す';
  if (台帳.別のlibで足す && 台帳.別のlibで足す[f]) return '別lib';
  if (台帳.別名で動く && 台帳.別名で動く[f]) return '別名で動く';
  if (台帳.足さない[f]) return '動かない';
  return null;
}

console.log('\n[betsumei-zenbu] 打てば 動く 別名を 機械で 全部 出して 台帳と 突き合わせる');

T('★プラグインを 積めた（積み忘れると 素の engine が 答えて 嘘の 緑に なる）', () => {
  if (!積めた) throw new Error('registerExallyFunctions が true を 返さない');
});

T('★★機械が 名前を 拾えている（拾えないと 全部 素通り）★★', () => {
  if (名たち.length < 3) throw new Error('★' + 名たち.length + '個しか 拾えていない★'
    + '＝拾い方が 壊れています（convertFormula の 書き方が 変わった？）');
});

const 抜け = 台帳の抜け(名たち, 押す, 棚を引く);

T('★★打てば 動くのに「動かない」棚に 居ない（AIが 客に 嘘を 言わない）★★', () => {
  if (抜け.動くのに動かない棚.length) {
    throw new Error('★動くのに「動かない」と 台帳に 書いてあります★: ' + 抜け.動くのに動かない棚.join(' / ')
      + '\n   → ★AIが お客さんに「使えません」と 言います★');
  }
});

T('★★打てば 動く 名前が どの 棚にも 無い、が 無い（AIが 知らない を 無くす）★★', () => {
  if (抜け.動くのにどの棚にも無い.length) {
    throw new Error('★どの 棚にも 無い★: ' + 抜け.動くのにどの棚にも無い.join(' / ')
      + '\n   → ★AIは この 名前を 知りません★（客が 打てば 動くのに 教えられない）'
      + '\n   → lib/formula-extra.js の ★別名で動く★ に 1行 足してください');
  }
});

T('★逆向き＝動かないのに「別名で動く」に 居ない', () => {
  if (抜け.動かないのに動く棚.length) {
    throw new Error('★動かないのに「動く」と 書いてあります★: ' + 抜け.動かないのに動く棚.join(' / '));
  }
});

console.log('\n── 実測（★手で 並べず 機械が 拾った★） ──');
console.log('  拾った 名前 … ' + 名たち.length + '個');
for (const f of 名たち) {
  const a = 押した答え[f] || {};
  console.log('  ' + f.padEnd(13) + String(a.後 || '').slice(1, 32).padEnd(33)
    + String(a.答 || '').padEnd(16) + (棚を引く(f) || '★どこにも 無い★'));
}
console.log('\n' + pass + ' passed, ' + fail + ' failed');
hf.destroy();
process.exit(fail ? 1 : 0);
