/* version-scope.test.mjs — ★台帳に書いた「版対応はここまで」が実態と合っているかを見張る★
 *
 * なぜ必要か（2026-08-01）:
 *   「版対応」は客にも投資家にも説明する言葉なので、書いてある事と実装がズレたら
 *   それは説明の誤りになる。そこで台帳(known-diffs.json の version_scope)の主張を機械で照合する。
 *
 *   今の主張は2つ:
 *     ① golden（真値）は Excel 16.0.20228 の【1版だけ】。他の版の実機突合は未実施。
 *     ② book.html の版セレクタは【表示だけ】で、計算にも xlsx 書き出しにも効いていない。
 *
 *   ②が特に大事。実装を足して「版で出し分ける」ようにしたのに台帳が古いままだと
 *   逆向きの嘘になるし、逆に台帳だけ直して実装を忘れても嘘になる。どちらでも赤にする。
 *
 * 使い方: node tests/xlsx-harness/version-scope.test.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..', '..');
const KD = JSON.parse(fs.readFileSync(path.join(__dirname, 'known-diffs.json'), 'utf8'));
const VS = KD.version_scope;

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); } };

console.log('\n[version-scope] 「版対応はここまで」の記述と実装が合っているか');

// ── ① goldenは何版あるか ──
const goldens = fs.readdirSync(path.join(__dirname, 'golden')).filter(f => /^excel-.*\.json$/.test(f));
T('① goldenの本数と台帳の記述が合っている（1版だけなら「1版だけ」と書いてある）', () => {
  if (!VS || !VS.golden_source) throw new Error('known-diffs.json に version_scope.golden_source が無い');
  const one = /この1版だけ/.test(VS.golden_source);
  if (goldens.length === 1 && !one) throw new Error('goldenは1本（' + goldens[0] + '）なのに、台帳が「1版だけ」と書いていない');
  if (goldens.length > 1 && one) throw new Error('goldenが ' + goldens.length + ' 本あるのに、台帳が「1版だけ」のまま: ' + goldens.join(', '));
});
T('① goldenのmetaに書いてある版が、台帳の記述と一致している', () => {
  const g = JSON.parse(fs.readFileSync(path.join(__dirname, 'golden', goldens[0]), 'utf8'));
  const v = g.meta && g.meta.version;
  if (!v) throw new Error('golden の meta.version が無い');
  if (VS.golden_source.indexOf(v) < 0) throw new Error('台帳に版 ' + v + ' が書かれていない: ' + VS.golden_source);
});
T('① 「実機で突合していない版がある」ことが明記されている', () => {
  const nv = (VS.not_verified || []).join(' ');
  for (const v of ['2016', '2019', '2021']) {
    if (nv.indexOf(v) < 0) throw new Error('not_verified に ' + v + ' の記載が無い');
  }
});

// ── ② 版セレクタが何かに効いているか ──
//   ★2026-08-02 (a-lite) で「飾り」ではなくなった。今は【警告】に効く（計算そのものは今も版を見ない）。
//   台帳の記述と実装が【どちら向きにズレても】赤にする、という趣旨は変えていない。
const book = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
const formula = fs.readFileSync(path.join(ROOT, 'exally-formula.js'), 'utf8');
const KEY = 'exally_excel_version';
// 計算そのものが版を見ているか（＝答えが版で変わるか）
const engineReadsVersion = new RegExp(KEY).test(formula) || /getBookVer\s*\(\s*\)/.test(formula);
// 表示用の3箇所を超えて版を読んでいるか（＝警告など、表示以外の用途に使っているか）
const UI_ONLY_CALLS = 3;   // getBookVer本体 / updateBookVerBadge / openBookVerModal
const callSites = (book.match(/getBookVer\s*\(\s*\)/g) || []).length;
const usedBeyondUI = callSites > UI_ONLY_CALLS;

const s = VS.product_version_selector || {};
const claimsWarnOnly = /警告/.test(s['実測 2026-08-02(a-lite実装後)'] || '');
const claimsCosmetic = /表示だけ/.test(s['実測 2026-08-02(a-lite実装後)'] || '');

T('② 版セレクタの実態と台帳の記述が一致している（どちら向きのズレも赤）', () => {
  if (!VS.product_version_selector) throw new Error('version_scope.product_version_selector が無い');
  if (usedBeyondUI && claimsCosmetic) {
    throw new Error('版を表示以外(' + callSites + '箇所)で使っているのに、台帳は「表示だけ」のまま。台帳を更新すること。');
  }
  if (!usedBeyondUI && claimsWarnOnly) {
    throw new Error('台帳は「警告に効く」と書いているのに、版を表示用の' + UI_ONLY_CALLS + '箇所でしか読んでいない。'
      + '実装が外れたか、台帳が先走っている。');
  }
});
T('② ★計算そのものが版で変わるようになったら台帳を書き直す（今は変わらないはず）', () => {
  const claimsCalcUnchanged = /計算そのものは今も版を見ない/.test(s['実測 2026-08-02(a-lite実装後)'] || '');
  if (engineReadsVersion && claimsCalcUnchanged) {
    throw new Error('計算側(exally-formula.js)が版を見るようになっているのに、'
      + '台帳は「計算そのものは版を見ない」のまま。どちらが正しいか決めて揃えること。');
  }
});
T('② 「まだ決めていない」ことに期限と選択肢が書いてある', () => {
  const s = VS.product_version_selector;
  if (!s.due) throw new Error('due（期限）が無い');
  if (!s['判断すること'] || s['判断すること'].length < 20) throw new Error('判断すること（選択肢）が書かれていない');
});
T('版を足す手順が書いてある（次の人が同じ事を調べ直さない）', () => {
  if (!VS.how_to_add_a_version) throw new Error('how_to_add_a_version が無い');
});

console.log('\n── 実測 ──');
console.log('  golden: ' + goldens.length + '本 (' + goldens.join(', ') + ')');
console.log('  計算そのものが版で変わるか: ' + (engineReadsVersion ? '★はい' : 'いいえ（答えは版によらず同じ）'));
console.log('  版セレクタの用途: ' + (usedBeyondUI ? '★警告に効いている(' + callSites + '箇所で参照)' : '表示だけ'));
console.log('  getBookVer() の参照: ' + callSites + '箇所');
console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
