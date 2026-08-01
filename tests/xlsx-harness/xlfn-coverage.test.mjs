/* xlfn-coverage.test.mjs — ★書き出す関数名が「分類済み」であることを機械で強制する★
 *
 * なぜ必要か（2026-08-01 第3波P2で実際に踏んだ）:
 *   新しい関数は xlsx の中では `_xlfn.` を付けた名前で保存する決まり。付け忘れると
 *   【Excelがその式を #NAME? にする／物によってはブックごと開けない】。
 *   実際に RANK.EQ は一覧に入っていたのに **RANK.AVG だけ抜けていて**、
 *   書き出したブックを実Excelで開いた時だけ その式が #NAME? になっていた。
 *   ★これは画面のテストでは絶対に見つからない。そして CI には Excel が無いので
 *     tools/verify-workbook-excel.ps1（実Excel）でしか気づけない＝人が走らせ忘れたら素通りする。
 *
 *   そこで「分類していない関数が式に出てきたら赤」にする。
 *   分類は3つのどれか:
 *     ① xlsx-io.js の XLFN 一覧    … _xlfn. を付ける新しい関数
 *     ② xlsx-io.js の ALIAS       … 日本語UIの表示名（JIS→DBCS / YEN→DOLLAR）
 *     ③ xlfn-legacy.json の legacy … 接頭辞の要らない古い関数
 *   どれにも無い関数は「まだ実Excelで確かめていない」という意味なので赤にする。
 *
 * 使い方: node tests/xlsx-harness/xlfn-coverage.test.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..', '..');
const require = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const IO = require(path.join(__dirname, 'xlsx-io.js'));
const LEG = JSON.parse(fs.readFileSync(path.join(__dirname, 'xlfn-legacy.json'), 'utf8'));

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); } };

console.log('\n[xlfn-coverage] 書き出す関数名が分類済みか（_xlfn. の付け忘れを止める）');

// ケースの式から関数名を集める（文字列リテラルの中は数えない）
const CASEDIR = path.join(__dirname, 'cases');
const used = new Set();
for (const f of fs.readdirSync(CASEDIR)) {
  if (f === '_inputs.json') continue;
  const j = JSON.parse(fs.readFileSync(path.join(CASEDIR, f), 'utf8'));
  for (const c of j.cases) {
    const s = c.f.replace(/"(?:[^"]|"")*"/g, '""');
    for (const m of s.matchAll(/([A-Z][A-Z0-9.]*)\s*\(/gi)) used.add(m[1].toUpperCase());
  }
}

const xlfn = new Set(IO.xlfnNames);
const alias = new Set(['JIS', 'YEN']);          // xlsx-io.js の ALIAS と対
const legacy = new Set(LEG.legacy);

const unclassified = [...used].filter(n => !xlfn.has(n) && !alias.has(n) && !legacy.has(n)).sort();

T('★分類していない関数が式に出てこない（出てきたら実Excelで確かめて振り分ける）', () => {
  if (unclassified.length) {
    throw new Error('分類されていない関数:\n   - ' + unclassified.join('\n   - ')
      + '\n   → tools/verify-workbook-excel.ps1 を実Excelで走らせ、'
      + '\n     その式が #NAME? になるなら xlsx-io.js の XLFN 一覧へ、ならないなら xlfn-legacy.json へ足す。');
  }
});

T('legacy 一覧に「新しい関数」が紛れていない（XLFNと重複していない）', () => {
  const both = [...legacy].filter(n => xlfn.has(n));
  if (both.length) throw new Error('両方に載っている: ' + both.join(', ') + '（どちらか片方にする）');
});

T('別名(表示名)は legacy にも XLFN にも入れない（本名に直してから書き出すため）', () => {
  const bad = [...alias].filter(n => legacy.has(n) || xlfn.has(n));
  if (bad.length) throw new Error('別名が一覧に入っている: ' + bad.join(', '));
});

T('検査が空振りしていない（式から関数名を実際に拾えている）', () => {
  if (used.size < 40) throw new Error('拾えた関数が少なすぎます: ' + used.size);
});

T('legacy 一覧に「いつ実Excelで確かめたか」が書いてある', () => {
  if (!LEG.verified || LEG.verified.length < 20) throw new Error('verified（実Excelで確かめた時点）が書かれていない');
  if (!LEG.howToAdd) throw new Error('howToAdd（足し方）が書かれていない');
});

console.log('\n── 実測 ──');
console.log('  式に出てくる関数: ' + used.size + '個');
console.log('  _xlfn. を付ける : ' + [...used].filter(n => xlfn.has(n)).sort().join(', '));
console.log('  日本語UIの表示名: ' + ([...used].filter(n => alias.has(n)).sort().join(', ') || '(ケースには出てこない=alias.test.mjsで見張る)'));
console.log('  接頭辞不要(古い): ' + [...used].filter(n => legacy.has(n)).length + '個');
console.log('  未分類          : ' + unclassified.length + '個' + (unclassified.length ? '\n   - ' + unclassified.join('\n   - ') : ''));

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
