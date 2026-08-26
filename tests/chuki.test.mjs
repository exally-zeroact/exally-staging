/* chuki.test.mjs — ★注記を外してから読む（見張りの共通部品）★
 *
 *  ★指示役 2026-08-26「同じ型を3回 踏んだら 決まりにする。1本ずつ直すと 4回目が来る」★
 *
 *  使い方: node tests/chuki.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { fileURLToPath, pathToFileURL } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const SELF = process.argv.includes('--self-test');
const OVERRIDE = process.env.EXALLY_CHUKI_OVERRIDE ? JSON.parse(process.env.EXALLY_CHUKI_OVERRIDE) : {};
const 部品 = await import(pathToFileURL(OVERRIDE['scripts/lib/chuki.mjs'] || path.join(ROOT, 'scripts/lib/chuki.mjs')).href);
const { 注記を外す, 動く所に在る } = 部品;

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ok   ' + n); } catch (e) { fail++; console.log('  NG   ' + n + '\n       ' + (e && e.message)); } };
const ok = (c, m) => { if (!c) throw new Error(m || 'expected truthy'); };
const eq = (a, b, m) => { if (a !== b) throw new Error((m ? m + ': ' : '') + '期待=' + JSON.stringify(b) + ' 実際=' + JSON.stringify(a)); };

console.log('');
console.log('[chuki] ★注記を外してから読む★');

T('★/* … */ を 外す★', () => {
  ok(!動く所に在る('/* env: "prod" */ var a = 1;', 'prod'));
  ok(動く所に在る('var env = "prod";', 'prod'));
});
T('★// 行末までを 外す★', () => {
  ok(!動く所に在る('var a = 1; // env: prod\nvar b = 2;', 'prod'));
  ok(動く所に在る('var a = 1; // env: prod\nvar b = "prod";', 'prod'));
});
T('★行の頭でない // も 外す（前は 行頭だけ見ていた）★', () => {
  eq(注記を外す('a; // だめ', { 残す: false }), 'a; ');
});
T('★★字の中の // は 注記ではない（URL）★', () => {
  ok(動く所に在る("var u = 'https://exally.vercel.app/prod';", 'prod'),
    '★URL の // を 注記と読んで 中身を捨てている★');
});
T('★字の中の /* は 注記ではない★', () => {
  ok(動く所に在る('var s = "/* prod */";', 'prod'));
});
T('★逃がした引用符（バックスラッシュ）で 字の終わりを 間違えない★', () => {
  /* ★字の中に /* … *​/ が在っても 注記ではない★。
     逃がし方を見落とすと ★字がそこで終わったと思い、中身を注記として捨てる★。 */
  const src = 'var s = ' + String.fromCharCode(39) + 'あ' + String.fromCharCode(92) + String.fromCharCode(39) + '/* AI_LIMIT */い' + String.fromCharCode(39) + ';';
  ok(動く所に在る(src, 'AI_LIMIT'), '★字の中を 注記と読んで 捨てた★');
});
T('★行と桁を ずらさない（場所を出す検査が狂わない）★', () => {
  const src = 'var a = 1;\n/* ここは\n   注記 */\nvar b = 2;\n';
  const 出 = 注記を外す(src);
  eq(出.split('\n').length, src.split('\n').length, '★行数が 変わった★');
  eq(出.split('\n')[3], 'var b = 2;');
  eq(出.indexOf('var b'), src.indexOf('var b'), '★桁が ずれた★');
});
T('★終わりが無い /* でも 落ちない★', () => {
  eq(注記を外す('a; /* 終わらない', { 残す: false }), 'a; ');
});
T('★SQL の -- を 外す（sql:true の時だけ）★', () => {
  ok(動く所に在る('-- drop table は しない\ncreate table a;', 'drop'), 'JSの時は -- を外さない');
  ok(!動く所に在る('-- drop table は しない\ncreate table a;', 'drop', { sql: true }));
  ok(動く所に在る("insert into t values ('--drop');", 'drop', { sql: true }), '★字の中の -- を 外している★');
});
T('★HTML の <!-- --> を 外す（html:true の時だけ）★', () => {
  ok(動く所に在る('<!-- 昔は AI_LIMIT だった --><div>x</div>', 'AI_LIMIT'));
  ok(!動く所に在る('<!-- 昔は AI_LIMIT だった --><div>x</div>', 'AI_LIMIT', { html: true }));
});
T('★空・null でも 落ちない★', () => {
  eq(注記を外す(''), '');
  eq(注記を外す(null), '');
  eq(注記を外す(undefined), '');
});
T('★注記の外は 1バイトも変えない★', () => {
  const src = 'var a = "★あ★";\nvar b = 2;\n';
  eq(注記を外す(src), src);
});

T('★閉じていない ひとつだけの 引用符で 字の中に居続けない（正規表現の中で踏んだ）★', () => {
  /* ★2026-08-26 実際に踏んだ★ … 正規表現 /[^']/ の ' で 字の中に入ったまま になり、
     その後ろの注記を 落とし損ねて book.html の1行が 残った。 */
  const src = 'var re = /[^' + String.fromCharCode(39) + ']/;' + String.fromCharCode(10)
    + '/* この注記は 落ちるべき AI_LIMIT */' + String.fromCharCode(10) + 'var b = 2;';
  ok(!動く所に在る(src, 'AI_LIMIT'), '★字の中に居続けて 注記を落とし損ねた★');
});
T('★バッククォートの字は 何行でも またげる★', () => {
  const B = String.fromCharCode(96);
  const src = 'var t = ' + B + 'あ' + String.fromCharCode(10) + 'い // これは字の中' + B + ';' + String.fromCharCode(10) + 'var b = 2;';
  ok(動く所に在る(src, 'これは字の中'), '★複数行の字を 切ってしまった★');
});

/* ══ ★実物で 使われているか（配ったか）★ ══ */
T('★見張りたちが この部品を使っている（1本ずつ書かない）★', () => {
  const 使う人 = ['tests/ai-reason.test.mjs', 'tests/book-scan-ui.test.mjs', 'tests/xlsm-vba.test.mjs'];
  const 足りない = [];
  for (const rel of 使う人) {
    const src = fs.readFileSync(path.join(ROOT, rel), 'utf8');
    if (src.indexOf('chuki.mjs') < 0) 足りない.push(rel);
  }
  eq(足りない.join(','), '', '★まだ 自前で書いている★');
});
T('★自前の「注記外し」が 見張りの中に 残っていない★', () => {
  const 残り = [];
  const 見る = fs.readdirSync(path.join(ROOT, 'tests')).filter((f) => f.endsWith('.mjs'));
  for (const f of 見る) {
    if (f === 'chuki.test.mjs') continue;
    const src = 注記を外す(fs.readFileSync(path.join(ROOT, 'tests', f), 'utf8'));
    /* ★自前で /*…*​/ を消している所★（この部品を使えば要らない） */
    if (/replace\(\s*\/\\\/\\\*\[/.test(src)) 残り.push('tests/' + f);
  }
  eq(残り.join(','), '', '★まだ 自前で書いている所が 在る★');
});

console.log('');
console.log('  ' + pass + ' 緑 / ' + fail + ' 赤');

if (SELF) {
  const { spawnSync } = await import('node:child_process');
  const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'exally-chuki-'));
  console.log('');
  console.log('[self-test] わざと壊して 赤くなるかを数える（★repo は読むだけ★）');
  const BREAKS = [
    ['★字の中を 見ない（URL の // を 注記にする）★',
      (s) => s.replace("    if (c === \"'\" || c === '\"' || c === '`') {", '    if (false) {')],
    ['★// を 外さない★', (s) => s.replace("if (行の注記 && c === '/' && c2 === '/') {", "if (false) {")],
    ['★/* を 外さない★', (s) => s.replace("    if (c === '/' && c2 === '*') {", '    if (false && c2 === \'*\') {')],
    ['★行と桁を ずらす（空白で埋めない）★', (s) => s.replace('  const 埋める = (opt.残す === undefined) ? true : !!opt.残す;', '  const 埋める = false;')],
    ['★SQL でないのに -- を 外す★', (s) => s.replace("if (o.sql && c === '-' && c2 === '-') {", "if (c === '-' && c2 === '-') {")],
    ['★HTMLでない字まで <script/<style で切り分ける★',
      (s) => s.replace('  if (!opt.html) return 素で外す(s, { sql: !!opt.sql, 埋める: 埋める });', '')],
    ['★閉じていない 引用符でも 字の中だと思い続ける★',      (s) => s.replace('      const 端 = 字の終わり(s, i);', '      const 端 = i + 1;')],
    ['★` を 1行で閉じさせる（複数行の字が 切れる）★',
      (s) => s.replace("  const 行末 = (q === '`') ? s.length :", "  const 行末 = (false) ? s.length :")],
    ['★終わりが無い /* で 落ちる★', (s) => s.replace('      const 端 = (終 < 0) ? n : 終 + 2;', '      const 端 = 終 + 2;')],
    ['★逃がした引用符で 字の終わりを 間違える★',
      (s) => s.replace('    if (s[j] === String.fromCharCode(92)) { j += 2; continue; }', '')],
  ];
  let red = 0;
  for (const [name, brk] of BREAKS) {
    const orig = fs.readFileSync(path.join(ROOT, 'scripts/lib/chuki.mjs'), 'utf8');
    const bad = brk(orig);
    if (bad === orig) { console.log('  ★置換できず★  ' + name); continue; }
    const f = path.join(TMP, 'chuki.mjs');
    fs.writeFileSync(f, bad, 'utf8');
    const env = Object.assign({}, process.env, { EXALLY_CHUKI_OVERRIDE: JSON.stringify({ 'scripts/lib/chuki.mjs': f }) });
    const r = spawnSync(process.execPath, ['--max-old-space-size=4096', path.join(__dirname, 'chuki.test.mjs')], { encoding: 'utf8', env });
    if (r.status !== 0) { red++; console.log('  赤くなった  ' + name); }
    else console.log('  ★素通り★  ' + name);
  }
  try { fs.rmSync(TMP, { recursive: true, force: true }); } catch (e) { /* 消せなくても検査は済んでいる */ }
  console.log('');
  console.log('  ' + red + '/' + BREAKS.length + ' 通りで赤くなった');
  process.exit(red === BREAKS.length ? 0 : 1);
}

process.exit(fail ? 1 : 0);
