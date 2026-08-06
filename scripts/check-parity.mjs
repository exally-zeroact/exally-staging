/* check-parity.mjs — ★本番(exally)とテスト(exally-staging)の食い違いを、中身で数える★
 *
 * なぜ必要か（2026-08-05・指示役）:
 *   本番だけに39件・stagingだけに39件あり、24時間ぶん離れていた。
 *   さらに悪いことに ★同じ物を両方で別々に作っていた★（入口の見張り）。
 *   これが続くと ★どちらが正か誰にも分からなくなる。★
 *
 * ★コミットの履歴ではなく「ファイルの中身」で比べる。★
 *   履歴で比べると、同じ物を別々に作った時に「両方に独自のコミットがある」としか出ず、
 *   中身が同じなのか違うのかが分からない。
 *
 * 比べる前にそろえる（ここを揃えないと、意味のない差で埋まって本物の差が見えない）:
 *   ・改行コード（CRLF / LF）… Windowsとgitの設定で勝手に変わる。中身の違いではない
 *   ・?v=<刻印> のキャッシュ避け … ビルドのたびに変わる。中身の違いではない
 *   実測(2026-08-05): 見かけの差50件 → そろえると ★15件は改行/刻印だけ★ で、本物は35件だった。
 *
 * ★片側にしか無くてよい物は、理由つきで一覧に書く。★
 *   理由が書けない差は許さない（＝一覧は「逃げ道」ではなく「決めた事の記録」）。
 *
 * 使い方: node scripts/check-parity.mjs [stagingの場所]   （既定 ../exally-staging）
 *         node scripts/check-parity.mjs --json
 *         node scripts/check-parity.mjs --self-test
 */
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');

/* ★片側にしか無くてよい／中身が違ってよい物。必ず理由を書く。 */
const ALLOWED = {
  // ── テスト(staging)にしか無くてよい物 ──
  '.nojekyll': 'GitHub Pages 専用。本番(Vercel)には要らない',
  'index.html': 'GitHub Pages のサブパス配信用の入口。本番は vercel.json の rewrites で hub.html を出す',
  'tests/pages-hosting.test.mjs': 'GitHub Pages のサブパス配信を守る検査。本番はサブパス配信ではない',
  'tests/repo-supa.mjs': 'テスト用の倉庫(DB-test)を見る道具。本番の倉庫には向けない',
  'tools/verify-grid-export-excel.ps1': '実Excelでグリッド書き出しを確かめる道具。開発機でだけ使う',

  // ── 中身が違ってよい物 ──
  'js/supa-config.js': '★倉庫の接続先。本番=本番の倉庫 / staging=DB-test。同じにしてはいけない',
  '.github/workflows/ci.yml': '走らせる検査が違う（stagingには第3波ハーネス、本番には入口/設定の守り）。★中身は追いかける',
  'manifest.json': 'PWAの入口URL。配信の場所が違うので start_url が違う',
  'kyuyo/manifest.json': '同上',
  'kyuyo/admin-manifest.json': '同上',
  'kyuyo/admin.html': "Service Worker の場所。staging はサブパス配信なので '../sw.js'、本番は '/sw.js'",
  'tests/run.js': '登録する検査が違う（staging に Pages配信ガードと版対応ハーネスがある）。★中身は追いかける',
  'tests/hub-ui.mjs': 'staging はサブパス配信なので「絶対パスで書かない」検査を余分に持つ',
  'tests/ci-coverage.test.mjs': 'CIから外している物の一覧が違う（staging=repo-supa / 本番=xlsx-harness の入出力部品）',
  'tests/xlsx-harness/report.md': '★走らせた日の数字が入る生成物（=TODAY() の日付シリアル）。日をまたぐと必ず違う＝中身の差ではない',
};

/* 比べる前にそろえる（改行コードと刻印は「中身の違い」ではない） */
export function normalize(text) {
  return String(text)
    .replace(/\r\n/g, '\n')                 // 改行コード
    .replace(/\?v=[0-9a-f]{6,}/g, '?v=');   // キャッシュ避けの刻印
}

/* ★純関数: 2つのファイル一覧と中身から、差を分類する。self-testで作り物を通せる。 */
export function classify(a, b, allowed) {
  const names = [...new Set([...Object.keys(a), ...Object.keys(b)])].sort();
  const out = { onlyProd: [], onlyStaging: [], differ: [], same: 0, allowedCount: 0 };
  for (const n of names) {
    const inA = Object.prototype.hasOwnProperty.call(a, n), inB = Object.prototype.hasOwnProperty.call(b, n);
    const why = allowed[n];
    if (inA && inB) {
      if (normalize(a[n]) === normalize(b[n])) { out.same++; continue; }
      if (why) { out.allowedCount++; continue; }
      out.differ.push(n);
    } else if (inA) {
      if (why) { out.allowedCount++; continue; }
      out.onlyProd.push(n);
    } else {
      if (why) { out.allowedCount++; continue; }
      out.onlyStaging.push(n);
    }
  }
  return out;
}

/* ══ self-test ═══════════════════════════════════════════════════════ */
if (process.argv.includes('--self-test')) {
  let pass = 0, fail = 0;
  const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); } };
  console.log('\n[check-parity --self-test] 数え方が正しいか');
  T('★改行コードだけの差は「差」と数えない', () => {
    const r = classify({ 'a.js': 'x\r\ny\r\n' }, { 'a.js': 'x\ny\n' }, {});
    if (r.differ.length) throw new Error('数えてしまっている');
    if (r.same !== 1) throw new Error('同じと数えていない');
  });
  T('★刻印(?v=)だけの差も「差」と数えない', () => {
    const r = classify({ 'a.html': '<script src="x.js?v=abc123">' }, { 'a.html': '<script src="x.js?v=999fff">' }, {});
    if (r.differ.length) throw new Error('数えてしまっている');
  });
  T('★中身が本当に違えば差として出す', () => {
    const r = classify({ 'a.js': 'const A=1;' }, { 'a.js': 'const A=2;' }, {});
    if (r.differ.length !== 1) throw new Error('見つけられていない');
  });
  T('★片側にしか無ければ、どちら側かを分けて出す', () => {
    const r = classify({ 'p.js': 'x' }, { 's.js': 'y' }, {});
    if (r.onlyProd[0] !== 'p.js' || r.onlyStaging[0] !== 's.js') throw new Error('分けられていない: ' + JSON.stringify(r));
  });
  T('理由を書いた物は差から外れる（ただし数は残る）', () => {
    const r = classify({ 'js/supa-config.js': 'A' }, { 'js/supa-config.js': 'B' }, { 'js/supa-config.js': '倉庫が違う' });
    if (r.differ.length) throw new Error('外れていない');
    if (r.allowedCount !== 1) throw new Error('数が残っていない');
  });
  T('★理由が無い差は必ず出る（一覧を逃げ道にできない）', () => {
    const r = classify({ 'x.js': 'A' }, { 'x.js': 'B' }, { 'y.js': '関係ない理由' });
    if (r.differ.length !== 1) throw new Error('逃げられてしまう');
  });
  console.log('\n' + pass + ' passed, ' + fail + ' failed');
  process.exit(fail ? 1 : 0);
}

/* ══ 本番（実物を見る） ═══════════════════════════════════════════════ */
const argPath = process.argv.slice(2).find(a => !a.startsWith('--'));
const OTHER = path.resolve(argPath || path.join(ROOT, '..', 'exally-staging'));
const JSON_OUT = process.argv.includes('--json');

if (!fs.existsSync(OTHER)) {
  console.error('★テスト側が見つかりません: ' + OTHER + '\n  場所を引数で渡してください: node scripts/check-parity.mjs <path>');
  process.exit(2);
}

function tracked(dir) {
  const out = execFileSync('git', ['ls-files'], { cwd: dir, encoding: 'utf8' });
  return out.split('\n').map(s => s.trim()).filter(Boolean);
}
function readAll(dir, files) {
  const m = {};
  for (const f of files) {
    const p = path.join(dir, f);
    try { m[f] = fs.readFileSync(p, 'utf8'); } catch { /* バイナリ/欠落は中身なしとして扱う */ m[f] = ' binary-or-missing'; }
  }
  return m;
}

const prodFiles = tracked(ROOT), stgFiles = tracked(OTHER);
const A = readAll(ROOT, prodFiles), B = readAll(OTHER, stgFiles);
const r = classify(A, B, ALLOWED);
const total = r.differ.length + r.onlyProd.length + r.onlyStaging.length;

if (JSON_OUT) {
  console.log(JSON.stringify({ total, ...r }, null, 1));
} else {
  console.log('\n[check-parity] 本番(exally) と テスト(exally-staging) の食い違い\n');
  console.log('  本番   : ' + ROOT);
  console.log('  テスト : ' + OTHER + '\n');
  const show = (title, arr) => {
    console.log('■ ' + title + '（' + arr.length + '件）');
    console.log(arr.length ? arr.map(x => '  ・' + x).join('\n') : '  （なし）');
    console.log('');
  };
  show('★中身が違う', r.differ);
  show('★本番にしか無い', r.onlyProd);
  show('★テストにしか無い', r.onlyStaging);
  console.log('■ 理由つきで許している差（' + r.allowedCount + '件）');
  console.log(Object.entries(ALLOWED).map(([k, v]) => '  ・' + k.padEnd(42) + v).join('\n'));
  console.log('\n── 実測 ──');
  console.log('  同じ ' + r.same + ' / ★差 ' + total + '★（中身違い ' + r.differ.length
    + ' / 本番のみ ' + r.onlyProd.length + ' / テストのみ ' + r.onlyStaging.length + '）');
  console.log('  ※ 改行コードと ?v= の刻印はそろえてから比べています（中身の違いではないため）');
  if (total > 10) console.log('\n  ★差が10件を超えています。新しい作業に入らないでください（指示役の決まり）。★');
}

process.exitCode = total > 10 ? 3 : 0;
