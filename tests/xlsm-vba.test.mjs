/* xlsm-vba.test.mjs — ★VBA入り(.xlsm)を どう扱うか★（2026-08-25 に決めた）
 *
 *  ★決めた事★
 *    ①★開ける★（「読むだけ」ではない＝直して書き出せる）
 *    ②★VBAには 触らない。そのまま残す★（★実測：書き出しても 1バイトも変わらない★）
 *    ③★VBAは 動かさない★（うちはブラウザ。Windows＋マクロ有効化が要る物は 勧めない）
 *    ④★VBAが要る仕事は うちの側（手順を覚えさせる＝レシピ）で済ませる★
 *  ★「まだ出来ません」ではなく「どうなるか」を 先に言う★（画面にも出す）
 *
 *  使い方: node tests/xlsm-vba.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const SELF = process.argv.includes('--self-test');
const OVERRIDE = process.env.EXALLY_VBA_OVERRIDE ? JSON.parse(process.env.EXALLY_VBA_OVERRIDE) : {};
const 読む = (rel) => fs.readFileSync(OVERRIDE[rel] || path.join(ROOT, rel), 'utf8');
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ok   ' + n); } catch (e) { fail++; console.log('  NG   ' + n + '\n       ' + (e && e.message)); } };
const AT = async (n, fn) => { try { await fn(); pass++; console.log('  ok   ' + n); } catch (e) { fail++; console.log('  NG   ' + n + '\n       ' + (e && e.message)); } };
const ok = (c, m) => { if (!c) throw new Error(m || 'expected truthy'); };
const eq = (a, b, m) => { if (a !== b) throw new Error((m ? m + ': ' : '') + '期待=' + JSON.stringify(b) + ' 実際=' + JSON.stringify(a)); };

console.log('');
console.log('[xlsm-vba] ★VBA入り(.xlsm)の扱い★');

const bookOpen = 読む('js/book-open.js');
const book = 読む('book.html');

T('★VBAが入っている事に 気づく（xl/vbaProject.bin を見る）★', () => {
  ok(bookOpen.indexOf('vbaProject.bin') > 0, '見ていない');
  ok(/hasVba = true/.test(bookOpen), '★気づいても 覚えていない★');
  ok(/hasVba: !!hasVba/.test(bookOpen), '★呼ぶ側へ 渡していない★');
});
T('★客に「どうなるか」を 先に言う（画面に出す言葉が在る）★', () => {
  const i = bookOpen.indexOf('MSG_VBA');
  ok(i > 0, '言葉が無い');
  const 言葉 = bookOpen.slice(i, i + 400);
  ok(言葉.indexOf('そのまま残します') > 0, '★残る事を 言っていない★');
  ok(言葉.indexOf('動きません') > 0, '★動かない事を 言っていない★');
  ok(言葉.indexOf('代わりに') > 0 || 言葉.indexOf('出来ます') > 0, '★次の手を 言っていない★');
  for (const だめ of ['まだ出来ません', '未対応', '対応していません']) {
    ok(言葉.indexOf(だめ) < 0, '★「' + だめ + '」と言っている（どうなるかを言う）★');
  }
});
T('★画面が その言葉を 実際に出す（持っているだけにしない）★', () => {
  ok(/res\.hasVba/.test(book), '★画面が 見ていない★');
  ok(/BookOpen\.MSG_VBA/.test(book), '★画面が 言葉を出していない★');
});
T('★客に見せる字に ★ を書かない★', () => {
  /* ★注記(コメント)には ★ が在ってよい。客に出る字だけを見る★（検査の側の間違いを直した） */
  const i = bookOpen.indexOf('var MSG_VBA');
  const 言葉 = bookOpen.slice(i, bookOpen.indexOf(';', i + 200)).replace(/\/\*[\s\S]*?\*\//g, '');
  const 中身 = (言葉.match(/'[^']*'/g) || []).join('');
  ok(中身.length > 20, '言葉が取れていない');
  ok(中身.indexOf('★') < 0, '★客の字に ★ が出ている★：' + 中身.slice(0, 60));
});
T('★VBAを 動かす物を 1つも持っていない★', () => {
  for (const だめ of ['vbaProject.execute', 'runMacro', 'Application.Run']) {
    ok(bookOpen.indexOf(だめ) < 0, '★' + だめ + ' が在る★');
  }
});

/* ── ★実物で測る：書き出しても VBAが1バイトも変わらない★ ── */
await AT('★VBAの部品は 書き出しても 1バイトも変わらない（実際に往復させる）★', async () => {
  /* ★見本＝tests/fixtures/vba-sample.xlsm（実Excelで作った .xlsm に VBAの部品を1つ足した物）★
     ★中身は作り物★＝見るのは「1バイトも変わらないか」だけ。 */
  const 見本 = path.join(ROOT, 'tests/fixtures/vba-sample.xlsm');
  if (!fs.existsSync(見本)) { console.log('       ★未測定★ 見本が無い'); return; }
  const ZipSurgeon = require_(path.join(ROOT, 'lib/zip-surgeon.js'));
  const XlsxEdit = require_(path.join(ROOT, 'lib/xlsx-edit.js'));
  const 元 = new Uint8Array(fs.readFileSync(見本));
  const z1 = ZipSurgeon.read(元);
  ok(z1.has('xl/vbaProject.bin'), '見本に VBAの部品が無い');
  const 前 = Buffer.from(await z1.bytes('xl/vbaProject.bin'));

  const bk = await XlsxEdit.open(元);
  await XlsxEdit.setValues(bk, (bk.sheets && bk.sheets[0].name) || '計算', { A1: { v: 11, t: 'n' } });
  const out = await XlsxEdit.save(bk);
  const 出 = out.bytes || out;
  const z2 = ZipSurgeon.read(new Uint8Array(出));
  ok(z2.has('xl/vbaProject.bin'), '★書き出したら VBAの部品が 消えた★');
  const 後 = Buffer.from(await z2.bytes('xl/vbaProject.bin'));
  eq(後.equals(前), true, '★VBAの中身が 変わった★');
  console.log('       … 元 ' + 元.length + 'バイト → 書き出し ' + 出.length + 'バイト／VBA ' + 前.length + 'バイトは 不変');
});

console.log('');
console.log('  ' + pass + ' 緑 / ' + fail + ' 赤');

if (SELF) {
  const { spawnSync } = await import('node:child_process');
  const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'exally-vbaself-'));
  console.log('');
  console.log('[self-test] わざと壊して 赤くなるかを数える（★repo は読むだけ★）');
  const BREAKS = [
    ['js/book-open.js', '★VBAに気づかない★', (s) => s.replace('{ kind = \'xlsm\'; hasVba = true; }', "kind = 'xlsm';")],
    ['js/book-open.js', '★気づいても 呼ぶ側へ渡さない★', (s) => s.replace('hasVba: !!hasVba', 'hasVba: false')],
    ['js/book-open.js', '★「まだ出来ません」と言う★', (s) => s.replace('マクロは そのまま残しますが、ここでは 動きません。', 'マクロは まだ出来ません。')],
    ['js/book-open.js', '★次の手を 言わない★', (s) => s.replace('毎月の繰り返しは、このあと「手順を覚えさせる」で 代わりに出来ます。', '')],
    ['book.html', '★画面が 出さない★', (s) => s.replace('if (res.hasVba && typeof BookOpen', 'if (false && typeof BookOpen')],
  ];
  let red = 0;
  for (const [rel, name, brk] of BREAKS) {
    const orig = fs.readFileSync(path.join(ROOT, rel), 'utf8');
    const bad = brk(orig);
    if (bad === orig) { console.log('  ★置換できず★  ' + name); continue; }
    const f = path.join(TMP, rel.replace(/[\\/]/g, '_'));
    fs.writeFileSync(f, bad, 'utf8');
    const env = Object.assign({}, process.env, { EXALLY_VBA_OVERRIDE: JSON.stringify({ [rel]: f }) });
    const r = spawnSync(process.execPath, ['--max-old-space-size=4096', path.join(__dirname, 'xlsm-vba.test.mjs')], { encoding: 'utf8', env });
    if (r.status !== 0) { red++; console.log('  赤くなった  ' + name); }
    else console.log('  ★素通り★  ' + name);
  }
  try { fs.rmSync(TMP, { recursive: true, force: true }); } catch (e) { /* 消せなくても検査は済んでいる */ }
  console.log('');
  console.log('  ' + red + '/' + BREAKS.length + ' 通りで赤くなった');
  process.exit(red === BREAKS.length ? 0 : 1);
}

process.exit(fail ? 1 : 0);
