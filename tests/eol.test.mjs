/* eol.test.mjs — ★行の 終わりを CRLF に して 試験を 黙って 壊すな★（2026-09-05）
 *
 *  ★何が 起きたか（実測）★
 *    私が python で ファイルを 読んで 書き戻した時、★CRLF に なった★。
 *    repo の 決まりは `.gitattributes` で ★eol: lf★。
 *    ⇒ わざと壊す 試験は ★複数行の 字を そのまま 掴む★ので、
 *      改行が 1文字 違うだけで ★「置換できず」＝壊せない＝赤★に なる。
 *    ⇒★1回の 書き戻しで 17本が 落ちた★。
 *    ⇒ しかも ★中身は 1文字も 間違っていない★ので、
 *      赤の 中身を 読んでも ★理由が 分からない★（私は 30分 探した）。
 *
 *  ★なぜ 機械で 見張るか★
 *    ・git は ★commit する時に 直してくれる★＝★手元では 壊れたまま★
 *    ・つまり ★ci-same（手元）だけが 赤に なり、CI（押した物）は 緑★
 *      ＝★「手元は赤・CIは緑」という 一番 分かりにくい 形★
 *
 *  使い方: node tests/eol.test.mjs
 *          node tests/eol.test.mjs --self-test
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };

/** ★純関数★＝中身を 見て CRLF が 混ざっているか（self-test で 作り物を 通せる） */
export function CRLFが在るか(src) {
  const 全 = (src.match(/\r\n/g) || []).length;
  return { 有る: 全 > 0, 数: 全 };
}

if (process.argv.includes('--self-test')) {
  console.log('\n[eol --self-test] わざと壊して赤になるか');
  T('★CRLF が 1つでも 在れば 見つける', () => {
    if (!CRLFが在るか('a\r\nb').有る) throw new Error('見つけていない');
  });
  T('★LF だけなら 通す', () => {
    if (CRLFが在るか('a\nb\nc').有る) throw new Error('誤検知');
  });
  T('★数も 数える', () => {
    if (CRLFが在るか('a\r\nb\r\n').数 !== 2) throw new Error('数が ちがう');
  });
  T('★空でも 落ちない', () => { CRLFが在るか(''); });
  console.log('\n' + pass + ' passed, ' + fail + ' failed');
  process.exit(fail ? 1 : 0);
}

/* ══ 本番 ═══════════════════════════════════════════════════ */
const SKIP = new Set(['node_modules', '.git', 'tmp', '.vercel', 'dist']);
function 集める(rel, out = []) {
  for (const name of fs.readdirSync(path.join(ROOT, rel || '.'))) {
    if (SKIP.has(name)) continue;
    const r = rel ? rel + '/' + name : name;
    if (fs.statSync(path.join(ROOT, r)).isDirectory()) 集める(r, out);
    else if (/\.(js|mjs|cjs|html|json|md|yml|css)$/i.test(name)) out.push(r);
  }
  return out;
}
const 見る = 集める('');
const 悪い = [];
for (const r of 見る) {
  const c = CRLFが在るか(fs.readFileSync(path.join(ROOT, r), 'utf8'));
  if (c.有る) 悪い.push({ f: r, n: c.数 });
}

console.log('\n[eol] 行の終わりが LF に そろっているか（.gitattributes の 決まり）');

T('★★CRLF の ファイルが 1本も 無い★★', () => {
  if (悪い.length) {
    throw new Error('CRLF が 混ざっています:\n'
      + 悪い.slice(0, 20).map((x) => '   - ' + x.f + '（' + x.n + '行）').join('\n')
      + (悪い.length > 20 ? '\n   … ほか ' + (悪い.length - 20) + '本' : '')
      + '\n   → ★書き戻す道具が CRLF にしています★。LF で 書き直してください。'
      + '\n     ★中身が 合っていても、わざと壊す試験が 掴めず 黙って 赤に なります★'
      + '\n     （2026-09-05 に これで 17本 落ちた）');
  }
});

T('検査が空振りしていない（実際にファイルを読んでいる）', () => {
  if (見る.length < 50) throw new Error('見たファイルが 少なすぎます: ' + 見る.length);
});

console.log('\n── 実測 ──');
console.log('  見た ファイル … ' + 見る.length + '本 ／ CRLF … ' + 悪い.length + '本');
console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
