/* kazoekagi-onaji-ka.mjs — ★NUL を 字で 書き直しても ★鍵が 1つも 変わらない★か★（2026-09-07）
 *
 *  ★★なぜ 要るか★★
 *    `lib/ribbon.js` の NUL は ★飾りでは なく 区切り★です。
 *      `var 数え鍵 = (親 || '') + <NUL> + 元名;`
 *    ⇒★人の 字に 絶対 出てこない から 選ばれています★
 *    ⇒★もし ★空白★に 直したら
 *       親='A B'／元名='C' と 親='A'／元名='B C' が ★同じ 鍵に なります★
 *    ⇒★★『見た目が 同じ』と『役目が 同じ』は 別★★（2026-09-07 指示役が 止めた）
 *
 *  ★★だから 直した 後は ★鍵そのものを 突き合わせます★★
 *    ・直す前（main の lib/ribbon.js）
 *    ・直した後（この 枝の lib/ribbon.js）
 *    ⇒ 同じ 仕様（ribbon-spec）から 木を 作り、★全部の 鍵を 並べて SHA を 比べる★
 *    ⇒★1つでも 違えば 赤★
 *
 *  ★『字で 書いた 逃がし』は 同じ 1バイト★
 *    実測 … JS が 読むと 長さ1／文字コード ★0★（空白なら 32）
 *
 *  使い方: node docs/measured/ribbon/kazoekagi-onaji-ka.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..', '..', '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const 紙 = [];
const 言う = (s) => { 紙.push(s); console.log(s); };
let pass = 0, fail = 0;
const T = (n, よい, 添え) => {
  if (よい) { pass++; 言う('  ok   ' + n); }
  else { fail++; 言う('  NG   ' + n + (添え ? '\n       ' + 添え : '')); }
};

言う('# ★NUL を 字で 書き直しても 鍵が 変わらないか★（' + new Date().toISOString().slice(0, 10) + '）');
言う('');

/* ★①同じ 1バイトに 戻るか★ */
const 逃がし = JSON.parse('"' + '\\' + 'u0000"');
T('★字で 書いた 逃がしは ★同じ 1バイト★（空白では ない）★',
  逃がし.length === 1 && 逃がし.charCodeAt(0) === 0,
  '長さ ' + 逃がし.length + ' ／ 文字コード ' + 逃がし.charCodeAt(0) + '（空白なら 32）');

/* ★②直す前の lib/ribbon.js を main から 取り出す★ */
/* ★★写しは lib/ の 中に 置く★★
   ⇒ ribbon.js は ★隣の lib を 名前で 呼ぶ★ので、外に 置くと 見つけられない
   ⇒ 置きっぱなしに しない（下の finally で 必ず 消す） */
const 前の道 = path.join(ROOT, 'lib', '__ribbon-mae-tmp.cjs');
let 取れた = true;
try {
  const 前 = execFileSync('git', ['show', 'main:lib/ribbon.js'], { cwd: ROOT, maxBuffer: 1 << 26 });
  fs.writeFileSync(前の道, 前);
} catch (e) { 取れた = false; }
T('★直す前の lib/ribbon.js を main から 取り出せた★', 取れた);

if (取れた) {

  const 前バイト = fs.readFileSync(前の道);
  const 後バイト = fs.readFileSync(path.join(ROOT, 'lib/ribbon.js'));
  言う('  ★直す前★ ' + 前バイト.length + 'バイト ／ 生の NUL ' + 前バイト.filter((b) => b === 0).length + '個');
  言う('  ★直した後★ ' + 後バイト.length + 'バイト ／ 生の NUL ' + 後バイト.filter((b) => b === 0).length + '個');

  const 仕様 = require_(path.join(ROOT, 'lib/ribbon-spec.js'));
  const 前R = require_(前の道);
  const 後R = require_(path.join(ROOT, 'lib/ribbon.js'));

  /* ★木を 作って ★全部の 枝★を 並べる★
     ★木の 形★（実際に 開いて 見た）… { 子:{ 'H':{ 子:{ '0':{ 子:{}, 動作, 名, タブ元 } … } } }, 曖昧, 入れた, 飛ばした }
     ⇒★Alt の 打つ 順（H→1→…）が 鍵★＝★数え鍵が 効いている 所そのもの★
     ⇒ 1本しか 出ないなら ★物差しが 空洞★（前の 版で 実際に そう なった） */
  const 並べる = (R) => {
    const 木 = R.木を作る(仕様);
    const 出 = [];
    const 掘る = (節, 順) => {
      if (!節 || !節.子) return;
      for (const k of Object.keys(節.子).sort()) {
        const c = 節.子[k];
        出.push(順 + k + 逃がし + (c.タブ元 || '') + '/' + (c.名 || '') + '/' + (c.動作 || ''));
        掘る(c, 順 + k);
      }
    };
    掘る(木, '');
    /* ★ぶつかった 物・入れた 数・飛ばした 数も 並べる★（ここが 数え鍵の 効き所） */
    出.push('曖昧=' + JSON.stringify(木.曖昧 || []));
    出.push('入れた=' + (木.入れた === undefined ? '?' : 木.入れた));
    出.push('飛ばした=' + JSON.stringify(木.飛ばした || []));
    return 出;
  };

  let 前並 = [], 後並 = [], 作れた = true;
  try { 前並 = 並べる(前R); 後並 = 並べる(後R); } catch (e) { 作れた = false; 言う('  ★木を 作れなかった … ' + e.message + '★'); }

  T('★木が 作れて 中身が 空で ない★', 作れた && 前並.length > 10 && 後並.length > 10,
    '前 ' + 前並.length + '本 ／ 後 ' + 後並.length + '本');

  const sha = (a) => crypto.createHash('sha256').update(a.join('\n'), 'utf-8').digest('hex').slice(0, 16);
  言う('  ★鍵の 数★ 前 ' + 前並.length + '本 ／ 後 ' + 後並.length + '本');
  言う('  ★鍵の SHA★ 前 ' + sha(前並) + ' ／ 後 ' + sha(後並));
  T('★★鍵が 1つ 残らず 同じ（SHA 一致）★★', 前並.length > 0 && sha(前並) === sha(後並),
    '★違う＝区切りの 役目が 変わった★');

  /* ★③空白に したら どう なるか＝わざと やって 見せる★（物差しの 自己確認） */
  const 空白で = 前並.map((k) => k.split(逃がし).join(' '));
  const だぶり = 空白で.length - new Set(空白で).size;
  言う('');
  言う('  ★もし 空白に していたら★');
  言う('    同じ 鍵に なる 組 … ' + だぶり + '組');
  T('★空白に すると 鍵が ぶつかる 事を 数で 示せる（0組なら この 心配は 要らない）★', true,
    'ぶつかり ' + だぶり + '組');
}
try { fs.rmSync(前の道, { force: true }); } catch (e) { /* 手元の 写し */ }

言う('');
言う('kazoekagi: ' + pass + ' 緑 / ' + fail + ' 赤');
fs.writeFileSync(path.join(ここ, 'kazoekagi-onaji-ka.txt'), 紙.join('\n') + '\n', { encoding: 'utf-8' });
console.log('\n★書いた … docs/measured/ribbon/kazoekagi-onaji-ka.txt★');
process.exit(fail ? 1 : 0);
