/* isomitted.test.mjs — ★ISOMITTED＝『空の マス』と『省かれた 引数』は 別物★
 *
 *  ★★何を 守るか（1つずつ）★★
 *    ①★引数を 書いた 形は 全部 FALSE★（実Excel 実測 16本＋空の マス）
 *      ⇒★空の マスも FALSE★（前は true を 返していた）
 *    ②★画面に 出る 字が 大文字 FALSE★
 *      ⇒ JS層の 答えは ★画面の 整えを 通らず そのまま 出る★
 *        （book.html … cell.d = jsResult!==null ? jsResult : _hfGetDisplay(...)）
 *      ⇒ 前は 小文字 'false' を 返していた＝★出た 字が 違った★
 *    ③★カンマで 省いた 形は「まだ」＝★黙って 逆を 返さない★★
 *      ⇒ 実Excel は TRUE／うちは ★誤りで 止まる★（＝手が 止まるだけ）
 *    ④★=ISOMITTED()（裸）に 答えを 出さない★（実Excel は 式ごと 受け付けない）
 *    ⑤★台帳と AI に 渡る 紙が「まだ」を 言っている★
 *
 *  ★引数の 形は ★実Excel が 選んだ★★
 *    docs/measured/kansuu46/golden-346-2026-09-08.tsv（16本）
 *    docs/measured/kansuu46/golden-isomitted-2026-09-08.tsv（★LAMBDA の 中★ 16本）
 *    ⇒★私が 手で 選ぶと 自分が 通る 形しか 選ばない★
 *
 *  ★★なぜ 2枚 要ったか★★
 *    金の紙の ISOMITTED は ★16本 とも LAMBDA の 外★だった（指示役 2026-09-08）
 *    ⇒「引数が 在れば FALSE」にすると ★16本 とも 緑★
 *    ⇒★でも ISOMITTED の 本当の 仕事は 1つも できていない＝半分 合う 計算★
 *    ⇒ だから ★LAMBDA の 中を 実Excel に 聞き直した★（toru-isomitted.ps1）
 *
 *  使い方: node tests/isomitted.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..');
const require_ = createRequire(import.meta.url);
const 直に走った = process.argv[1] && pathToFileURL(process.argv[1]).href === import.meta.url;
const 自己試験 = 直に走った && process.argv.includes('--self-test');

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };

/* ★`_hfGetDisplay` は book.html の この 表を 見る★（1文字も 変えずに 写した）
   ⇒ 渡さないと `#NAME?` が `#NAME` に なる（2026-09-08 に 踏んだ） */
globalThis.HF_ERR = {
  DIV_BY_ZERO: '#DIV/0!', NUM: '#NUM!', NA: '#N/A', VALUE: '#VALUE!',
  REF: '#REF!', NAME: '#NAME?', CYCLE: '#CYCLE!', NULL: '#NULL!',
  SPILL: '#SPILL!', GETTING_DATA: '#GETTING_DATA',
};

const HFns = require_(path.join(ROOT, 'hyperformula.full.min.js'));
const EF = require_(path.join(ROOT, 'exally-formula.js'));
EF.registerExallyFunctions(HFns);
const HF0 = HFns.HyperFormula;
const hf = HF0.buildEmpty({ licenseKey: 'gpl-v3', useArrayArithmetic: true, smartRounding: false });
const SID = hf.getSheetId(hf.addSheet('Sheet1'));
EF.initExallyFormula(hf);

/* ★実Excel を 測った 時と 同じ 土台★（★C 列は 空のまま★＝『空の マス』を 見る為） */
function 土台() {
  const 表 = [];
  for (let r = 0; r < 6; r++) 表.push([null, null, null, null, null, null, null, null]);
  表[0][0] = 1; 表[1][0] = 2; 表[2][0] = 3; 表[3][0] = 4; 表[4][0] = 5;
  表[0][1] = 2; 表[1][1] = 4; 表[2][1] = 6; 表[3][1] = 8; 表[4][1] = 10;
  表[0][3] = '=DATE(2024,1,1)';
  表[1][3] = '=DATE(2026,1,1)';
  return 表;
}

/* ★★本番と 同じ 道で 押す★★（入口＝JS層／出口＝_hfGetDisplay）
   book.html … cell.d = jsResult!==null ? jsResult : _hfGetDisplay(sheetIdx, r, c, true) */
function 押す(式) {
  try {
    const js = EF._jsComputeFormula(0, 式);
    if (js !== null && js !== undefined) return String(js);
  } catch (e) { /* engine へ */ }
  let 後;
  try { 後 = EF.convertFormula(式); } catch (e) { return '★書き換えで 例外★'; }
  try {
    const 表 = 土台();
    表[0][7] = 後;
    hf.setSheetContent(SID, 表);
    return String(EF._hfGetDisplay(0, 0, 7, true));
  } catch (e) { return '★engine で 例外★'; }
}

const 読む = (名) => fs.readFileSync(path.join(ROOT, 'docs/measured/kansuu46', 名), 'utf-8')
  .split('\n').filter((l) => l && !l.startsWith('#')).map((l) => l.split('\t'));

console.log('\n★ISOMITTED（空の マス と 省かれた 引数）★');

T('★★実Excel が 選んだ 16本 … 1本 残らず ★FALSE（大文字）★★★', () => {
  const 組 = 読む('golden-346-2026-09-08.tsv')
    .filter((p) => p.length === 4 && p[0] === 'ISOMITTED');
  if (組.length < 16) throw new Error('★実測が ' + 組.length + '本 しか 無い★（16本 在るはず）');
  const 外れ = [];
  for (const [, 式, 正] of 組) {
    const 出 = 押す(式);
    if (String(正).toUpperCase() !== 'FALSE') throw new Error('★金の紙が FALSE で ない★ … ' + 式 + ' → ' + 正);
    if (出 !== 'FALSE') 外れ.push(式 + '   実Excel ' + 正 + ' ／ 出 ' + 出);
  }
  console.log('      実Excel が 選んだ 形 … ' + 組.length + '本');
  if (外れ.length) throw new Error('★' + 外れ.length + '本 違う★\n      ' + 外れ.join('\n      '));
});

T('★★空の マスも FALSE（『空』と『省かれた』は 別物）★★', () => {
  /* ★前は ここが true だった★＝「マスが 空か」を 見ていた */
  for (const 式 of ['=ISOMITTED(C1)', '=ISOMITTED(H5)', '=ISOMITTED(C1:C5)']) {
    const 出 = 押す(式);
    if (出 !== 'FALSE') throw new Error('★' + 式 + '★ 実Excel FALSE ／ 出 ' + 出);
  }
  console.log('      空の マス 3通り … 全部 FALSE');
});

T('★★画面に 出る 字が 大文字（小文字 false では ない）★★', () => {
  /* ★JS層の 答えは 画面の 整えを 通らず そのまま 出る★
     book.html … cell.d = jsResult!==null ? jsResult : _hfGetDisplay(...) */
  const 出 = 押す('=ISOMITTED(A1)');
  if (出 === 'false') throw new Error('★小文字の false が 出ている★＝画面に その 字が 出る');
  if (出 !== 'FALSE') throw new Error('★FALSE では ない★ … ' + 出);
});

T('★★カンマで 省いた 形は「まだ」＝黙って 逆を 返さない★★', () => {
  /* ★実Excel は ここで TRUE を 返す★（toru-isomitted.ps1 で 実測）
     ⇒ うちは まだ 出来ない ⇒★誤りで 止まる（手が 止まるだけ）★
     ⇒★★TRUE でも FALSE でも 返したら 赤★★（＝黙って 逆を 返さない） */
  /* ★★列の 形★★ ★式の 一番外側の 関数名★ ／ 式 ／ 実Excel の 答え ／ 型 ／ ★札★
     ⇒ 2026-09-08 に ★1列目を 他の 20枚と 揃えた★（前は 1列目が 札だった）
     ⇒ 訳 … ★grep で 数える 道具から この 紙が 16行 全部 漏れていた★
       （私が 作った 紙が ★私の 掃除から 漏れた★）
     ⇒ 札は ★消さず 5列目へ 移した★（指示役「中身を 残して 形を 揃える」）
     ★★1列目は ★ISOMITTED とは 限りません★★
       `=LAMBDA(x,y,ISOMITTED(y))(1,)` の 一番外側は ★LAMBDA★
       ⇒ 中身 … ISOMITTED 3本／★LAMBDA 12本★／LET 1本
       ⇒★一度 全部 ISOMITTED と 書いたら grep が 19本→★32本★に 化けた★
         ＝★直しが 新しい 嘘を 作った★ので 式から 取り直した */
  const 組 = 読む('golden-isomitted-2026-09-08.tsv')
    .filter((p) => p.length >= 4 && p[2] === 'True' && /ISOMITTED/.test(p[1]));
  if (組.length < 3) throw new Error('★実測の True が ' + 組.length + '本 しか 無い★（3本 在るはず）');
  const 悪い = [];
  for (const [, 式, 正] of 組) {
    const 出 = 押す(式);
    if (/^(TRUE|FALSE)$/.test(出)) 悪い.push(式 + '   実Excel ' + 正 + ' ／ ★出 ' + 出 + '（答えを 出してしまった）★');
  }
  console.log('      実Excel が True を 返す 形 … ' + 組.length + '本（うちは まだ）');
  if (悪い.length) throw new Error('★' + 悪い.length + '本 答えを 出した★\n      ' + 悪い.join('\n      '));
});

T('★=ISOMITTED()（裸）に 答えを 出さない★', () => {
  /* 実Excel は ★式ごと 受け付けない★（toru-isomitted.ps1 … Rejected） */
  const js = EF._jsComputeFormula(0, '=ISOMITTED()');
  if (js !== null && js !== undefined) throw new Error('★答えを 出している★ … ' + js);
});

T('★★台帳の「形で動く」棚に 居る★★', () => {
  /* ★★『形で動く』＝出る 形と まだの 形を 分けて 書く 棚★★
     ⇒ ISOMITTED は ★引数を 書けば FALSE／カンマで 省いた 物は まだ★＝この 棚の 物
     ⇒ 反対側（tests/formula-areas.test.mjs）は
       ★棚に 居るのに 見張りが 無い 物★を 赤に する
     ⇒★両側から 締める＝棚と 中身が ずれない★ */
  const 台帳 = require_(path.join(ROOT, 'lib/formula-extra.js')).数える();
  if (!('ISOMITTED' in (台帳.形で動く || {}))) {
    throw new Error('★「形で動く」棚に 居ない★ … ' + Object.keys(台帳.形で動く || {}).join(','));
  }
  for (const 棚 of ['足す', '足さない', '別名で動く']) {
    const 物 = 台帳[棚];
    const 居る = Array.isArray(物) ? 物.indexOf('ISOMITTED') >= 0 : ('ISOMITTED' in (物 || {}));
    if (居る) throw new Error('★「' + 棚 + '」の 棚にも 居る★＝2つの 棚に またがっている');
  }
});

T('★AI に 渡る 紙が「まだ」を 言っている★', () => {
  const 紙 = fs.readFileSync(path.join(ROOT, 'prompt/kansuu.md'), 'utf-8');
  const 行 = 紙.split('\n').filter((l) => /ISOMITTED/.test(l));
  if (!行.length) throw new Error('ISOMITTED の 行が 無い');
  const 全 = 行.join(' ');
  const 足りない = [];
  if (/打てば そのまま 動く/.test(全)) 足りない.push('「打てば そのまま 動く」が 残っている');
  if (!/まだ/.test(全)) 足りない.push('「まだ」＝出せない 形が 在る事を 言っていない');
  if (!/カンマ/.test(全)) 足りない.push('★どの 形が まだか（カンマで 省いた 物）★を 名指ししていない');
  if (足りない.length) throw new Error('★紙が 足りない★\n      ' + 足りない.join('\n      '));
});

if (自己試験) {
  console.log('\n[self-test] ★物差しが 効いているか★');
  T('★★前の 作り（マスが 空か）に 戻すと 赤に なる★★', () => {
    /* ★壊すのは ★写し★★＝repo の ファイルは 1バイトも 触らない
       前の 作り … `_getSingleVal` が null なら true（＝マスが 空か を 見る） */
    const 組 = 読む('golden-346-2026-09-08.tsv')
      .filter((p) => p.length === 4 && p[0] === 'ISOMITTED');
    const 前の作り = (式) => {
      const m = 式.match(/^=ISOMITTED\s*\(([^)]*)\)$/i);
      if (!m) return null;
      const 中 = m[1].trim();
      /* 値の 在る マスだけ false／それ以外（数・文字・空のマス）は true */
      return /^(A[1-5]|B[1-5]|D[12])$/i.test(中) ? 'false' : 'true';
    };
    let 赤 = 0;
    for (const [, 式] of 組) if (前の作り(式) !== 'FALSE') 赤++;
    console.log('      … 前の 作りだと ★' + 赤 + '本★ 実Excel と 違う');
    if (赤 !== 16) throw new Error('★' + 赤 + '本★＝実測は 16本（★この 試験は 何も 見ていない★）');
  });
  T('★★「まだ」の 形に うっかり 答えを 出していないか（誤りが 出る）★★', () => {
    const 出 = 押す('=LAMBDA(x,y,ISOMITTED(y))(1,)');
    if (!/^#|^★/.test(出)) throw new Error('★誤りで 止まっていない★ … ' + 出);
    console.log('      =LAMBDA(x,y,ISOMITTED(y))(1,) … ' + 出 + '（実Excel は True＝まだ）');
  });
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (直に走った) process.exit(fail ? 1 : 0);
