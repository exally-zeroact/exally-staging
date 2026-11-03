/* formula-complex.test.mjs — ★IM系（複素数）の 出す 字が 実Excel と 同じか★
 *
 *  ★★何を 守るか（1つずつ）★★
 *    ①★実Excel が 選んだ IM系 226本の うち ★197本★が 同じ 字★
 *      ⇒★合わない 29本は 全部 ★別の 山★と 名指しできる★
 *        （文字を 数として 読めていない 22／大きい 数 2／計算の 精度 5）
 *      ⇒★「226本 全部 合う」とは ★書かない★★
 *        ＝この 直しの 担当では ない 物まで 背負うと ★半分 合う 計算★に なる
 *      ⇒★名前の 分からない 落ち方が 1本でも 出たら 赤★
 *    ②★数を 返す 4個（IMABS IMAGINARY IMARGUMENT IMREAL）を ★触っていない★★
 *      ⇒ そこに 15桁を かけると ★壊れる★（実Excel の =IMARGUMENT(-2) は ★17桁★）
 *    ③★包んだのは 21個 ちょうど★（★半分 包んで 黙って 進まない★）
 *      ⇒★名簿が 元の ★26個★ を 1つ 残らず 名指ししている★
 *        （21 包む ＋ 4 触らない ＋ 1 まだ 測っていない）
 *      ⇒★2026-09-08 に 指示役が 見つけた★ … 名簿は 25個しか 無く
 *        ★COMPLEX が 名簿に 無いまま 黙って 素通り★していた
 *    ④★計算を 書き直していない★＝HyperFormula の 元の 実装を 継いでいる
 *
 *  ★★『切る』では ない・『丸める』★★（指示役 2026-09-08）
 *    元 0.5403023058681398 ／ ★丸め 0.54030230586814（実Excel と 同じ）★／ 切り 0.540302305868139
 *    ★両方 作って 走らせて 測った★ … ★丸め 78本 直る ／ 切り 41本 直る★
 *    ⇒★数が 動いた＝物差しは 桁を 見ている★
 *
 *  ★引数の 形は ★実Excel が 選んだ★★（docs/measured/kansuu46/golden-346-2026-09-08.tsv）
 *    ⇒★私が 手で 選ぶと 自分が 通る 形しか 選ばない★
 *
 *  使い方: node tests/formula-complex.test.mjs [--self-test]
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

const F = require_(path.join(ROOT, 'lib/formula-complex.js'));
const P = require_(path.join(ROOT, 'lib/formula-complex-plug.js'));
const HFns = require_(path.join(ROOT, 'hyperformula.full.min.js'));
const EF = require_(path.join(ROOT, 'exally-formula.js'));
EF.registerExallyFunctions(HFns);
const HF0 = HFns.HyperFormula;

/* ★元の 物が 何個 在るかを 先に 数える★（★思い込みで 名前を 並べない★） */
const 元 = HF0.getFunctionPlugin('IMCOS');
const 元の名たち = Object.keys((元 && 元.implementedFunctions) || {});

const 包んだ = P.つなぐ(HF0, F);

const hf = HF0.buildEmpty({ licenseKey: 'gpl-v3', useArrayArithmetic: true, smartRounding: false });
const SID = hf.getSheetId(hf.addSheet('Sheet1'));
EF.initExallyFormula(hf);

const 赤の名 = (t) => ({
  NA: '#N/A', DIV_BY_ZERO: '#DIV/0!', VALUE: '#VALUE!', NUM: '#NUM!',
  NAME: '#NAME?', REF: '#REF!', CYCLE: '#CYCLE!', ERROR: '#ERROR!', SPILL: '#SPILL!',
}[t] || ('#' + t));

/* ★実Excel を 測った 時と 同じ 材料★（golden-346 を 取った 時の 土台） */
function 土台() {
  const 表 = [];
  for (let r = 0; r < 6; r++) 表.push([null, null, null, null, null, null, null, null]);
  表[0][0] = 1; 表[1][0] = 2; 表[2][0] = 3; 表[3][0] = 4; 表[4][0] = 5;
  表[0][1] = 2; 表[1][1] = 4; 表[2][1] = 6; 表[3][1] = 8; 表[4][1] = 10;
  表[0][3] = '=DATE(2024,1,1)';
  表[1][3] = '=DATE(2026,1,1)';
  return 表;
}
function 押す(式) {
  const 表 = 土台();
  表[0][7] = EF.convertFormula(式);
  hf.setSheetContent(SID, 表);
  const v = hf.getCellValue({ sheet: SID, row: 0, col: 7 });
  return (v && v.type) ? 赤の名(v.type) : v;
}

const 金道 = path.join(ROOT, 'docs/measured/kansuu46/golden-346-2026-09-08.tsv');
const 金 = fs.readFileSync(金道, 'utf-8').split('\n')
  .filter((l) => l && !l.startsWith('#'))
  .map((l) => l.split('\t')).filter((p) => p.length === 4 && /^IM/.test(p[0]))
  .map((p) => ({ 名: p[0], 式: p[1], 答: p[2], 型: p[3] }));

const 同じか = (出, 正, 型) => {
  const s = String(出);
  if (型 === 'Double') {
    const a = Number(出), b = Number(正);
    if (!isFinite(a) || !isFinite(b)) return s === 正;
    if (a === b) return true;
    const 大 = Math.max(Math.abs(a), Math.abs(b));
    return Math.abs(a - b) <= (大 > 1 ? 大 * 1e-9 : 1e-9);
  }
  return s === 正;
};

console.log('\n★IM系（複素数）の 字★');

T('★★包んだのは 21個 ちょうど（半分 包んで 進まない）★★', () => {
  if (包んだ !== 21) throw new Error('★' + 包んだ + '個 しか 包めていない★（21個 のはず）');
  if (F.文字を返す.length !== 21) throw new Error('名簿が ' + F.文字を返す.length + '個');
  if (F.数を返す.length !== 4) throw new Error('触らない 物が ' + F.数を返す.length + '個');
});

T('★★名簿が 元の 26個を ★1つ 残らず★ 名指ししている★★', () => {
  /* ★★2026-09-08 に 指示役が 見つけた★★
     私は「26個 全部（21包む／★5そのまま★）」と 書きながら 名簿は ★21+4＝25個★ だった。
     ⇒★COMPLEX が 名簿に 無いまま ★黙って 素通り★していた★
     ⇒★「70個と 書いて 60個しか 並んでいない」と 同じ 形★
     ⇒ だから ★数では なく 名簿を 突き合わせる★ */
  const 名簿 = F.文字を返す.concat(F.数を返す, F.まだ測っていない);
  const もれ = 元の名たち.filter((n) => 名簿.indexOf(n) < 0);
  const よけい = 名簿.filter((n) => 元の名たち.indexOf(n) < 0);
  console.log('      元 ' + 元の名たち.length + '個 ＝ 包む ' + F.文字を返す.length
    + ' ＋ 触らない ' + F.数を返す.length + ' ＋ まだ測っていない ' + F.まだ測っていない.length
    + '（' + F.まだ測っていない.join(' ') + '）');
  if (もれ.length) throw new Error('★名簿に 無い（＝黙って 素通りする）★ … ' + もれ.join(' '));
  if (よけい.length) throw new Error('★名簿に 在るが 元に 無い★ … ' + よけい.join(' '));
  if (名簿.length !== 元の名たち.length) {
    throw new Error('★名簿 ' + 名簿.length + '個 ／ 元 ' + 元の名たち.length + '個★');
  }
});

T('★★COMPLEX（まだ 測っていない）は 触っていない★★', () => {
  /* ★実Excel で 6本 測って 6本とも 合っている★
     ★「実Excel が COMPLEX を 15桁に 丸めるか」は まだ 測っていない★
     ⇒★測っていない 物を 直さない＝今 合っている 物を 当て推量で 触らない★ */
  const 組 = fs.readFileSync(金道, 'utf-8').split('\n')
    .map((l) => l.split('\t')).filter((p) => p.length === 4 && p[0] === 'COMPLEX');
  if (組.length < 6) throw new Error('★COMPLEX の 実測が ' + 組.length + '本 しか 無い★（6本 在るはず）');
  const 外れ = [];
  for (const [, 式, 正] of 組) {
    const 出 = String(押す(式));
    if (出 !== 正) 外れ.push(式 + '  実Excel ' + 正 + ' ／ 出 ' + 出);
  }
  console.log('      COMPLEX … 実測 ' + 組.length + '本 全部 そのまま');
  if (外れ.length) throw new Error('★' + 外れ.length + '本 変わった★\n      ' + 外れ.join('\n      '));
});

T('★元が 持っている 物を 1つも 落としていない★', () => {
  /* ★HF は 自前の prototype を 見る＝1つでも 置き忘れると 積み直しが 落ちる★
     （実物で 踏んだ … "Function method complex not found in plugin …"） */
  const 新 = HF0.getFunctionPlugin('IMCOS');
  const 新の名たち = Object.keys((新 && 新.implementedFunctions) || {});
  if (新の名たち.length !== 元の名たち.length) {
    throw new Error('★元 ' + 元の名たち.length + '個 → 今 ' + 新の名たち.length + '個★');
  }
  const 足りない = 元の名たち.filter((n) => 新の名たち.indexOf(n) < 0);
  if (足りない.length) throw new Error('★落とした★ … ' + 足りない.join(' '));
  console.log('      元の 名前 ' + 元の名たち.length + '個 … 1つも 落としていない');
});

/* ★★残る 29本は ★この 直しの 担当では ない★★
   ⇒★「226本 全部 合う」と 書くと ★他の 山の 分まで 背負って 赤に なる★
   ⇒ だから ★残る 物の 形を 名指しで 書き★、
     ★名前の 分からない 落ち方が 1本でも 出たら 赤★に する
   ★実測の 内訳（docs/measured/kansuu46/mae-ato-im-ato-2026-09-08.tsv）★
     ・22本 … 引数が ★"12:30"（文字）★ ⇒ ★文字を 数として 読めていない★（別の 山）
     ・ 2本 … IMCSCH(D1) IMSECH(D1) ⇒ ★大きい 数で 音を 上げる★（別の 山）
     ・ 5本 … IMSIN(D1) IMSQRT×4 ⇒ ★計算の 精度★（15桁に 丸めても 最後の 1桁が 違う） */
/* ★★分けるのは ★式の 字★では なく ★出た 物★★
   ⇒ 最初 式の 字（"…" が 在るか／D1 が 在るか）で 分けたら
     `=IMSIN(D1)` が ★2つの 山に 当てはまって★ 数が ずれた
   ⇒★誤りを 返したのか・値を 返したのか★が 本当の 区別 */
const 別の山 = [
  { 名: '文字を 数として 読めていない', 見分け: (g, 出) => /^#/.test(出) && /"/.test(g.式), 実測: 22 },
  { 名: '大きい 数で 音を 上げる', 見分け: (g, 出) => /^#/.test(出), 実測: 2 },
  { 名: '計算の 精度（値は 出ている）', 見分け: (g, 出) => !/^#/.test(出), 実測: 5 },
];

T('★★実Excel が 選んだ IM系 ' + 金.length + '本 … 合わない 物は 全部 ★別の 山★と 名指しできる★★', () => {
  if (金.length < 200) throw new Error('★実測が ' + 金.length + '本 しか 無い★（226本 在るはず）');
  let 合 = 0;
  const 山ごと = {};
  const 名前が分からない = [];
  for (const g of 金) {
    const 出 = 押す(g.式);
    if (同じか(出, g.答, g.型)) { 合++; continue; }
    const 山 = 別の山.find((y) => y.見分け(g, String(出)));
    if (!山) { 名前が分からない.push(g.式 + '  実Excel ' + g.答 + ' ／ 出 ' + 出); continue; }
    山ごと[山.名] = (山ごと[山.名] || 0) + 1;
  }
  console.log('      合った ' + 合 + '本 ／ 合わない ' + (金.length - 合) + '本');
  for (const y of 別の山) console.log('        ' + y.名 + ' … ' + (山ごと[y.名] || 0) + '本（実測 ' + y.実測 + '本）');
  if (名前が分からない.length) {
    throw new Error('★名前の 分からない 落ち方が ' + 名前が分からない.length + '本★\n      '
      + 名前が分からない.slice(0, 10).join('\n      '));
  }
  /* ★実測より 減っていない事★＝直したのに 合う 数が 減ったら 赤 */
  if (合 < 197) throw new Error('★合った数が ' + 合 + '本★＝実測 197本より 減っている');
  /* ★別の山が 膨らんでいない事★＝この 直しで 他所を 壊したら 赤 */
  for (const y of 別の山) {
    if ((山ごと[y.名] || 0) > y.実測) {
      throw new Error('★' + y.名 + ' が ' + 山ごと[y.名] + '本★＝実測 ' + y.実測 + '本より 増えている');
    }
  }
});

T('★★数を 返す 4個は 触っていない（15桁を かけたら 壊れる）★★', () => {
  /* ★★実Excel の =IMARGUMENT(-2) は ★3.1415926535897931（17桁）★★
     ⇒ 15桁に 丸めると 3.14159265358979 に なって ★壊れる★
     ⇒ だから ★押して 確かめる★（名簿に 書いてあるだけでは 足りない） */
  const 見本 = [
    ['=IMARGUMENT(-2)', Math.PI],
    ['=IMABS(-2)', 2],
    ['=IMREAL(2)', 2],
    ['=IMAGINARY(2)', 0],
  ];
  for (const [式, 正] of 見本) {
    const 出 = 押す(式);
    if (typeof 出 !== 'number') throw new Error('★' + 式 + ' が 数で 返っていない★ … ' + 出 + '（' + typeof 出 + '）');
    if (Math.abs(出 - 正) > 1e-12) throw new Error('★' + 式 + '★ 正 ' + 正 + ' ／ 出 ' + 出);
  }
  /* ★桁を 落としていない事★ … 丸めたら 15桁に なる 所を 17桁 持っているか */
  const p = 押す('=IMARGUMENT(-2)');
  if (String(p) === String(Number(Number(p).toPrecision(15)))) {
    /* Math.PI は 15桁に 丸めても 字が 変わらない 事が 在るので 数で 見る */
  }
  if (String(p).length < 16) throw new Error('★桁が 落ちている★ … ' + p);
});

T('★丸めている（切っていない）★', () => {
  /* ★言葉と 中身が 合っているか★＝『切る』に 直されたら ここが 赤に なる */
  if (F.十五桁('0.5403023058681398') !== '0.54030230586814') {
    throw new Error('★丸めになっていない★ … ' + F.十五桁('0.5403023058681398'));
  }
  if (F.十五桁('0.5403023058681398') === '0.540302305868139') {
    throw new Error('★切ってしまっている★');
  }
});

T('★指数は 大文字 E（実Excel の 書き方）★', () => {
  const r = F.十五桁('8.659560562354932e-17');
  if (!/E/.test(r) || /e/.test(r)) throw new Error('★小文字の e が 残っている★ … ' + r);
});

if (自己試験) {
  console.log('\n[self-test] ★物差しが 効いているか★');
  T('★★直しを 入れる 前の 実測（生の 字）に かけると 78本 直る★★', () => {
    /* ★★「丸めた 後の 字」から 元に 戻す 事は できない★★
       ⇒ だから ★直す 前に 実際に 測った 紙★を 読んで かけ直す
       ⇒ 紙 … `mae-ato-im-mae-2026-09-08.tsv`（origin/main 5ee2236 を 別の 置き場で 押した 物）
       ⇒★これが 78本 に ならなければ ★直しは 何も していない★★ */
    const 道 = path.join(ROOT, 'docs/measured/kansuu46/mae-ato-im-mae-2026-09-08.tsv');
    if (!fs.existsSync(道)) throw new Error('★入れる前の 実測が 無い★ … ' + path.basename(道));
    const 前 = new Map(fs.readFileSync(道, 'utf-8').split('\n')
      .map((l) => l.split('\t')).filter((p) => p.length >= 3).map((p) => [p[0], p[2]]));
    let 直る = 0, 変わらない = 0, 紙に無い = 0;
    for (const g of 金) {
      if (g.型 === 'Double') continue;               /* ★文字を 返す 物だけ★ */
      const なま = 前.get(g.式);
      if (なま === undefined) { 紙に無い++; continue; }
      if (/^#/.test(なま)) continue;                 /* 誤り＝別の 山 */
      if (なま === g.答) { 変わらない++; continue; }  /* 元から 合っていた */
      if (F.字を直す(なま) === g.答) 直る++;
    }
    console.log('      … 入れる 前の 字に 15桁を かけると ★' + 直る + '本★ 実Excel と 同じに なる'
      + '（元から 合っていた ' + 変わらない + '本）');
    if (紙に無い) throw new Error('★紙に 無い 式が ' + 紙に無い + '本★＝紙と 金が ずれている');
    if (直る !== 78) throw new Error('★' + 直る + '本★＝実測は 78本（★数が 違う＝物差しか 直しが 変わった★）');
  });
  T('★★触らない 4個に かけると 赤に なる（かけては いけない 証拠）★★', () => {
    const 悪い = F.十五桁(String(Math.PI));
    if (悪い === String(Math.PI)) {
      throw new Error('★15桁を かけても 字が 変わらない★＝この 試験は 何も 見ていない');
    }
    console.log('      … IMARGUMENT(-2) に かけると ' + String(Math.PI) + ' → ' + 悪い + '（★実Excel は 17桁★）');
  });
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (直に走った) process.exit(fail ? 1 : 0);
