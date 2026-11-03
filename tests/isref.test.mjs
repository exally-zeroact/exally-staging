/* isref.test.mjs — ★ISREF が 実Excel と 同じ 答えを 出すか★（2026-09-07）
 *
 *  ★★本番に 在った 壊れ（実測）★★
 *    `convertFormula` の ISREF は `[^)]+` で かっこの 中を 取っていた。
 *    ⇒★最初の `)` で 止まる＝入れ子を 数えない★
 *      `=ISREF(SUM(A1:A2))` → `=FALSE())` … ★閉じかっこが 1つ 多い★ ⇒★#ERROR!★
 *      `=ISREF(IF(TRUE,A1,B1))` → `=FALSE(),A1,B1))`
 *    ⇒ さらに TRUE に する 形が ★`A1` と `A1:B2` だけ★
 *      `=ISREF(A:A)` `=ISREF(1:1)` `=ISREF(Sheet1!A1)` は
 *      ★実Excel が TRUE／うちは FALSE★＝★誤りも 出ない「黙った 逆の 答え」★
 *    ⇒★AI に 渡る 紙には「★打てば そのまま 動く★」と 書いてあった★
 *
 *  ★答え★ … `docs/measured/kansuu46/golden-isref-2026-09-07.tsv`（25本）
 *    Excel 16.0 build 20326
 *
 *  ★★今回は「嘘を 止める」所まで★★（指示役 2026-09-07・2a）
 *    ・★字だけで 確かに 決まる 形★ … TRUE / FALSE を 出す
 *    ・★かっこが 入る 形は 全部★ …… ★出さない★ ⇒ `#NAME?`（＝「まだです」）
 *      （関数入り／★シート名に かっこが 在る 物も★）
 *      ⇒`INDIRECT("A1")` は TRUE ／`INDIRECT("zzz")` は FALSE
 *        ＝★指せたかは 字だけでは 決まらない★
 *      ⇒★★間違った TRUE/FALSE より `#NAME?`（＝まだです）の 方が 良い★★
 *    ・★名前★ も まだ 出さない（名前の 一覧を 見ないと 決まらない）
 *
 *  使い方: node tests/isref.test.mjs [--self-test]
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

const HFns = require_(path.join(ROOT, 'hyperformula.full.min.js'));
const EF = require_(path.join(ROOT, 'exally-formula.js'));
const 積 = EF.registerExallyFunctions(HFns) === true;
const HF0 = HFns.HyperFormula;
const hf = HF0.buildEmpty({ licenseKey: 'gpl-v3' });
const SID = hf.getSheetId(hf.addSheet('Sheet1'));
EF.initExallyFormula(hf);

const 赤の名 = (t) => ({
  NA: '#N/A', DIV_BY_ZERO: '#DIV/0!', VALUE: '#VALUE!', NUM: '#NUM!',
  NAME: '#NAME?', REF: '#REF!', CYCLE: '#CYCLE!', ERROR: '#ERROR!', SPILL: '#SPILL!',
}[t] || ('#' + t));

/* ★実Excel を 測った 時と 同じ 中身★（A1:E5 に 行*10+列／E5 だけ =1/0） */
function 土台() {
  const 表 = [];
  for (let r = 0; r < 6; r++) {
    const 行 = [];
    for (let c = 0; c < 9; c++) 行.push(r < 5 && c < 5 ? (r + 1) * 10 + (c + 1) : null);
    表.push(行);
  }
  表[4][4] = '=1/0';
  return 表;
}
function 押す(式) {
  let 後;
  try { 後 = EF.convertFormula(式); } catch (e) { return { 値: '★書き換えで 例外★', 後: '' }; }
  try {
    const 表 = 土台();
    表[0][7] = 後;                         /* H1 … 測った 時と 同じ 場所 */
    hf.setSheetContent(SID, 表);
    const v = hf.getCellValue({ sheet: SID, row: 0, col: 7 });
    return { 値: (v && v.type) ? 赤の名(v.type) : v, 後: 後 };
  } catch (e) { return { 値: '★engine で 例外★', 後: 後 }; }
}

const 金 = fs.readFileSync(path.join(ROOT, 'docs/measured/kansuu46/golden-isref-2026-09-07.tsv'), 'utf-8')
  .split('\n').filter((l) => l && !l.startsWith('#'))
  .map((l) => { const p = l.split('\t'); return { 式: p[1], 答: p[2], 型: p[3] }; });

console.log('\n[isref] ISREF — 実Excel の 答えと 突き合わせ');
console.log('  実Excel の 答え … ' + 金.length + '本（Excel 16.0 build 20326）');

T('★★物差しが 生きている★★', () => {
  if (!積) throw new Error('exally-formula を 積めていない');
  const r = 押す('=ISREF(A1)');
  if (r.値 !== true) throw new Error('=ISREF(A1) が TRUE で ない … ' + r.値);
});

T('★★かっこが 合わない 式を 作らない★★（本番に 在った 壊れ）', () => {
  const 悪い = [];
  for (const 行 of 金) {
    let 後; try { 後 = EF.convertFormula(行.式); } catch (e) { 悪い.push(行.式 + ' → 例外'); continue; }
    const 開 = (後.match(/\(/g) || []).length, 閉 = (後.match(/\)/g) || []).length;
    if (開 !== 閉) 悪い.push(行.式 + ' → ' + 後 + '（開' + 開 + ' 閉' + 閉 + '）');
  }
  if (悪い.length) throw new Error('★' + 悪い.length + '本 かっこが 合わない★\n      ' + 悪い.join('\n      '));
});

T('★★お客さんの 画面に #ERROR! を 出さない★★', () => {
  const 悪い = [];
  for (const 行 of 金) {
    const r = 押す(行.式);
    if (String(r.値) === '#ERROR!') 悪い.push(行.式 + ' → ' + r.後 + ' → #ERROR!');
  }
  if (悪い.length) throw new Error('★' + 悪い.length + '本 #ERROR!★\n      ' + 悪い.join('\n      '));
});

T('★★出した 答えは 1本 残らず 実Excel と 同じ★★（出さない 物は 別に 数える）', () => {
  const 外れ = [];
  let 出した = 0, まだ = 0;
  for (const 行 of 金) {
    const r = 押す(行.式);
    /* ★まだ 出さない＝#NAME?★（＝「まだです」と 言っている） */
    if (String(r.値) === '#NAME?') { まだ++; continue; }
    出した++;
    let 正 = 行.答;
    let よい;
    if (正 === 'True' || 正 === 'False') よい = String(r.値).toUpperCase() === 正.toUpperCase();
    else if (/^-?[0-9.]+$/.test(正)) よい = Math.abs(Number(r.値) - Number(正)) <= 1e-9;
    else よい = String(r.値) === 正;
    if (!よい) 外れ.push(行.式 + '\n        正 ' + 正 + ' ／ 出 ' + r.値 + '（' + r.後 + '）');
  }
  console.log('      出した … ' + 出した + '本 ／ ★まだ 出さない（#NAME?）… ' + まだ + '本★');
  if (出した < 8) throw new Error('★出した 数が 少なすぎる★（' + 出した + '本）＝物差しが 空洞');
  if (外れ.length) throw new Error('★' + 外れ.length + '本 違う★\n      ' + 外れ.join('\n      '));
});

T('★★前は 逆の 答えを 出していた 3本が 直っている★★（誤りも 出ない＝一番 見つけにくい 形）', () => {
  const 直す前は逆 = { '=ISREF(A:A)': true, '=ISREF(1:1)': true, '=ISREF(Sheet1!A1)': true };
  const 外れ = [];
  for (const 式 of Object.keys(直す前は逆)) {
    const r = 押す(式);
    if (r.値 !== 直す前は逆[式]) 外れ.push(式 + ' → ' + r.値 + '（' + r.後 + '）');
  }
  if (外れ.length) throw new Error('★' + 外れ.length + '本 まだ 逆★\n      ' + 外れ.join('\n      '));
});

T('★入れ子でも 効く（式の 一番 外側でなくても）★', () => {
  const 期待 = { '=IF(ISREF(A1),"は","い")': 'は', '=SUM(ISREF(A1)*1,1)': 2 };
  for (const 式 of Object.keys(期待)) {
    const r = 押す(式);
    if (String(r.値) !== String(期待[式])) {
      throw new Error(式 + ' … 正 ' + 期待[式] + ' ／ 出 ' + r.値 + '（' + r.後 + '）');
    }
  }
});

T('★★かっこが 入る 形は 全部 まだ（シート名の かっこも）★★', () => {
  /* ★2026-09-07 指示役が 突いた★
     `_isRefKotae` は ★かっこが 1つでも 在れば 出さない★（`if (/[()]/.test(中)) return null;`）
     ⇒ だから ★関数入りだけ★では なく ★シート名の かっこ★も「まだ」に なる
     ⇒ 実測 … `=ISREF('売上(旧)'!A1)` は 実Excel ★TRUE★／うちは ★#NAME?★
       （★答えが 違うのでは なく「出さない」＝安全な 側★）
     ⇒★★『まだの 範囲』を 実際より 狭く 書かない★★
     ⇒★`'売上(旧)'` の ような シート名は ★珍しく ない★★ */
  const 後 = EF.convertFormula("=ISREF('売上(旧)'!A1)");
  if (!/ISREF\.MADA/.test(後)) throw new Error('かっこ入りの シート名を 出してしまった … ' + 後);
  /* ★閉じない かっこが シート名に 在っても ★式を 壊さない★★ */
  const 壊れそう = "=ISREF('売上(旧'!A1)";
  const 後2 = EF.convertFormula(壊れそう);
  if (後2 !== 壊れそう) throw new Error('★閉じない かっこで 式を 触った★ … ' + 後2);
});

T('★★実Excel が 選んだ 16本と 1本 残らず 同じ★★（2026-09-08 の 総ざらい）', () => {
  /* ★★司さん 2026-09-08「答え確かめて」で 見つかった★★
     ⇒ `=ISREF(TRUE)` が ★#NAME?★に なっていた（実Excel は ★False★）
     ⇒★#45 で 私が 本番に 入れた 戻り★（★直す前は FALSE＝合っていた★）
     ⇒ 訳 … `convertFormula` は ★先に★ `TRUE` を `TRUE()` に 直すので
       「かっこが 在れば 決められない」に 引っかかっていた
     ★★引数の 形は 実Excel が 選んだ★★
       `docs/measured/kansuu46/golden-346-2026-09-08.tsv`
     ⇒★私が 手で 選ぶと ★自分が 通る 形しか 選ばない★★ */
  const 道 = path.join(ROOT, 'docs/measured/kansuu46/golden-346-2026-09-08.tsv');
  if (!fs.existsSync(道)) throw new Error('実測の 紙が 無い … golden-346-2026-09-08.tsv');
  const 組 = fs.readFileSync(道, 'utf-8').split('\n')
    .filter((l) => l.startsWith('ISREF\t') && l.split('\t').length === 4)
    .map((l) => l.split('\t'));
  if (組.length < 16) throw new Error('★実測が ' + 組.length + '本 しか 無い★（16本 在るはず）');
  const 外れ = [];
  for (const [, 式, 正] of 組) {
    const 後 = EF.convertFormula(式);
    const 出 = (後 === '=TRUE()') ? 'True' : (後 === '=FALSE()') ? 'False' : '(まだ)';
    if (出 !== 正) 外れ.push(式 + '  正 ' + 正 + ' ／ 出 ' + 出 + '（' + 後 + '）');
  }
  console.log('      実Excel が 選んだ 形 … ' + 組.length + '本');
  if (外れ.length) throw new Error('★' + 外れ.length + '本 違う★\n      ' + 外れ.join('\n      '));
});

T('★AI に 渡る 紙が「動く」と 言い切っていない★', () => {
  const 紙 = fs.readFileSync(path.join(ROOT, 'prompt/kansuu.md'), 'utf-8');
  const 行 = 紙.split('\n').filter((l) => /ISREF/.test(l));
  if (!行.length) throw new Error('ISREF の 行が 無い');
  for (const l of 行) {
    if (/打てば そのまま 動く/.test(l)) {
      throw new Error('★「打てば そのまま 動く」が 残っている★＝AI が 客に 嘘を 言う\n      ' + l);
    }
  }
  /* ★★2026-09-08 に 書き直した★★
     ⇒ 前は ★昔の 文を そのまま★ 探していた（「かっこが 入っている 形は 全部 まだ」）
     ⇒ 台帳を ★正しく★ 直したのに 見張りが 赤に なった
       （TRUE()／FALSE() は かっこが 在っても 動くので「全部 まだ」は ★もう 嘘★）
     ⇒★★見張りが 見るのは ★文の 形★では なく ★守りたい 中身★★★
     ★この 見張りが 今 守る 物（1つずつ）★
       ①「打てば そのまま 動く」と 言い切っていない（上の 段で 見ている）
       ②★出せない 形が 在る★と 言っている（「まだ」）
       ③★その時 何が 出るか★を 言っている（「#NAME?」）
       ④★「全部 動く」と 読める 言い方を していない★ */
  const 全 = 行.join(' ');
  const 足りない = [];
  if (!/まだ/.test(全)) 足りない.push('②「まだ」＝出せない 形が 在る事を 言っていない');
  if (!/#NAME\?/.test(全)) 足りない.push('③「#NAME?」＝その時 何が 出るかを 言っていない');
  if (/全部\s*動く|どの 形でも 動く/.test(全)) 足りない.push('④「全部 動く」と 読める 言い方が 在る');
  if (足りない.length) {
    throw new Error('★AI に 渡る 紙が 足りない★\n      ' + 足りない.join('\n      ')
      + '\n      ---\n      ' + 行.join('\n      '));
  }
});

/* ★★作り物の 名前が お客さんの 外へ 出ないか★★
   `ISREF.MADA` は ★Excel に 無い 名前★です。
   ⇒ もし お客さんの ファイルに 残ったら ★開くたびに 壊れた 式が 出ます★
   ⇒★ここ（node の 中）では 書き出しまで 押せません★＝
     ★本物の ブラウザで 押した 記録★を 名指しで 見ます */
T('★★「外へ 漏れない」は ★本物で 押した 記録★が 在る★★', () => {
  const 道 = path.join(ROOT, 'docs/measured/kansuu46/isref-mada-ga-moreru-ka.txt');
  if (!fs.existsSync(道)) throw new Error('証拠の 紙が 無い … docs/measured/kansuu46/isref-mada-ga-moreru-ka.txt');
  const 紙 = fs.readFileSync(道, 'utf-8');
  if (!/0 赤/.test(紙)) throw new Error('赤が 残っている 紙');
  for (const 印 of [
    '数式バーが 元の 式の まま',
    '控えに 作り物の 名前が 無い',
    '書き出した .xlsx に 作り物の 名前が 1つも 無い',
    'ファイルの 生の バイトにも 作り物の 名前が 無い',
    '「Excelに書き出す」を 本物の マウスで 押せた',
  ]) {
    if (紙.indexOf(印) < 0) throw new Error('紙に 無い … ' + 印);
  }
  /* ★書き出した 式が 本当に 載っているか★（紙だけ 置いて 中身が 空を 止める） */
  if (!/Sheet1!D1\s+=ISREF\(INDEX/.test(紙)) throw new Error('書き出した 式が 紙に 無い');
});

if (自己試験) {
  console.log('\n[self-test] ★物差しが 効いているか★');
  T('★★前の 作り（[^)]+）に 戻すと 赤に なる★★', () => {
    /* ★壊すのは ★写し★★＝repo の ファイルは 1バイトも 触らない */
    const 前の作り = (f) => f.replace(/\bISREF\s*\(([^)]+)\)/gi,
      (m, arg) => (/^[A-Z]+\d+(:[A-Z]+\d+)?$/i.test(arg.trim()) ? 'TRUE()' : 'FALSE()'));
    let 壊れ = 0;
    for (const 行 of 金) {
      const 後 = 前の作り(行.式);
      const 開 = (後.match(/\(/g) || []).length, 閉 = (後.match(/\)/g) || []).length;
      if (開 !== 閉) 壊れ++;
    }
    console.log('      … 前の 作りだと ★' + 壊れ + '本★ かっこが 合わない');
    if (壊れ < 5) throw new Error('★壊れが ' + 壊れ + '本 しか 出ない★＝この 試験は 何も 見ていない（実測は 8本）');
  });
  T('★決められない 形を うっかり 出していないか（#NAME? が 1本 以上 在る）★', () => {
    const r = 押す('=ISREF(INDIRECT("A1"))');
    if (String(r.値) !== '#NAME?') {
      throw new Error('★字だけでは 決まらない 形に 答えを 出した★ … ' + r.値 + '（' + r.後 + '）');
    }
  });
  /* ★★2026-09-08 に 足した＝紙の 見張りが ★本当に 噛むか★★
     ⇒ 訳 … 前の 見張りは ★昔の 文を そのまま★ 探していたので
       ★台帳を 正しく 直した だけで 赤★に なった（＝守りたい 物を 見ていなかった）
     ⇒ だから ★悪い 紙を 3通り 作って 3通りとも 赤に なるか★を 見る
     ★壊すのは 写しだけ★＝prompt/kansuu.md は 1バイトも 触らない */
  T('★★紙の 見張りは 悪い 紙 3通りを 全部 捕まえる★★', () => {
    const 見る = (l) => {
      const 全 = l;
      const 足りない = [];
      if (!/まだ/.test(全)) 足りない.push('②');
      if (!/#NAME\?/.test(全)) 足りない.push('③');
      if (/全部\s*動く|どの 形でも 動く/.test(全)) 足りない.push('④');
      return 足りない;
    };
    const 悪い紙 = [
      ['②「まだ」が 無い', '- ISREF … マスを 指しているか。A1 は 動く。関数入りは #NAME? に なる'],
      ['③「#NAME?」が 無い', '- ISREF … マスを 指しているか。A1 は 動く。関数入りは まだ'],
      ['④「全部 動く」と 書いた', '- ISREF … マスを 指しているか。★全部 動く★（まだ／#NAME? も 書いてある）'],
    ];
    const 逃した = [];
    for (const [札, 紙] of 悪い紙) if (!見る(紙).length) 逃した.push(札);
    console.log('      … 悪い 紙 ' + 悪い紙.length + '通りを 試した ／ 逃した ' + 逃した.length + '通り');
    if (逃した.length) throw new Error('★見張りが 素通りさせた★ … ' + 逃した.join(' ／ '));
    /* ★良い 紙（＝今 本番に 在る 物）は 通らないと おかしい★ */
    const 本物 = fs.readFileSync(path.join(ROOT, 'prompt/kansuu.md'), 'utf-8')
      .split('\n').filter((l) => /ISREF/.test(l)).join(' ');
    if (見る(本物).length) throw new Error('★今の 紙を 落とした★ … ' + 見る(本物).join(' '));
  });
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (直に走った) process.exit(fail ? 1 : 0);
