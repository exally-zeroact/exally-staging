/* formula-areas.test.mjs — ★AREAS を 実Excel の 答えと 突き合わせる★（2026-09-08）
 *
 *  ★★本番の 道で 押す★★
 *    JS層 → convertFormula → HyperFormula
 *
 *  ★答え★ 実Excel に ★4回★ 打たせた ★89本★
 *    `golden-areas-2026-09-07.tsv` … 36本（基本）
 *    `golden-areas2-…` … 22本（範囲で ない 物）
 *    `golden-areas3-…` … 16本（★答えが 割れる 組★）
 *    `golden-areas4-…` … 15本（★残していた 穴★）
 *    Excel 16.0 build 20326
 *
 *  ★★なぜ convertFormula（式の 字を 書き換える 段）で やるか★★
 *    エンジンは ★とびとびを 式を 読む 所で 断る★
 *      `=AREAS((A1:B3,D1:D2))` … Parsing error
 *    ⇒★関数の 所まで 届かない＝プラグインでは 受け取れない★
 *    ⇒★JS層も 駄目★（式の 一番 外側でしか 効かない／AREAS は 入れ子で 使われる）
 *    ⇒★★字を 数に 置き換えれば 入れ子も そのまま 正しく なる★★
 *
 *  ★★出さない 形★★
 *    ★かっこが 入る 物 全部★（関数入り／シート名の かっこ）と
 *    ★重ならない 重なり★（実Excel は #NULL!／★エンジンに その 誤りが 無い★）
 *    ⇒ `#NAME?`（＝「まだです」）
 *
 *  使い方: node tests/formula-areas.test.mjs [--self-test]
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

const A = require_(path.join(ROOT, 'lib/formula-areas.js'));
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

/* ★実Excel を 測った 時と 同じ 中身★（A1:I9 に 行*10+列） */
function 土台() {
  const 表 = [];
  for (let r = 0; r < 10; r++) {
    const 行 = [];
    for (let c = 0; c < 11; c++) 行.push(r < 9 && c < 9 ? (r + 1) * 10 + (c + 1) : null);
    表.push(行);
  }
  return 表;
}
function 押す(式) {
  let 後;
  try { 後 = EF.convertFormula(式); } catch (e) { return { 値: '★書き換えで 例外★', 後: '' }; }
  try {
    const 表 = 土台();
    表[0][10] = 後;                        /* K1 … 測った 時と 同じ 場所 */
    hf.setSheetContent(SID, 表);
    const v = hf.getCellValue({ sheet: SID, row: 0, col: 10 });
    return { 値: (v && v.type) ? 赤の名(v.type) : v, 後: 後 };
  } catch (e) { return { 値: '★engine で 例外★', 後: 後 }; }
}

const 金 = [];
for (const f of ['golden-areas-2026-09-07.tsv', 'golden-areas2-2026-09-07.tsv',
  'golden-areas3-2026-09-07.tsv', 'golden-areas4-2026-09-07.tsv']) {
  const t = fs.readFileSync(path.join(ROOT, 'docs/measured/kansuu46', f), 'utf-8');
  for (const l of t.split('\n')) {
    if (!l || l.startsWith('#')) continue;
    const p = l.split('\t');
    if (!/^=AREAS\(/.test(p[1] || '')) continue;
    金.push({ 式: p[1], 答: p[2], 元: f });
  }
}

console.log('\n[formula-areas] AREAS — 実Excel の 答えと 突き合わせ');
console.log('  実Excel の 答え … ' + 金.length + '本（4回・Excel 16.0 build 20326）');

T('★★物差しが 生きている★★', () => {
  if (!積) throw new Error('exally-formula を 積めていない');
  const r = 押す('=AREAS(B2:D4)');
  if (r.値 !== 1) throw new Error('=AREAS(B2:D4) が 1 で ない … ' + r.値 + '（' + r.後 + '）');
});

T('★★かっこが 合わない 式を 作らない★★', () => {
  const 悪い = [];
  for (const 行 of 金) {
    let 後; try { 後 = EF.convertFormula(行.式); } catch (e) { 悪い.push(行.式 + ' → 例外'); continue; }
    const 開 = (後.match(/\(/g) || []).length, 閉 = (後.match(/\)/g) || []).length;
    if (開 !== 閉) 悪い.push(行.式 + ' → ' + 後);
  }
  if (悪い.length) throw new Error('★' + 悪い.length + '本★\n      ' + 悪い.join('\n      '));
});

T('★★お客さんの 画面に #ERROR! を 出さない★★（とびとびは 式を 読む 所で 止まる 形）', () => {
  const 悪い = [];
  for (const 行 of 金) {
    const r = 押す(行.式);
    if (String(r.値) === '#ERROR!') 悪い.push(行.式 + ' → ' + r.後 + ' → #ERROR!');
  }
  if (悪い.length) throw new Error('★' + 悪い.length + '本 #ERROR!★\n      ' + 悪い.join('\n      '));
});

T('★★出した 答えは 1本 残らず 実Excel と 同じ★★（出さない 物は 別に 数える）', () => {
  const 外れ = [];
  let 出した = 0, まだ = 0, 聞けない = 0;
  for (const 行 of 金) {
    /* ★実Excel が 式として 受け取らなかった 物は ★答えが 無い★＝突き合わせない★ */
    if (/受け取らない|打てない/.test(行.答)) { 聞けない++; continue; }
    const r = 押す(行.式);
    if (String(r.値) === '#NAME?') { まだ++; continue; }
    出した++;
    const 正 = 行.答;
    const よい = /^-?[0-9.]+$/.test(正)
      ? Math.abs(Number(r.値) - Number(正)) <= 1e-9
      : String(r.値) === 正;
    if (!よい) 外れ.push(行.式 + '\n        正 ' + 正 + ' ／ 出 ' + r.値 + '（' + r.後 + '）');
  }
  console.log('      出した … ' + 出した + '本 ／ ★まだ（#NAME?）… ' + まだ + '本★'
    + ' ／ 実Excel が 受け取らない … ' + 聞けない + '本');
  /* ★★下限は ★実測の 値★を 書く★★（2026-09-08）
     ⇒ はじめ 30 と 書いて 赤に なった。★数が 悪いのでは なく 私の 下限が 当て推量★だった
     ⇒★訳★ 名前（ひとつ／とびとび／みっつ）は ★node には 一覧が 無い★ので 出せない
       ⇒ ★画面では 名前は 先に 参照に 開かれる★ので お客さんの 道では 出ます
     ⇒★実測の 本番の 道＝★29本★（`docs/measured/kansuu46/areas-awaseru.txt`）
     ⇒★1本でも 減ったら 赤★＝★静かに 減る 間違いを 止める★ */
  if (出した < 29) throw new Error('★出した 数が 減った★（' + 出した + '本／実測は 29本）'
    + '＝★『出さない』側に 倒れる 間違いが 入った★');
  if (外れ.length) throw new Error('★' + 外れ.length + '本 違う★\n      ' + 外れ.join('\n      '));
});

T('★★エンジンが 断る「とびとび」が 通る★★（ここが 出来なかった 所）', () => {
  /* ★直す前は 式を 読む 所で 止まっていた★
     `=AREAS((A1:B3,D1:D2))` … Parsing error */
  for (const [式, 正] of [['=AREAS((B2:D4,E5,F6:I9))', 3], ['=AREAS((A1,A1))', 2],
    ['=AREAS((A1,B2,C3,D4,E5,F6,G7,H8))', 8]]) {
    const r = 押す(式);
    if (r.値 !== 正) throw new Error(式 + ' … 正 ' + 正 + ' ／ 出 ' + r.値 + '（' + r.後 + '）');
  }
});

T('★入れ子でも 効く（式の 一番 外側でなくても）★', () => {
  for (const [式, 正] of [['=SUM(AREAS((A1,A2)),1)', 3], ['=AREAS((A1,A2))+AREAS(B2:D4)', 3],
    ['=IF(AREAS((A1,A2))=2,"ふたつ","ちがう")', 'ふたつ']]) {
    const r = 押す(式);
    if (String(r.値) !== String(正)) throw new Error(式 + ' … 正 ' + 正 + ' ／ 出 ' + r.値 + '（' + r.後 + '）');
  }
});

T('★★出せる 答えを「まだ」に していない★★（#VALUE! は エンジンが 持っている）', () => {
  const r = 押す('=AREAS((Sheet1!A1,二枚目!A1))');
  if (String(r.値) !== '#VALUE!') {
    throw new Error('★別の シートを またぐ とびとびは #VALUE! のはず★ … ' + r.値 + '（' + r.後 + '）');
  }
});

T('★★まだの 形は #NAME?（間違った 数を 出さない）★★', () => {
  const まだ = ['=AREAS(INDEX(A1:C3,1,1))', '=AREAS(INDIRECT("A1"))',
    "=AREAS('売上(旧)'!A1)", '=AREAS(B2:D4 A1)'];
  const 悪い = [];
  for (const 式 of まだ) {
    const r = 押す(式);
    if (String(r.値) !== '#NAME?') 悪い.push(式 + ' → ' + r.値 + '（' + r.後 + '）');
  }
  if (悪い.length) throw new Error('★' + 悪い.length + '本 答えを 出してしまった★\n      ' + 悪い.join('\n      '));
});

T('★閉じない かっこで 式を 壊さない★', () => {
  const 壊れそう = "=AREAS('売上(旧'!A1)";
  const 後 = EF.convertFormula(壊れそう);
  if (後 !== 壊れそう) throw new Error('★触った★ … ' + 後);
});

T('★★AI に 渡る 紙が「動く」とも「動かない」とも 言っていない★★', () => {
  const 紙 = fs.readFileSync(path.join(ROOT, 'prompt/kansuu.md'), 'utf-8');
  const 行 = 紙.split('\n').filter((l) => /AREAS/.test(l));
  if (!行.length) throw new Error('AREAS の 行が 無い');
  const 全 = 行.join(' ');
  if (!/出る 形/.test(全) || !/まだの 形/.test(全)) {
    throw new Error('★『出る 形』と『まだの 形』を 名指ししていない★\n      ' + 全);
  }
  /* ★3つ目の 棚の 見出しが 在るか★ */
  if (!/形に よって 出る 物/.test(紙)) throw new Error('★3つ目の 棚の 見出しが 無い★');
});

T('★★名前は ★本番と 同じ 順★で 押す（先に 開いてから 数える）★★', () => {
  /* ★★はじめ ここは 間違った 試験でした（そのまま 書きます）★★
       部品に「名前 → か所の 数」の 表を 渡して 押していました。
       ⇒★でも 本番では ★その 表は 1度も 使われません★★
         画面は `convertFormula` の 頭で ★`名前の箱.開く` を 先に 走らせる★ので、
         AREAS に 届く 時には ★もう 参照に なっています★
       ⇒★★お客さんが 通らない 道を 試して「動く」と 言っていました★★
       ⇒ 見張り `unused-param` が「使っていない 口」と 数えて 教えてくれた
       ⇒★口を 消して、ここを ★本番と 同じ 順★に 書き直した★

     ★本番の 順で 押すと 何が 出るか（実測）★
       `=AREAS((とびとび,A1))` → 開く → `=AREAS((($B$2:$D$4,$F$6:$I$9),A1))`
       ⇒★かっこが 入れ子に なる★＝★これを 数えられないと 名前を 使った 人だけ #NAME?★ */
  const N = require_(path.join(ROOT, 'lib/named-ranges.js'));
  const 箱 = N.作る();
  箱.足す('ひとつ', '$B$2:$D$4');
  箱.足す('とびとび', '$B$2:$D$4,$F$6:$I$9');
  箱.足す('みっつ', '$A$1,$C$3,$E$5');
  const 組 = [['=AREAS(ひとつ)', 1], ['=AREAS(とびとび)', 2], ['=AREAS(みっつ)', 3],
    ['=AREAS((ひとつ,A1))', 2], ['=AREAS((とびとび,A1))', 3], ['=AREAS((A1,とびとび))', 3],
    ['=AREAS((とびとび,とびとび))', 4], ['=AREAS((みっつ,とびとび))', 5]];
  const 外れ = [];
  for (const [式, 正] of 組) {
    const 開いた = 箱.開く(式, 'Sheet1');
    const 後 = EF.convertFormula(開いた);
    if (後 !== '=' + 正) 外れ.push(式 + ' → 開く ' + 開いた + ' → ' + 後 + '（正 =' + 正 + '）');
  }
  if (外れ.length) throw new Error('★' + 外れ.length + '本 違う★\n      ' + 外れ.join('\n      '));
});

T('★★開けなかった 名前は 出さない★★（実Excel も #NAME?）', () => {
  /* ★名前が 決めてなければ 画面は 開けない★ ⇒ そのまま AREAS に 来る
     ⇒★実Excel も #NAME?★なので ★出さないのが 正しい★ */
  const r = 押す('=AREAS(しらないなまえ)');
  if (String(r.値) !== '#NAME?') throw new Error('★答えを 出してしまった★ … ' + r.値 + '（' + r.後 + '）');
});

T('★出す 名簿は lib が 正本★', () => {
  /* ★★2026-09-08 に 直した＝★この 棚は もう AREAS 専用では ない★★
     ⇒ 前は 「棚 ＝ formula-areas.js が 足した 名前」と ★完全に 同じ★を 求めていた
     ⇒ ISOMITTED が 棚に 来た 日に ★赤★に なった
       （ISOMITTED は lib では なく `exally-formula.js` の JS層に 居る）
     ⇒★★見張りを 書いた 時に ★自分しか 居ない★と 決め打っていた★★
     ⇒ 今は こう 守る
       ①★formula-areas.js が 足した 名前は ★全部★ 棚に 在る★（棚が 落としていない）
       ②★AREAS が「足さない」棚に 居ない★
       ③★棚に 居る 他人は ★自分の 見張りを 持っている★★
         ⇒ 名前と 見張りの 対を ここに 書き、★居るのに 見張りが 無い★物を 赤に する
         ⇒★これが 無いと「棚に 名前を 足しただけ」で AI の 紙に 出てしまう★ */
  const 台帳 = require_(path.join(ROOT, 'lib/formula-extra.js')).数える();
  const 棚 = Object.keys(台帳.形で動く || {});
  const 足した = [...A.足した名前()];
  const 落ちた = 足した.filter((n) => 棚.indexOf(n) < 0);
  if (落ちた.length) {
    throw new Error('★lib が 足したのに 棚に 無い★ … ' + 落ちた.join(',')
      + ' ／ 台帳=' + 棚.join(','));
  }
  if ('AREAS' in (台帳.足さない || {})) throw new Error('★AREAS が まだ「足さない」の 棚に 居る★');

  /* ★棚に 居る 他人★＝名前 → その 名前を 守る 試験（★居るのに 見張りが 無いと 赤★）
     ★★この 対の 表は ★手で 書いている★★（指示役 2026-09-08 の 指摘）
       ⇒ 手で 写す 物は ★減る／ずれる★＝今日 3回 出た 病気
       ⇒ だから ★在るか だけでは 足りない★
       ⇒★★その ファイルの 中で ★その 名前を 実際に 押しているか★ まで 見る★★
         （「＝名前（」の ような ★式の 形★で 出てくるか
          ＝『名前を 書いただけ』の ファイルを 対に しても 通らない）
       ★★ここに その 字を 書かない★★
         ⇒ 一度 説明に 実物の 字を 書いたら ★この ファイル自身が 条件を 満たして★
           『対が 間違っている』を ★見逃した★（2026-09-08 に 踏んだ）
         ⇒★見張りの 説明が 見張りを 通してしまう＝真っ先に 自分に 通す★ */
  const 他人の見張り = { ISOMITTED: 'tests/isomitted.test.mjs' };
  const 見張り無し = [], 押していない = [];
  for (const n of 棚) {
    if (足した.indexOf(n) >= 0) continue;              /* AREAS 自身＝この ファイルが 見張る */
    const 道 = 他人の見張り[n];
    if (!道 || !fs.existsSync(path.join(ROOT, 道))) { 見張り無し.push(n); continue; }
    /* ★この ファイル自身を 他人の 見張りに しない★（自分で 自分を 通さない） */
    if (path.resolve(ROOT, 道) === path.resolve(fileURLToPath(import.meta.url))) {
      押していない.push(n + ' → ' + 道 + '（★この ファイル自身★）');
      continue;
    }
    const 中身 = fs.readFileSync(path.join(ROOT, 道), 'utf-8');
    if (中身.indexOf('=' + n + '(') < 0) 押していない.push(n + ' → ' + 道);
  }
  if (見張り無し.length) {
    throw new Error('★棚に 居るのに 見張りが 無い★ … ' + 見張り無し.join(',')
      + '（★棚に 名前を 足しただけで AI の 紙に 出てしまう★）');
  }
  if (押していない.length) {
    throw new Error('★対に 書いた ファイルが その 名前を 1度も 押していない★ … '
      + 押していない.join(' ／ ') + '（★対の 表を 手で 書いた ずれ★）');
  }
});

if (自己試験) {
  console.log('\n[self-test] ★物差しが 効いているか★');
  T('★★書き換えを 外すと 赤に なる★★（とびとびが 通らなく なる）', () => {
    /* ★壊すのは ★写し★★＝repo の ファイルは 1バイトも 触らない */
    const 生 = fs.readFileSync(path.join(ROOT, 'exally-formula.js'), 'utf-8');
    if (!/_rewriteAreas\(f\);/.test(生)) throw new Error('★書き換えを 呼んでいない★');
    /* 部品を 直に 押して「出さない」形を 数える */
    let 出た = 0;
    for (const 中 of ['(A1:B3,D1:D2)', 'B2:D4', '(A1,A1)']) {
      if (typeof A.区画を数える(中, null) === 'number') 出た++;
    }
    if (出た !== 3) throw new Error('★部品が 3本 とも 数を 返さない★＝' + 出た + '本');
  });
  T('★★『出さない』に 倒れる 間違いを 数で 見張る★★', () => {
    /* ★2026-09-08 に 2つ 踏んだ★
       ・名前が シートの 数え方を 汚していた（(とびとび,A1) が #VALUE! に なった）
       ・シート名の 中の 空白で 割っていた（'二 枚目'!A1 が 出せなく なった）
       ⇒★どちらも「出さない」側に 倒れる＝★赤に ならない／静かに 減る★★
       ⇒★だから ★出す 数に 下限を 置く★ */
    let 数 = 0;
    for (const 行 of 金) {
      if (/受け取らない|打てない/.test(行.答)) continue;
      if (typeof A.区画を数える(行.式.replace(/^=AREAS\(/, '').replace(/\)$/, ''), null) === 'number') 数++;
    }
    console.log('      … 部品が 数を 返した … ' + 数 + '本');
    if (数 < 27) throw new Error('★' + 数 + '本 しか 出ない★＝静かに 減っている（実測は 27本）');
  });
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (直に走った) process.exit(fail ? 1 : 0);
