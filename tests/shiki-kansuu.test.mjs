/* shiki-kansuu.test.mjs — ★四角（A1:A3）と 関数★（2026-09-13）
 *
 *  ★★土台を 自分で 作る ④枚目の 後半★★
 *
 *  ★★下の 37通りは 全部 実Excel に 打って 読んだ 物です★★
 *    （2026-09-13・道具 `docs/measured/toru-shikaku-kansuu.ps1`／Excel 16.0 build 20326）
 *    ★1つも 当て推量が 在りません★
 *
 *  ★★材料★★
 *    A1=1(数) A2="2"(★字★) A3=TRUE(真偽) A4=★空★ A5=2(数) A6==1/0(★誤り★)
 *    B1:B3=★全部 空★ ／ C1=0.1 C2=0.2 C3=-0.3(消え残り) ／ C5=-1 C6=-2(負だけ)
 *
 *  ★★わざと 外した 3通り（★半分 合う 計算を 出さない★）★★
 *    `=SUMPRODUCT(A1:A5,A1:A5)` … ★形の 違う 四角の 決まりが 未測定★
 *    `=A1:A3` `=A1:A3*2`        … ★溢れ＝土台⑤★（実測では 下に 広がりました）
 *    ⇒ 紙（golden）には 40通り 在り、ここは ★37通り★。★この 差は わざと★です。
 */
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const 切 = require_(path.join(ROOT, 'lib/shiki-kiru.js'));
const 形 = require_(path.join(ROOT, 'lib/shiki-katachi.js'));
const K = require_(path.join(ROOT, 'lib/shiki-keisan.js'));
const F = require_(path.join(ROOT, 'lib/shiki-kansuu.js'));

let pass = 0, fail = 0;
const T = (n, f) => {
  try { f(); pass++; console.log('  ✓ ' + n); }
  catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); }
};

/* ══ ★材料（実Excel に 置いたのと 同じ）★ ══ */
const マス = {
  A1: K.数(1), A2: K.字('2'), A3: K.真偽(true), A4: K.空, A5: K.数(2), A6: K.誤('#DIV/0!'),
  B1: K.空, B2: K.空, B3: K.空,
  C1: K.数(0.1), C2: K.数(0.2), C3: K.数(-0.3), C4: K.空, C5: K.数(-1), C6: K.数(-2)
};
const マスか = (s) => /^[A-Z]+[0-9]+$/.test(s) && !/^(TRUE|FALSE)$/.test(s);
const 読む = (名) => Object.prototype.hasOwnProperty.call(マス, 名) ? マス[名] : K.空;

/* ★四角を ほどく★（A1:A5 → A1 A2 A3 A4 A5）＝★1列の 物しか 測って いません★ */
function 四角をほどく(左, 右) {
  const a = /^([A-Z]+)([0-9]+)$/.exec(左), b = /^([A-Z]+)([0-9]+)$/.exec(右);
  if (!a || !b || a[1] !== b[1]) throw new Error('この 測りでは 1列の 四角しか 見ない: ' + 左 + ':' + 右);
  const 並び = [];
  for (let r = Math.min(+a[2], +b[2]); r <= Math.max(+a[2], +b[2]); r++) 並び.push(読む(a[1] + r));
  return 並び;
}

/* ══ ★測る 為の ちいさな 木歩き★ ══ */
function 歩く(木, 手) {
  switch (木.種) {
    case '数': return K.数(K.打った数(木.値));
    case '字': return K.字(String(木.値).slice(1, -1).split('""').join('"'));
    case '誤': return K.誤(木.値);
    case '括': return 歩く(木.子, 手);
    case '前': return K.前置き(木.記, 歩く(木.子, 手), 手);
    case '後': return K.後置き(木.記, 歩く(木.子, 手), 手);
    case '二': {
      if (木.記 === ':') throw new Error('四角は 関数の 引数の 所で ほどきます');
      return K.つなぎ(木.記, 歩く(木.左, 手), 歩く(木.右, 手), 手);
    }
    case '名': {
      const n = String(木.値).toUpperCase();
      if (n === 'TRUE') return K.真偽(true);
      if (n === 'FALSE') return K.真偽(false);
      if (マスか(n)) return 読む(n);
      throw new Error('この 測りでは 見ない 名前: ' + 木.値);
    }
    case '呼': {
      /* ★★引数が どこから 来たかを 落とさない★★＝ここが この 台の 肝 */
      const 引数たち = 木.引数.map((子) => {
        if (子.種 === '二' && 子.記 === ':') {
          return { 種: '四角', 並び: 四角をほどく(String(子.左.値).toUpperCase(), String(子.右.値).toUpperCase()) };
        }
        if (子.種 === '名' && マスか(String(子.値).toUpperCase())) {
          return { 種: 'マス', 値: 読む(String(子.値).toUpperCase()) };
        }
        return { 種: '直', 値: 歩く(子, 手) };
      });
      const 出 = F.呼ぶ(木.名, 引数たち, 手);
      if (出 === null) throw new Error('この 台は 知らない 関数: ' + 木.名);
      return 出;
    }
    default: throw new Error('この 測りでは 見ない 形: ' + 木.種);
  }
}

function 出す(式, 手) {
  const か = 切.切る(式.charAt(0) === '=' ? 式.slice(1) : 式);
  if (!か.ok) throw new Error('切れない … ' + か.なぜ);
  const y = 形.形にする(か.出);
  if (!y.ok) throw new Error('形に ならない … ' + y.なぜ);
  /* ★★本番と 同じ 最後の 一手を 通します★★（2026-09-16）
       `lib/shiki-hyou.js` の `計算する` は ★マスに 入れる 前★に
       `計.足し引きを締める` を 通します（＋－の 消え残りを 消す 所）。
       ★この 見張りは 台を 直に 叩く★ので 通さないと
       `=0.1+0.2-0.3` が 0 に なりません＝★本番より 短い 道を 押す 事に なる★
     ⇒★「誰かから 呼ばれる」と 「本番の 道から 呼ばれる」は 別★ */
  return K.足し引きを締める(歩く(y.形, 手));
}
function 字にして(v) {
  if (v.型 === '誤') return v.値;
  if (v.型 === '真偽') return v.値 ? 'TRUE' : 'FALSE';
  if (v.型 === '数') return K.数を字に(v.値);
  if (v.型 === '空') return '';
  return v.値;
}

/* ══ ★★実Excel の 答え★★ ══ 〔打った 式, 出た 字〕 */
const 実Excel = [
  ['=SUM(A1:A5)', '3'],
  ['=SUM(A1,"2",TRUE)', '4'],
  ['=SUM(A1:A6)', '#DIV/0!'],
  ['=SUM(A2)', '0'],
  ['=SUM(A1:A3,A5)', '3'],
  ['=SUM(B1:B3)', '0'],
  ['=COUNT(A1:A5)', '2'],
  ['=COUNT("1",TRUE)', '2'],
  ['=COUNT(B1:B3)', '0'],
  ['=COUNTA(A1:A5)', '4'],
  ['=AVERAGE(A1:A5)', '1.5'],
  ['=AVERAGE(A1,A3)', '1'],
  ['=AVERAGE(B1:B3)', '#DIV/0!'],
  ['=MAX(A1:A5)', '2'],
  ['=MAX("3",1)', '3'],
  ['=MAX(B1:B3)', '0'],
  ['=MIN(A1:A5)', '1'],
  ['=MIN(A2,1)', '1'],
  ['=PRODUCT(A1:A5)', '2'],
  ['=MIN(B1:B3)', '0'],
  ['=PRODUCT(B1:B3)', '0'],
  ['=COUNTA(B1:B3)', '0'],
  ['=MAX(A1:A6)', '#DIV/0!'],
  ['=COUNT(A1:A6)', '2'],
  ['=COUNTA(A1:A6)', '5'],
  ['=AVERAGE(A1:A6)', '#DIV/0!'],
  ['=SUM("x")', '#VALUE!'],
  ['=SUM(A1:A5,"x")', '#VALUE!'],
  ['=AVERAGE(A2)', '#DIV/0!'],
  ['=MAX(A2)', '0'],
  ['=SUM(0.1,0.2,-0.3)', '0'],
  ['=SUM(C1:C3)', '0'],
  ['=MAX(C5:C6)', '-1'],
  ['=MIN(C5:C6)', '-2'],
  ['=MAX(C5)', '-1'],
  ['=PRODUCT(2,3)', '6'],
  ['=AVERAGE(1,2)', '1.5'],
];

T('★★実Excel の ' + 実Excel.length + '通りと 1つも 違わない★★', () => {
  const 違い = [];
  for (const [式, 欲] of 実Excel) {
    let 得;
    try { 得 = 字にして(出す(式)); }
    catch (e) { 違い.push(式 + ' … ★落ちた★ ' + e.message); continue; }
    if (得 !== 欲) 違い.push(式 + ' … うち `' + 得 + '` ／ ★実Excel `' + 欲 + '`★');
  }
  if (違い.length) {
    throw new Error('★' + 違い.length + ' / ' + 実Excel.length + ' 違う★' + String.fromCharCode(10) + '      ' + 違い.slice(0, 8).join(String.fromCharCode(10) + '      '));
  }
});

T('★★値が どこから 来たかで 扱いが 変わる（この 台の 肝）★★', () => {
  /* ★同じ SUM でも 中身の 扱いが 違う★ */
  const 四角 = 字にして(出す('=SUM(A1:A5)'));       /* 字 "2" と TRUE を 無視 → 3 */
  const 直 = 字にして(出す('=SUM(A1,"2",TRUE)'));   /* 字と 真偽を 数に する → 4 */
  if (四角 === 直) throw new Error('★四角と 直に 書いた 物を 同じに 扱って いる（どちらも ' + 四角 + '）★');
  if (四角 !== '3' || 直 !== '4') throw new Error('★' + 四角 + ' と ' + 直 + '★（実Excel は 3 と 4）');
  /* ★1マス指しも「マスから 来た 物」＝直に 書いたのとは 違う★ */
  if (字にして(出す('=SUM(A2)')) !== '0') throw new Error('★1マス指しの 字を 数に して いる★');
});

T('★★知らない 関数は 断る（半分 合う 計算を 出さない）★★', () => {
  /* ★★ここに `SUMPRODUCT` と 焼き付けて いました★★（2026-09-16 に 外した）
       ★すぐ 下の 覚書きと ★同じ 過ち★です★（VLOOKUP で 1度 踏んで いる）
       ★SUMPRODUCT は 測って 書きました★（棚67）
         `golden-shikaku-kansuu-2026-09-13.tsv` … `=SUMPRODUCT(A1:A5,A1:A5)` → ★ 5★
         （四角の 中の 字・真偽は ★ 0 と して 掛ける★）
       ⇒★名前を 焼き付けず 下の ★台に 無い 名前★だけで 見ます★ */
  /* ★★名前を 焼き付けません★★（2026-09-16）
       ★前は `VLOOKUP` と 書いて いました★ ⇒ ★VLOOKUP を 作った 途端 赤★
       ＝★見張りが「まだ 作って いない」という 思い込みを 守って いた★
       ⇒★★その場で 台に 無い 名前を 選ぶ★★（★これなら 何を 作っても 死にません★） */
  const 無い名 = ['ZZZNOTAFUNCTION', 'QQQNOTAFUNCTION'].find((n) => !F.表[n]);
  if (!無い名) throw new Error('★台に 無い 名前が 選べません★');
  if (F.呼ぶ(無い名, [], {}) !== null) throw new Error('★知らない 関数を 答えて いる★');
  /* ★知って いる 物は ちゃんと 答える★ */
  if (F.呼ぶ('SUM', [{ 種: '直', 値: K.数(1) }], {}).値 !== 1) throw new Error('★SUM が 答えない★');
});

T('★★拾い方が 出どころで 変わって いるか（土台の 肝を 名指しで 見る）★★', () => {
  /* ★四角からは 数だけ 2つ（字 "2" と TRUE と 空は 飛ばす）★ */
  const 元 = F.数を拾う([{ 種: '四角', 並び: [K.数(1), K.字('2'), K.真偽(true), K.空, K.数(2)] }], {});
  if (元.数たち.length !== 2) throw new Error('★四角から 数を 2つ 拾えて いない（' + 元.数たち.length + '個）★');
  /* ★直に 書いた 字は 数に なる★ */
  const 直 = F.数を拾う([{ 種: '直', 値: K.字('2') }], {});
  if (直.数たち.length !== 1 || 直.数たち[0] !== 2) throw new Error('★直に 書いた 字を 数に して いない★');
  /* ★誤りは 伝わる★ */
  const 誤 = F.数を拾う([{ 種: '四角', 並び: [K.数(1), K.誤('#DIV/0!')] }], {});
  if (!誤.誤 || 誤.誤.値 !== '#DIV/0!') throw new Error('★四角の 中の 誤りを 伝えて いない★');
});

console.log('');
console.log('★締め★ 通った ' + pass + ' ／ ★落ちた ' + fail + '★');
process.exit(fail ? 1 : 0);
