/* osu-86-karimono-no-ima.mjs — ★★借り物が 建って いる ★今★ の 答えを 紙に 取る★★（2026-09-15）
 *
 *  ★★なぜ 今 取るのか（★外した 後では 取れません★）★★
 *    自前の 関数 ★86個★が ★借り物（HyperFormula）の 骨組みの 上に 載って います★
 *      ・★足す 65個★ … `lib/formula-*-plug.js` が 借り物に 登録する
 *      ・★包む 21個★ … `formula-complex-plug` が IM系の ★出す 字だけ★ 実Excel に 合わせる
 *    ★司さんの 実物 1冊は この 86個を ★1つも★ 使いません★（実測・2026-09-15）
 *    ⇒★★だから 外す時に 86個が 黙って 変わっても この 1冊では 気づけません★★
 *    ⇒★★借り物が 建って いる 今 の 答えを 焼いて おく★★
 *      ＝★外した 後に「外す前と 同じか」が 測れる★
 *
 *  ★★この 紙で 言える 事／言えない 事★★
 *    ★言える★ … ★★外す前と 同じか★★
 *    ★★言えない★★ … ★正しいか★
 *      ＝★これは「うちの 今の 答え」であって「実Excel の 答え」では ありません★
 *      ★実Excel の 答えが 在る 分★は `docs/measured/kansuu46/` の 紙が 持って います
 *      （2026-09-15 実測＝★86個の うち 66個は 紙が 在る／20個は 無い★）
 *
 *  ★★外へ 出る 4個は 答えを 焼きません★★（指示役1 の 決め・2026-09-15）
 *    WEBSERVICE ／ STOCKHISTORY ／ TRANSLATE ／ DETECTLANGUAGE
 *    ★毎回 変わる★ので ★「外す前と 同じか」も 測れません★
 *    ⇒★本番の 道は 元から 外へ 出させません★（`取る`／`聞く` が 投げる）
 *      ＝★出るのは「外へ 出ようとした」という 印だけ★＝★呼び方だけ★
 *    ⇒★お金・秘密・相手の 迷惑の 3つとも 避けられます★
 *
 *  ★★通り数を 必ず 出します★★（指示役1 の 注文）
 *    ★1通りしか 押せて いない 関数は「守れて いる」に 近くない★
 *    ⇒★1通り ◯個／2〜4通り ◯個／5通り以上 ◯個★を 出します
 *
 *  ★★この 台の 数は 画面の 数では ありません★★
 *    ★`setSheetContent` で 板ごと 入れて います★（本番は 1マスずつ）。
 *    ⇒ 2026-09-10 … 板ごと 入れた せいで 裸の `=LINEST(…)` が #VALUE! に なり
 *      ★「本番が 壊れて いる」と 報告する 一歩 手前★まで 行きました。
 *    ⇒★画面の 事を 言いたい なら ブラウザで 押して ください★
 *
 *  ★★見て いない 事★★
 *    ・★引数の 組は 17通りの 決め打ち★＝★その 関数に とって 意味の 在る 組とは 限りません★
 *      （★だから 通り数を 出します★＝★1通りは 弱い★）
 *    ・★答えが 正しいかは 見て いません★（上に 書いた 通り）
 *
 *  使い方: node docs/measured/osu-86-karimono-no-ima.mjs
 *  出す先: docs/measured/golden-86-karimono-2026-09-15.tsv
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const 出す先 = path.join(ROOT, 'docs/measured/golden-86-karimono-2026-09-15.tsv');

const { 建てる } = await import(pathToFileURL(path.join(ROOT, 'docs/measured/honban-no-michi.mjs')).href);
const { 指紋の行 } = await import(pathToFileURL(path.join(ROOT, 'docs/measured/zairyou-no-yubimon.mjs')).href);
const 土台 = await import(pathToFileURL(path.join(ROOT, 'docs/measured/kansuu46-no-dodai.mjs')).href);
const 道 = await 建てる({ 板の名: 'Sheet1' });   /* ★紙を 取った 時と 同じ 板の 名前★ */
const 台 = 道;
const { EF, hf, SID } = 道;
console.log('★本番と 同じ ' + 道.積んだ + '本の プラグインを つないだ★');

/* ══ ★相手の 86個を 名簿から 集める★（★字で 探さない★＝2026-09-15 に 踏んだ） ══ */
const 札 = new Set();
for (const f of fs.readdirSync(path.join(ROOT, 'lib')).filter((x) => /^formula-.*-plug\.js$/.test(x))) {
  const s = fs.readFileSync(path.join(ROOT, 'lib', f), 'utf8');
  for (const m of s.matchAll(/['"]([A-Z][A-Z0-9_.]{1,30})['"]\s*:/g)) 札.add(m[1]);
  for (const m of s.matchAll(/\b([A-Z][A-Z0-9_.]{1,30})\s*:\s*(function|\()/g)) 札.add(m[1]);
}
const CX = require_(path.join(ROOT, 'lib/formula-complex.js'));
const 包む = new Set(CX['文字を返す']);
const 相手 = [...new Set([...札, ...包む])].sort();

/* ★★答えを 焼かない 物★★＝★毎回 変わるから です★
     ⇒★焼いても「外す前と 同じか」が 測れません★（★毎回 赤に なる★）
     ㋐★外へ 出る 4個★ … 網の 向こうが 変わる
     ㋑★★乱数★★ … ★RANDARRAY★
        ★見つけ方★＝★思い込みで 名前を 並べず 2回 走らせて 突き合わせた★
        （2026-09-15：1回目と 2回目で ★RANDARRAY の 4行だけ 違った★）
        ★これを 見逃すと どう なるか★
          ★紙が 毎回 違う★⇒★「外した せい」と 読み違える★
          ＝★今日 XIRR で やった 「直りか 壊れか」の 逆★
     ★焼く代わりに 出す 物★ … ★形だけ★（誤りか／数か／字か） */
const 外へ = new Set(['WEBSERVICE', 'STOCKHISTORY', 'TRANSLATE', 'DETECTLANGUAGE']);
const 毎回変わる = new Set(['RANDARRAY']);

/* ★★環境に 依る 2個★★（指示役1 の 断り・2026-09-15）
     ★CELL ／ INFO★ … ★書いて いる 本・開いて いる 板・機械の 事★を 返す
     ⇒★裸の 台では ★誤りが 正しい★ 事も 在る★
     ⇒★★実 Excel と 違っても「合わない」と 数えない★★
       （★直そうと して 嘘を 作る★のを 避ける） */
const 環境に依る = new Set(['CELL', 'INFO']);

/* ★★お金の 関数に 日付を 渡すと #VALUE!★★（★製品の 欠陥・棚㉘★・2026-09-15）
     ★実測★
       `=ACCRINT(39508,39691,39569,0.1,1000,2,0)`                  … ★16.666666666666664★
          ＝★実 Excel の 答えと 1字も 違いません★
       `=ACCRINT(DATE(2008,3,1),DATE(2008,8,31),…)`               … ★#VALUE!★
       `=ACCRINT(D1,D2,D1,…)`（マスで 渡す）                   … ★#VALUE!★
       `=DATE(2008,3,1)` 単体                                    … 39508（★DATE は 正しい★）
     ⇒★★中身は 合って いて 入口で 落ちて います★★
     ⇒★お客さんは 必ず `DATE()` か マスで 渡します★
        （`=ACCRINT(39508,…)` と 打つ 人は 居ません）
     ⇒★★実質 使えません★★
   ★ここに 名前を 書く 訳★
     ★「合わない」と「なぜ 合わないか」を ★同じ 行に 置く★★
     ＝★外した 後に 誰かが「外したせいだ」と 読むのを 止める★ */
const 日付で落ちる = new Set([
  'ACCRINT', 'ACCRINTM', 'AMORDEGRC', 'AMORLINC', 'COUPDAYBS', 'COUPDAYS', 'COUPDAYSNC',
  'COUPNCD', 'COUPNUM', 'COUPPCD', 'DISC', 'DURATION', 'INTRATE', 'MDURATION', 'PRICE',
  'PRICEDISC', 'PRICEMAT', 'RECEIVED', 'YIELD', 'YIELDDISC', 'YIELDMAT',
]);

/* ★★大きすぎて 押せない 式★★（★名指し＋実測つき★・2026-09-15）
     ★なぜ 上限を 上げないか★
       ★落ちた／遅い／重い は「邪魔」では なく ★合図★★
       ★上限を 上げて 通すのは ★合図を 消す 手★★
       （2026-09-15：上限を 上げて いたら ★#CYCLE だらけの 紙★を 通して いました）
     ★実測（上限 512MB／1本ずつ／押す たびに 板を 空に）★
       `=WRAPROWS(A1:A5,3)`      … ★ 1 MB★  → 1
       `=WRAPROWS(A1:A5,1.5)`    … ★ 0 MB★  → 1
       `=WRAPROWS(A1:A5,-1)`     … ★ 0 MB★  → #NUM
       ★`=WRAPROWS(A1:A5,45294)`  … ★37 MB★ → 1★
       ★`=WRAPCOLS(A1:A5,45294)`  … ★落ちる★（512MB で 死ぬ）★
     ⇒★★漏れでは ありません★★＝★45,294 幅の 表を 本当に 作る★から
       （★最初 私は「WRAPROWS が 漏れて いる」と 見ましたが
         ★単体で 10回 押しても 2MB★でした＝★前の 見立ては 間違い★）
     ★だから ★押さずに 紙に 書きます★★＝★黙って 落とさない★
     ★見て いない 事★ … ★この 2本の 答えは 外す 前も 後も 比べられません★ */
const 大きすぎる = new Set([
  '=WRAPROWS(A1:A5,45294)',
  '=WRAPCOLS(A1:A5,45294)',
]);

/* ★★`kansuu46` の 紙から 式と 実 Excel の 答えを 読む★★（2026-09-15）
     形 … `関数 式 実 Excel の 答え 型`（tab 区切り／`#` の 行は 覚書き）
     ★なぜ こちらを 先に 使うか★
       17通りの 決め打ちだと ★お金の 18個が 誤りしか 返さなかった★
       （★日付も 利率も 渡して いなかった★）…実測 2026-09-15
       ⇒★数えたら その 18個は ★全部★ `kansuu46` が 覚えて いた★
     ★こちらなら 実 Excel の 答えまで 焼けます★
       ＝★「外す前と 同じか」だけで なく「正しいか」も 測れる★ */
const 紙の組 = new Map();
{
  const D = path.join(ROOT, 'docs/measured/kansuu46');
  for (const f of fs.readdirSync(D)) {
    if (!/^golden-.*\.tsv$/.test(f)) continue;
    for (const 行 of fs.readFileSync(path.join(D, f), 'utf8').split('\n')) {
      if (!行 || 行.charAt(0) === '#') continue;
      const c = 行.split('\t');
      if (c.length < 2) continue;
      if (!/^[A-Z][A-Z0-9_.]*$/.test(c[0])) continue;
      if (c[1].charAt(0) !== '=') continue;
      if (!紙の組.has(c[0])) 紙の組.set(c[0], []);
      紙の組.get(c[0]).push({ 式: c[1], 正: (c[2] === undefined ? '' : c[2]).trim(), 紙: f });
    }
  }
}

/* ★実 Excel の 答えと 合うか★
     ★字で 見ます★＝★「中の 数が 同じ」は「同じ」では ない★
     ただし ★数は 15桁で 丸めてから 見ます★（実 Excel が そう 見せる） */
function 実と合うか(出, 正) {
  const a = String(出).trim(), b = String(正).trim();
  if (a === b) return true;
  const x = Number(a), y = Number(b);
  if (!Number.isFinite(x) || !Number.isFinite(y)) return false;
  if (x === y) return true;
  const j = (v) => Number(v.toPrecision(15));
  return j(x) === j(y);
}

/* ★紙が 無い 関数だけ この 17通り★（`osu-wakeru.mjs` と 同じ＝★別の 組を 作らない★） */
const 候補 = ['(1)', '()', '(A1:A2,1)', '(1,1)', '(A1:A2)', '(1,1,1)', '("a")', '(A1)', '(1,1,1,1)',
  '(A1:B2,1,A1:A2)', '(A1:A2,A1:A2)',
  '(A1:A2,LAMBDA(v,v*2))', '(0,A1:A2,LAMBDA(a,b,a+b))', '(2,2,LAMBDA(r,c,r*c))',
  '(x,2,x*3)', '(x,x+1)', '("<a><b>1</b></a>","//b")'];

/** ★1つの 式を 本番の 道で 押す★（JS層 → convertFormula → エンジン） */
/* ★★材料と 置き場と 押し方は 共通の 本★★（2026-09-15）
     `docs/measured/kansuu46-no-dodai.mjs`
   ★★なぜ 共通に したか（★私の 失敗★）★★
     ★最初 私は 材料を ★自分で 決めました★（A1:B5 = 1,3／2,4／…）
     ★紙の 材料は A1:A5 = 1,2,3,4,5★
     ⇒ `=IMSUM(A1:A5)` … ★うち 24 ／ 実 Excel 15★
     ⇒★★928本が「合わない」と 出た★★＝★★偽の 負け★★
     ⇒★「うちは 半分 負けて いる」と 報せる 一歩 手前★
     ★決まりは 前から 在りました★
       `feedback_kami_no_zairyou_ga_bunshou_dake_nara_nise_no_make`
       ＝★紙の 材料は 紙から 取る／道具が 勝手に 決めない★
     ⇒★★読んで いたのに 自分で 作りました★★
       ＝★一番 読まれないのは 自分が 書いた 紙★ */


/* ══ ★押す★ ══ */
/* ★★紙の 頭に「出どころ」を 書く★★（指示役1 の 決め・2026-09-15）
     ★書かないと「古い うちの 答え」を「正」として 突き合わせます★
     ＝★偽の 勝ちの 一番 静かな 形★＝★自分の 古い 答えと 自分を 比べて 100%★ */
const 行 = [
  '# ★★出どころは 行ごと★★ … ★下の 「出どころ」の 列を 見て ください★',
  '#   ★実 Excel★ … `kansuu46` の 紙の 組で 押した 行★★正しいかまで 言えます★★',
  '#   ★借り物★   … 紙が 無いので 17通りの 決め打ちで 押した 行'
    + '★言えるのは「外す前と 同じか」だけ★',
  '#   ★形だけ★   … ★毎回 変わる★ので 答えを 焼いて いない 行（乱数／外へ 出る）',
  '#   ★実 Excel(環境)★ … CELL／INFO★裸の 台では 誤りが 正しい 事も 在る★'
    + '⇒★違っても「合わない」と 数えて いません★',
  '# ★取った 日★ … 2026-09-15 ／ ★取った 道★ … docs/measured/osu-86-karimono-no-ima.mjs',
  '# ★使い方★ … ★外した 後に これと 突き合わせて「外す前と 同じか」を 見る★',
  '# ★言えない 事★ … ★正しいか★（実 Excel の 答えは docs/measured/kansuu46/ が 持つ・86個中 66個）',
  ...指紋の行(require_, import.meta.url),
  '関数\t式または引数の組\t★出どころ★\t道\t出た字\t実 Excel の 答え\t判\t紙',
];
const 通り数 = {};
const 誤りでない数 = {};
const 紙で押した = new Map();
let 押した = 0;
let 大きすぎて押さず = 0;
for (const f of 相手) {
  土台.板を空に(台);                 /* ★前の 溢れを 残さない★ */
  if (外へ.has(f)) {
    /* ★答えは 焼かない★＝★呼び方だけ★（本番の 道は 外へ 出させない） */
    const r = 土台.押す(台, '=' + f + '("x")');
    行.push(f + '\t★外へ 出る★\t★形だけ★\t' + r.道 + '\t★答えは 焼かない（呼び方だけ）＝' + r.字 + '★\t\t\t');
    通り数[f] = 0;
    continue;
  }
  if (毎回変わる.has(f)) {
    /* ★乱数★…★答えは 焼かず 形だけ 焼く★（★毎回 違う 物を 焼くと 紙が 使えない★） */
    let 数 = 0, 誤 = 0;
    for (const a of 候補) {
      const r = 土台.押す(台, '=' + f + a);
      if (r.字 === '★投げた★') continue;
      if (/^#[A-Z/0-9!?.]+$/.test(r.字.trim())) 誤++; else 数++;
    }
    行.push(f + '\t★毎回 変わる（乱数）★\t★形だけ★\t形だけ\t'
      + '★誤りでない ' + 数 + '通り／誤り ' + 誤 + '通り★（★答えは 焼かない★）\t\t\t');
    通り数[f] = 数 + 誤;
    誤りでない数[f] = 数;
    continue;
  }
  /* ★★紙の 組が 在れば そちら★★（★実 Excel の 正まで 付く★） */
  const 組 = 紙の組.get(f);
  if (組 && 組.length) {
    const 環 = 環境に依る.has(f);
    let m = 0, 中m = 0, 合 = 0, 否 = 0, 正無し = 0;
    for (const k of 組) {
      if (大きすぎる.has(k.式.replace(/\s+/g, ''))) {
        /* ★大きすぎて 押せない★＝★黙って 落とさず 紙に 書く★ */
        行.push(f + '\t' + k.式 + '\t★押して いない★\t—\t'
          + '★大きすぎる（45,294幅の 表を 作る）★\t' + k.正
          + '\t★押して いない★\t' + k.紙);
        大きすぎて押さず++;
        continue;
      }
      const r = 土台.押す(台, k.式);
      土台.板を空に(台);               /* ★押す たびに 空に★＝★溢れが 積もらない★ */
      if (r.字 === '★投げた★') continue;
      m++;
      if (!/^#[A-Z/0-9!?.]+$/.test(r.字.trim())) 中m++;
      let 判;
      if (k.正 === '') { 判 = '★紙に 正が 無い★'; 正無し++; }
      else if (実と合うか(r.字, k.正)) { 判 = '合う'; 合++; }
      else if (環) { 判 = '★環境に 依る＝数えない★'; }
      else {
        判 = 日付で落ちる.has(f)
          ? '★合わない（お金の 関数に 日付を 渡すと #VALUE!・製品の 欠陥・棚㉘）★'
          : '★合わない★';
        否++;
      }
      行.push(f + '\t' + k.式 + '\t' + (環 ? '★実 Excel(環境)★' : '★実 Excel★')
        + '\t' + r.道 + '\t' + r.字 + '\t' + k.正 + '\t' + 判 + '\t' + k.紙);
      押した++;
    }
    通り数[f] = m;
    誤りでない数[f] = 中m;
    紙で押した.set(f, { m, 合, 否, 正無し, 環 });
    continue;
  }
  let n = 0, 誤りでない = 0;
  for (const a of 候補) {
    const 式 = '=' + f + a;
    const r = 土台.押す(台, 式);
    if (r.字 === '★投げた★') continue;      /* ★投げた 組は 通り数に 入れない★ */
    n++;
    /* ★★誤りも「答え」に 数えたら ★全部 5通り以上★ に なりました★★（2026-09-15）
       ⇒★それでは「どれだけ 覚えて いるか」の 測りに なりません★
       ⇒★★誤りでない 答えを 別に 数えます★★（★分母を 2つ 出す★）
       ★誤りも 焼きます★＝★外した 後に 誤りの 出方が 変わったら それも 変化★ */
    if (!/^#[A-Z/0-9!?.]+$/.test(r.字.trim())) 誤りでない++;
    押した++;
    行.push(f + '\t' + a + '\t★借り物★\t' + r.道 + '\t' + r.字 + '\t\t\t');
  }
  通り数[f] = n;
  誤りでない数[f] = 誤りでない;
}

fs.writeFileSync(出す先, 行.join('\n') + '\n', 'utf8');

/* ══ ★数を 出す★ ══ */
function 段に分ける(表) {
  const 段 = { '★0通り★': 0, '1通り': 0, '2〜4通り': 0, '5通り以上': 0 };
  const 零 = [];
  for (const f of 相手) {
    if (外へ.has(f)) continue;
    const n = 表[f];
    if (n === 0) { 段['★0通り★']++; 零.push(f); }
    else if (n === 1) 段['1通り']++;
    else if (n <= 4) 段['2〜4通り']++;
    else 段['5通り以上']++;
  }
  return { 段, 零 };
}
const 全部 = 段に分ける(通り数);
const 中身 = 段に分ける(誤りでない数);
console.log('\n# ★借り物が 建って いる 今 の 答えを 焼いた★（2026-09-15）');
console.log('#   ★★出どころは 行ごと★★（実 Excel ／ 借り物 ／ 形だけ）\n');
console.log('★相手★ … ' + 相手.length + '個（足す ' + 札.size + ' ＋ 包む ' + 包む.size + '）');
console.log('★外へ 出る（答えを 焼かない）★ … ' + [...外へ].join(' '));
console.log('★焼いた 行★ … ' + 押した);
console.log('★★大きすぎて 押して いない 式★★ … ' + 大きすぎて押さず
  + '（★名指しで 紙に 書いて あります★＝★黙って 落として いません★）');
console.log('★出す先★ … ' + path.relative(ROOT, 出す先).replace(/\\/g, '/'));
console.log('\n★★通り数★★（★分母は 引数の 組 ' + 候補.length + '通り★）');
console.log('【①何か 返った】（★誤りも 数える★）');
for (const [k, v] of Object.entries(全部.段)) console.log('   ' + String(v).padStart(4) + '個  ' + k);
console.log('【★★②誤りでない 答えが 返った★★】（★こちらが 本当の 覚え具合★）');
for (const [k, v] of Object.entries(中身.段)) console.log('   ' + String(v).padStart(4) + '個  ' + k);
console.log('★★①だけ 見ると 全部 5通り以上に 見えます★★＝★分母を 出さない 緑は 嘘★');
/* ══ ★★実 Excel の 正が ★実際に★ 付いた のは 何個か★★ ══
     ★「覚えて いる はず★ と ★付いた★ は 別★（指示役1 の 断り・2026-09-15）
       紙の 覚え … ★66個★（名前が 出るか だけ 見た）
       ⇒★実際に 押せた 数を 出します★ */
{
  const 付いた = [...紙で押した.keys()].sort();
  let 合 = 0, 否 = 0, 正無し = 0, 行 = 0;
  const 合わない関数 = [];
  for (const f of 付いた) {
    const x = 紙で押した.get(f);
    合 += x.合; 否 += x.否; 正無し += x.正無し; 行 += x.m;
    if (x.否) 合わない関数.push(f + '(' + x.否 + ')');
  }
  console.log('\n★★実 Excel の 紙で 押した 分★★');
  /* ★★何を 何で 数えたか を 数の 隣に 書く★★（2026-09-15）
     ★前は「付いた 65個」とだけ 出して いました★
     ⇒★指示役1 が「66 と 65 の 差は 何か」と 聞いた★
     ⇒★差＝`RANDARRAY`（乱数なので 答えを 焼かず 形だけ）★
     ⇒★★どちらも 嘘では ない／何を 数えたかが 違うだけ★★
     ⇒★★数の 隣に 訳を 置きます★★ */
  const 乱数で除いた = [...相手].filter((f) => 毎回変わる.has(f) && 紙の組.has(f));
  console.log('   ★実Excel の 紙が 在る 関数★ … ' + (付いた.length + 乱数で除いた.length) + '個'
    + '（★相手 ' + 相手.length + '個の うち★）');
  console.log('   ★そのうち 乱数なので 形だけ 焼いた★ … ' + 乱数で除いた.length + '個'
    + (乱数で除いた.length ? '（' + 乱数で除いた.join(' ') + '）' : ''));
  console.log('   ★★⇒ 実Excel と 突き合わせた★★ … ' + 付いた.length + '個');
  console.log('   ★押した 行★ … ' + 行);
  console.log('   ★★合った ' + 合 + ' ／ 合わない ' + 否 + ' ／ 紙に 正が 無い ' + 正無し + '★★');
  if (合わない関数.length) {
    console.log('   ★★合わない 関数（名指し・行数）★★');
    for (let i = 0; i < 合わない関数.length; i += 6) {
      console.log('     ' + 合わない関数.slice(i, i + 6).join('  '));
    }
    console.log('   ⇒★★これは 「外す前から 合って いない」★★');
    console.log('     ＝★外した 後の 赤と 取り違えない 為 ここに 書いて います★');
  } else {
    console.log('   ★合わない 関数 … 0個★');
  }
  console.log('   ★CELL／INFOは 「環境に 依る」ので 合わないに 数えて いません★');
}

if (中身.零.length) {
  console.log('\n★★誤りしか 返らなかった 関数★★（★外す 時に 比べる 物が 誤りだけ★）');
  for (let i = 0; i < 中身.零.length; i += 8) console.log('  ' + 中身.零.slice(i, i + 8).join(' '));
}
