/* formula-soto.test.mjs — ★外へ 出る 関数を「作り物の 外」で 押す★（2026-09-07）
 *
 *  ★★外へは 1回も 出しません★★
 *    外と やり取りする 道具を ★作り物★に 差し替えて 押します。
 *    ⇒★お金も かからない／相手にも 迷惑を かけない★
 *    ⇒★でも 通る 道は 本番と 同じ★（JS層 → convertFormula → engine）
 *
 *  ★★ここで 押す 物★★
 *    ①1回目は ★#N/A（取りに 行っています）★
 *    ②届いた 後は ★答えが 出る★
 *    ③★同じ 住所は 1回しか 取りに 行かない★（1000行 在っても 外へは 1回）
 *    ④★お客さんの 打った 字を そのまま 住所に 載せない★（銘柄を 整える）
 *    ⑤届いた 字の 読み取り（CSV → 表／AI の 返事 → 言語コード・訳）
 *
 *  使い方: node tests/formula-soto.test.mjs
 *          node tests/formula-soto.test.mjs --self-test
 */
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..');
const require_ = createRequire(import.meta.url);
const F = require_(path.join(ROOT, 'lib/formula-soto.js'));
const P = require_(path.join(ROOT, 'lib/formula-soto-plug.js'));

const 直に走った = process.argv[1] && pathToFileURL(process.argv[1]).href === import.meta.url;
const 自己試験 = 直に走った && process.argv.includes('--self-test');

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };
const TA = async (n, fn) => { try { await fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };

/* ── 本番と 同じ 物を 積む（外だけ 作り物） ── */
const HFns = require_(path.join(ROOT, 'hyperformula.full.min.js'));
const EF = require_(path.join(ROOT, 'exally-formula.js'));
EF.registerExallyFunctions(HFns);
const HF0 = HFns.HyperFormula;
const H = Object.assign(Object.create(HF0), HFns,
  { registerFunctionPlugin: HF0.registerFunctionPlugin.bind(HF0) });
for (const n of ['extra', 'nokori', 'kane', 'yosoku']) {
  require_(path.join(ROOT, 'lib/formula-' + n + '-plug.js'))
    .つなぐ(H, require_(path.join(ROOT, 'lib/formula-' + n + '.js')));
}

/* ★作り物の 外★＝何を 何回 頼まれたかを 数える */
const 頼まれた = [];
const 返す = new Map();
let 再計算した = 0;
const 知らせ = [];
const 道具 = {
  取る: async (url) => { 頼まれた.push(url); if (!返す.has(url)) throw new Error('作り物に 用意が ない: ' + url); return 返す.get(url); },
  聞く: async (文) => { 頼まれた.push('ai:' + 文.slice(0, 20)); return 返す.get('ai') || ''; },
  再計算: () => { 再計算した++; },
  知らせる: (文) => { 知らせ.push(文); },
};
const 積 = P.つなぐ(H, F, 道具);
/* ══ ★★門が 1つ 増えました＝ここで 明示して 開けます★★ ══（2026-09-26）
     ★★何が 変わったか★★
       前 ･･･ ★開いた だけで 外へ 出て いました★
       今 ･･･ ★お客さんが 「つないで よい」と 言った 本だけ 出します★
       ＝`lib/formula-soto.js` の 頭「もらった `.xlsx` の 式は ★式のまま★ 入る＝★開いた だけで 走る★」
       ＝司さん 09-25「★おすすめで 直せ★」／経営者1 の 推し ⑶
     ★★この 見張りは 「出した 後の 振る舞い」を 見る 物です★★
       ⇒★だから ここで 1回 開けます★（★門が 在る事は 別の 見張りが 見ます★）
       ⇒`tests/soto-wo-kiku-webkit.mjs` ＝★門そのものを 見る 見張り★
     ★開け忘れたら 全部 `#N/A` に なります★＝★だから ここに 書いて おきます★ */
if (typeof P.外へ出す !== 'function') {
  console.log('  ★★測れません★★ `外へ出す` が 在りません＝門の 口が 消えて います');
  process.exit(8);
}
P.外へ出す(true);

const hf = HF0.buildEmpty({ licenseKey: 'gpl-v3', useArrayArithmetic: true, smartRounding: false });
const SID = hf.getSheetId(hf.addSheet('S'));
EF.initExallyFormula(hf);

function 押す(式) {
  const 表 = [[EF.convertFormula(式)]];
  for (let i = 0; i < 30; i++) 表.push([null]);
  hf.setSheetContent(SID, 表);
  const v = hf.getCellValue({ sheet: SID, row: 0, col: 0 });
  return (v && v.type) ? ('#' + v.type) : v;
}
const 待つ = () => new Promise((r) => setTimeout(r, 20));

console.log('\n[formula-soto] 外へ 出る 関数（外は 作り物）');
console.log('  積んだ … ' + 積 + '本');

T('★★積めた★★（積み忘れると 全部 #NAME で 全部 緑に なる）', () => {
  if (!(積 > 0)) throw new Error('積めていない');
});

T('★お客さんの 打った 字を そのまま 住所に 載せない★', () => {
  /* ★見るのは「危ない 字が 落ちたか」★（出来上がりの 字を 丸暗記しない） */
  const 整 = F.銘柄を整える('MS FT/../etc?x=1&y=2 #z');
  for (const 悪 of ['/', '?', '&', '#', ' ', ':']) {
    if (整.indexOf(悪) >= 0) throw new Error('危ない 字が 残った（' + 悪 + '）… ' + 整);
  }
  if (!/^[a-z0-9.\-]+$/.test(整)) throw new Error('通してよい 字だけに なっていない … ' + 整);
  const u = F.株の住所('msft');
  if (!/^https:\/\/stooq\.com\/q\/d\/l\?s=msft\.us&i=d$/.test(u)) throw new Error('住所が 違う … ' + u);
  if (F.株の住所('') !== null) throw new Error('空の 銘柄が 通った');
});

T('★CSV を 表に 直せる★', () => {
  const csv = 'Date,Open,High,Low,Close,Volume\n2024-01-02,10,12,9,11,100\n2024-01-03,11,13,10,12,200';
  const t = F.株の表(csv, 0, 1);
  if (t.length !== 3) throw new Error('行数が 違う … ' + t.length);
  if (t[0][0] !== '日付' || t[0][1] !== '終値') throw new Error('見出しが 違う');
  if (t[1][1] !== 11 || t[2][1] !== 12) throw new Error('終値が 違う');
  if (F.年月日(t[1][0]) !== '2024-01-02') throw new Error('日付が 違う … ' + F.年月日(t[1][0]));
  const t2 = F.株の表(csv, 3, 0);
  if (t2[0].length !== 6) throw new Error('列数が 違う（全部）… ' + t2[0].length);
});

T('★AI の 返事から 言語コードだけ 拾う★', () => {
  if (F.言語コードを拾う('これは ja です') !== 'ja') throw new Error('拾えない');
  if (F.言語コードを拾う('zh-Hans') !== 'zh-Hans') throw new Error('地域つきが 拾えない');
  if (!F.訳を整える('「こんにちは」').startsWith('こんにちは')) throw new Error('引用符が 落ちない');
});

await TA('★★1回目は 取りに 行っています（#N/A）／届いたら 答えが 出る★★', async () => {
  P.覚えを消す();
  返す.set('https://stooq.com/q/d/l?s=msft.us&i=d',
    'Date,Open,High,Low,Close,Volume\n2024-01-02,10,12,9,11,100');
  const 前 = 押す('=INDEX(STOCKHISTORY("MSFT"),2,2)');
  if (前 !== '#NA') throw new Error('1回目が #N/A で ない … ' + 前);
  await 待つ();
  const 後 = 押す('=INDEX(STOCKHISTORY("MSFT"),2,2)');
  if (後 !== 11) throw new Error('届いた 後の 答えが 違う … ' + 後);
});

await TA('★★同じ 住所は 1回しか 取りに 行かない★★（1000行 在っても 外へは 1回）', async () => {
  P.覚えを消す();
  頼まれた.length = 0;
  返す.set('https://stooq.com/q/d/l?s=aapl.us&i=d',
    'Date,Open,High,Low,Close,Volume\n2024-01-02,1,2,3,4,5');
  押す('=INDEX(STOCKHISTORY("AAPL"),2,2)');
  await 待つ();
  for (let i = 0; i < 10; i++) 押す('=INDEX(STOCKHISTORY("AAPL"),2,2)');
  await 待つ();
  const 回 = 頼まれた.filter((u) => u.indexOf('aapl') >= 0).length;
  if (回 !== 1) throw new Error('外へ ' + 回 + '回 出た（1回の はず）');
});

await TA('★WEBSERVICE … 届いた 字が そのまま 出る★', async () => {
  P.覚えを消す();
  返す.set('https://stooq.com/q/d/l?s=x.us&i=d', 'こんにちは');
  const 前 = 押す('=WEBSERVICE("https://stooq.com/q/d/l?s=x.us&i=d")');
  if (前 !== '#NA') throw new Error('1回目が #N/A で ない … ' + 前);
  await 待つ();
  if (押す('=WEBSERVICE("https://stooq.com/q/d/l?s=x.us&i=d")') !== 'こんにちは') throw new Error('字が 出ない');
});

T('★★住所を セルから 作った 式は 断る★★（2026-09-07 指示役の 問いで 足した）', () => {
  /* ★許した 相手だけに しても
     `="https://許した相手/?q=" & A1` と 書けたら ★A1 の 中身は 相手に 届く★
     ⇒★行き先は 安全でも 中身が 出る★
     ⇒★住所は「式の 中に 直に 書いた 字」だけ★に した */
  P.覚えを消す();
  頼まれた.length = 0;
  const 危ない = [
    /* ★式を 置くのは A1★なので、指す 先は 別の セルに する（自分を 指すと #CYCLE） */
    '=WEBSERVICE("https://stooq.com/q/d/l?s="&A5)',
    '=WEBSERVICE(A5)',
    '=WEBSERVICE(CONCAT("https://stooq.com/q/d/l?s=","x"))',
  ];
  for (const f of 危ない) {
    const v = 押す(f);
    if (v !== '#VALUE') throw new Error('断っていない … ' + f + ' → ' + v);
  }
  if (頼まれた.length) throw new Error('★外へ 行った★（' + 頼まれた.length + '回）');
});

T('★★セルの 中身を 銘柄に 載せても 短く 切る★★（少しずつ 運べない）', () => {
  const 長 = F.銘柄を整える('himitsunodetaganagaikoto');
  if (長.length > 12) throw new Error('切れていない … ' + 長);
});

T('★https で ない 住所は その場で 断る★（外へ 行かない）', () => {
  P.覚えを消す();
  頼まれた.length = 0;
  const v = 押す('=WEBSERVICE("http://example.com/")');
  if (v !== '#VALUE') throw new Error('断っていない … ' + v);
  if (頼まれた.length) throw new Error('外へ 行った（' + 頼まれた.length + '回）');
});

await TA('★TRANSLATE / DETECTLANGUAGE … AI の 返事を 整える★', async () => {
  P.覚えを消す();
  返す.set('ai', '「Hello」');
  const 前 = 押す('=TRANSLATE("こんにちは","ja","en")');
  if (前 !== '#NA') throw new Error('1回目が #N/A で ない … ' + 前);
  await 待つ();
  if (押す('=TRANSLATE("こんにちは","ja","en")') !== 'Hello') throw new Error('訳が 整っていない');
  P.覚えを消す();
  返す.set('ai', '言語は ja です');
  押す('=DETECTLANGUAGE("こんにちは")');
  await 待つ();
  if (押す('=DETECTLANGUAGE("こんにちは")') !== 'ja') throw new Error('言語コードが 出ない');
});

await TA('★★お客さんに「取りに 行っています」が 見える★★（#N/A だけだと 壊れたに 見える）', async () => {
  P.覚えを消す();
  知らせ.length = 0;
  返す.set('https://stooq.com/q/d/l?s=nnn.us&i=d', 'Date,Close' + '\n' + '2024-01-02,5');
  押す('=INDEX(STOCKHISTORY("NNN"),2,2)');
  if (!知らせ.length) throw new Error('取りに 行き始めた のに 何も 出ない');
  if (!/取りに 行っています/.test(知らせ[0])) throw new Error('文が 違う … ' + 知らせ[0]);
  await 待つ();
  const 後 = 知らせ[知らせ.length - 1];
  if (!/届きました/.test(後)) throw new Error('届いた 後の 知らせが 出ない … ' + 後);
  /* ★お客さんに 出る 字に ★ を 使わない★（会社の 決まり） */
  for (const 文 of 知らせ) if (文.indexOf('★') >= 0) throw new Error('お客さんの 字に ★ が 混 じっている … ' + 文);
});

await TA('★取れなかった 時も 黙らない★', async () => {
  P.覚えを消す();
  知らせ.length = 0;
  押す('=WEBSERVICE("https://stooq.com/q/d/l?s=yyy.us&i=d")');
  await 待つ();
  const 後 = 知らせ[知らせ.length - 1];
  if (!/取れませんでした/.test(後)) throw new Error('だめだった 事を 言っていない … ' + 後);
});

T('★出す 名簿は lib が 正本★（繋ぐ 側が 別の 名簿を 持っていない）', () => {
  const fs = require_('node:fs');
  const src = fs.readFileSync(path.join(ROOT, 'lib/formula-soto-plug.js'), 'utf-8');
  const 並 = [...src.matchAll(/^\s*'([A-Z][A-Z0-9.]*)':\s*\{\s*method:/gm)].map((m) => m[1]);
  const a = [...F.足した名前()].sort().join(','), b = [...並].sort().join(',');
  if (a !== b) throw new Error('名簿が 違う\n      lib  … ' + a + '\n      plug … ' + b);
});

if (自己試験) {
  console.log('\n[self-test] ★物差しが 効いているか★');
  T('★作り物に 用意が 無い 住所は だめに なる★（黙って 緑に しない）', async () => {
    P.覚えを消す();
    押す('=WEBSERVICE("https://stooq.com/q/d/l?s=zzz.us&i=d")');
  });
  T('★再計算を お願いしている★（届いても 画面が 変わらない を 止める）', () => {
    if (!(再計算した > 0)) throw new Error('1回も 再計算を お願いしていない');
  });
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (直に走った) process.exit(fail ? 1 : 0);
