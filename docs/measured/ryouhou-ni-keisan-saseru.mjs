/* ★両方に 計算させて 突き合わせる★（見張り②／2026-09-14）
 *
 *  ★経営者1 と 決めた 形★
 *    ①★どちらが 答えたかを 数える★ … `lib/shiki-hyou.js` に 入れた（見張り①）
 *    ②★両方が 答えられる 式は 両方に 計算させ、違ったら ★赤★★ … ★これ★
 *       ⇒★実Excel の 紙が 無くても 差が 出る★＝実測が 無い 45個の 穴が 軽く なる
 *       ★借り物を「正」に しません★＝違ったら ★人が 見る★
 *         （借り物にも 誤りは 在る＝`mae-ato` の「こちらだけ誤り 177本」）
 *
 *  ★式の 出どころ★ … `docs` の 紙に 在る 式（★手で 並べて いない★）
 *  ★材料★ … 紙の `#材料<タブ>マス<タブ>値`（★写さない★）
 *  ★借り物は ここでしか 動かしません★＝本番の 道では ない
 *
 *  使い方: node docs/measured/ryouhou-ni-keisan-saseru.mjs [何本まで]
 *          （repo の どこから 呼んでも 動きます＝手元の 絶対パスを 焼き込んで いません）
 */
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
/* ★手元の 絶対パスを 焼き込まない★＝この 紙の 場所から 上へ 2つ が repo の 根 */
const R = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', '..') + path.sep;
const require_ = createRequire(path.join(R, 'package.json'));

const H = require_(path.join(R, 'lib/shiki-hyou.js'));
/* ★材料の 読み方・置き方・「押してよいか」は 1か所★＝`tests/zairyou.mjs`
   （道具ごとに 書くと、材料の 無い 紙で 空の 表を 押して ★偽の 負け★が 出ます） */
const Z = await import('file:///' + path.join(R, 'tests/zairyou.mjs').replace(/\\/g, '/'));

/* ══ 借り物を 立てる（★ここだけ★） ══ */
const ctx = { console, setTimeout, clearTimeout, Date, Math, JSON, Intl, TextEncoder, TextDecoder };
ctx.window = ctx; ctx.self = ctx; ctx.globalThis = ctx;
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(R + 'hyperformula.full.min.js', 'utf8'), ctx, { filename: 'hf.js' });
/* ★名前空間（FunctionPlugin・CellError・ErrorType を 持つ）と 本体（buildFromArray）は 別物★ */
const HFの場 = ctx.HyperFormula;                       /* 名前空間 */
const HF = HFの場.HyperFormula || HFの場;               /* 本体 */

/* ★本番と 同じ 皮を 借り物に 積む★（2026-09-14）
   ＝本番（book.html）は 借り物に `lib/formula-*-plug.js` を 積んで います。
     ★素の 借り物と 比べると うちが 有利に 見えます★ので、★本番の 形★で 比べます。
   ★積めた 数を 必ず 出す★＝黙って 0本の まま 比べない */
const 積んだ = [];
for (const [皮名, 中名] of [
  ['formula-extra-plug.js', 'formula-extra.js'],
  ['formula-nokori-plug.js', 'formula-nokori.js'],
  ['formula-kane-plug.js', 'formula-kane.js'],
  ['formula-yosoku-plug.js', 'formula-yosoku.js'],
  ['formula-filterxml-plug.js', 'formula-filterxml.js'],
  ['formula-cell-plug.js', 'formula-cell.js'],
]) {
  try {
    const P = require_(path.join(R, 'lib/' + 皮名));
    const F = require_(path.join(R, 'lib/' + 中名));
    const n = P.つなぐ(HFの場, F);
    積んだ.push(皮名.replace('formula-', '').replace('-plug.js', '') + ' ' + n + '本');
  } catch (e) {
    積んだ.push('★' + 皮名 + ' が 積めなかった … ' + String(e.message).slice(0, 50) + '★');
  }
}

/* ══ 紙を なめる ══ */
const 判定の字 = /^(合った|違う|こちらだけ誤り|名前が通らない|数えない|同じ|一致)$|直った|変わらず（前から/;
const 紙 = [];
(function なめる(d) {
  for (const e of fs.readdirSync(d, { withFileTypes: true })) {
    const p = path.join(d, e.name);
    if (e.isDirectory()) { なめる(p); continue; }
    if (!/\.(tsv|csv)$/.test(e.name)) continue;
    if (/^cases-/.test(e.name) || /-awase-/.test(e.name)) continue;
    紙.push(p);
  }
})(R + 'docs');

const 仕事 = [];       /* { 式, 材料, 紙 } */
for (const p of 紙) {
  let t;
  try { t = fs.readFileSync(p, 'utf8'); } catch (e) { continue; }
  const 生 = t.split('\n');
  if (生.some((l) => l && !l.startsWith('#') && l.split('\t').some((v) => 判定の字.test((v || '').trim())))) continue;
  const 材料 = Z.材料を読む(p);
  const 見出しの答え = Z.見出しの答えの列(生);
  const 見た = new Set();
  for (const l of 生) {
    if (!l || l.startsWith('#')) continue;
    const c = l.split('\t');
    for (let i = 0; i < c.length; i++) {
      const x = (c[i] || '').trim();
      if (!/^=[A-Z]/.test(x)) continue;
      if (見た.has(x)) break;
      見た.add(x);
      /* ★紙の 答えも 一緒に 持つ★＝★機械が 3通りに 分けられる★（経営者1 の 注文）
         ⇒★どちらが 正しいかを 人が 見る★必要が ほとんど 無くなる */
      const 正 = Z.行の答え(c, i, 見出しの答え);
      仕事.push({ 式: x, 材料, 紙: path.basename(p), 正: 正.答, 正の型: 正.型 });
      break;
    }
  }
}

const 上限 = Number(process.argv[2] || 0) || 仕事.length;
const 選 = 仕事.slice(0, 上限);

console.log('# ★両方に 計算させて 突き合わせる★（2026-09-14）');
console.log('#   式の 出どころ … `docs` の 紙（記録の 紙は 外した）');
console.log('#   ★借り物を「正」に しません★＝違いは ★人が 見る★物として 出します');
console.log('#   ★借り物は「本番の 形」で 立てました★（素では ない）… ' + 積んだ.join(' ／ '));
console.log('#     ＝本番 `book.html` は 借り物に `lib/formula-*-plug.js` を 積んで います。');
console.log('#       素で 比べると ★うちが 有利に 見える★ので 揃えました。');
console.log('#   押す 式 … ★' + 選.length + '通り★（紙 ' + 紙.length + '本から 集めた ' + 仕事.length + '通り）');
console.log('');

/* ★見かけの 差を 作らない★＝2026-09-14 実測。
   ①うちの 真偽は `{型:'真偽',値:true}` ⇒ `String(v.値)` だと `true`／借り物は `TRUE`
     ＝★中身は 同じなのに 違うと 数えて いた★（REGEXTEST 2本）
   ②借り物の 誤りの 名は `NA` `DIV_BY_ZERO` 等 ⇒ `#N/A` `#DIV/0!` に 直してから 比べる */
const 借の誤り = { NA: '#N/A', DIV_BY_ZERO: '#DIV/0!', VALUE: '#VALUE!', REF: '#REF!',
  NAME: '#NAME?', NUM: '#NUM!', NULL: '#NULL!', CYCLE: '#CYCLE!', ERROR: '#ERROR!',
  SPILL: '#SPILL!', CALC: '#CALC!' };
const 字に = (v) => {
  if (v === null || v === undefined) return '';
  if (typeof v === 'object') {
    if (v.溢れ) return '溢れ' + v.行数 + 'x' + v.列数;
    if (v.型 === '誤') return String(v.値);
    if (v.型 === '真偽') return v.値 ? 'TRUE' : 'FALSE';
    if (v.型 !== undefined) return String(v.値);
    if (v.type) return 借の誤り[String(v.type)] || ('#' + v.type);   /* 借り物の CellError */
    return JSON.stringify(v);
  }
  if (typeof v === 'number') return String(v);
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  return String(v);
};
const 数か = (s) => /^-?\d+(\.\d+)?([eE][-+]?\d+)?$/.test(s);
const 近い = (a, b) => {
  const x = Number(a), y = Number(b);
  return isFinite(x) && isFinite(y) && Math.abs(x - y) <= Math.max(1e-9, Math.abs(y) * 1e-9);
};

const 土だけ例=[], 借だけ例=[], 無例=[], 溢れ例=[];
let 溢れ差 = 0, 材料待ち = 0, うち勝ち = 0, 借り勝ち = 0, 両方負け = 0;
const うち勝ち例 = [], 借り勝ち例 = [], 両負け例 = [];
const 借り物だけの関数 = {};
/* ★紙の 答えと 同じか★（数は 1e-9 まで・真偽は 大文字小文字を 見ない） */
const 紙と同じ = (出, 正) => {
  const a = String(出).trim(), b = String(正).trim();
  if (数か(a) && 数か(b)) return 近い(a, b);
  return a.toUpperCase() === b.toUpperCase();
};
let 土台だけ = 0, 借り物だけ = 0, どちらも無 = 0, 両方 = 0, 合 = 0;
const 違い = [];
let 押し数 = 0;
for (const w of 選) {
  /* ★途中で 数を 出す★（指示役1 の 注文）＝★殺された時に どこまで 測ったかが 分かる★
     ★借り物を 式ごとに 立て直す 作り★なので 時間が かかります（400通りで 約1分半）。 */
  if ((++押し数 % 500) === 0) {
    console.log('  … ' + 押し数 + '/' + 選.length
      + '（両方 ' + 両方 + ' ／ 借り物だけ ' + 借り物だけ + ' ／ 材料待ち ' + 材料待ち + '）');
  }
  /* ★押してよいか★＝材料が 無い 紙で ★中身を 見る 式★は 押さない
     ＝2026-09-14、これを 見ずに 押して `=PRODUCT(A1:A5)` が ★偽の 負け★に なった
       （紙は 2／空の 表なら 0＝★どちらも 正しい。条件が 違うだけ★）
     ★どの 関数が 中身を 見るかは `tests/zairyou.mjs` の 表★（CELL は 第1引数で 変わる） */
  const 可 = Z.押してよいか(w.式, w.材料);
  if (!可.よい) { 材料待ち++; continue; }

  /* ── 自前 ── */
  let うち = null, うち字 = '';
  try {
    const h = Z.表に置く(H.表(), w.材料);
    h.打つ('ZZ1', w.式);
    うち = h.値('ZZ1');
    うち字 = 字に(うち);
  } catch (e) { うち字 = '★落ちた★'; }
  const うちが知る = !(うち && うち.型 === '誤' && String(うち.値) === '#NAME?') && うち字 !== '★落ちた★';

  /* ── 借り物 ── */
  let 借字 = '';
  let 借が知る = true;
  try {
    const 板 = [[]];
    const 名から = (k) => { const m = /^([A-Z]+)([0-9]+)$/.exec(k); if (!m) return null; let c = 0; for (const ch of m[1]) c = c * 26 + (ch.charCodeAt(0) - 64); return { r: +m[2] - 1, c: c - 1 }; };
    const 置 = (k, v) => { const a = 名から(k); if (!a) return; while (板.length <= a.r) 板.push([]); while (板[a.r].length <= a.c) 板[a.r].push(null); 板[a.r][a.c] = v; };
    /* ★借り物にも「字は 字」として 渡す★（2026-09-14 実測）
       `buildFromArray` に 文字列 `'2'` を 渡すと ★型 NUMBER・値 2★に なります。
       ★頭に アポストロフィを 付けると 型 STRING・値 "2"★。
       ⇒ 付けないと ★借り物が 数として 計算し、うちだけ 正しく 見える★＝★嘘の 勝ち★
       （2026-09-14、SUM/COUNT/AVERAGE/PRODUCT の 9件が まさに それでした） */
    for (const k in w.材料) {
      const m = w.材料[k];
      置(k, m.型 === '空' ? null
        : m.型 === '字' ? ("'" + m.値)
        : m.型 === '真偽' ? (String(m.値).trim().toUpperCase() === 'TRUE')
        : m.値);
    }
    置('ZZ1', w.式);
    const e = HF.buildFromArray(板, { licenseKey: 'gpl-v3' });
    const a = 名から('ZZ1');
    const v = e.getCellValue({ sheet: 0, row: a.r, col: a.c });
    借字 = 字に(v);
    if (借字 === '#NAME?') 借が知る = false;
    e.destroy();
  } catch (er) { 借字 = '★落ちた★'; 借が知る = false; }

  if (うちが知る && 借が知る) {
    両方++;
    /* ★「うちは 溢れる／借り物は 溢れない」を 負けに 数えない★（2026-09-14）
       ＝借り物（HyperFormula 3.4.0）は ★溢れを 持って いません★。
         `=A1:A3` … うち ★溢れ3x1★ ／ 借り物 ★#VALUE!★
         `=A1:A3*2` … うち ★溢れ3x1★ ／ 借り物 ★0★（★黙って 0★＝実Excel は 溢れ）
       ⇒★実Excel に 近いのは うち★。★これを「違う」に 混ぜると 嘘の 負けに なる★
       ⇒ 別枠で 数え、★左上どうし★を 比べて 中身も 見る */
    if (/^溢れ/.test(うち字) && !/^溢れ/.test(借字)) {
      溢れ差++;
      const 左上 = 字に(うち && うち.並び ? うち.並び[0][0] : null);
      const 同左 = (数か(左上) && 数か(借字)) ? 近い(左上, 借字) : (左上 === 借字);
      if (!同左 && 溢れ例.length < 10) 溢れ例.push(w.式.slice(0, 42) + ' … うち 左上 `' + 左上.slice(0, 18) + '` ／ 借り物 `' + 借字.slice(0, 18) + '`');
      continue;
    }
    const 同 = (数か(うち字) && 数か(借字)) ? 近い(うち字, 借字) : (うち字 === 借字);
    if (同) {
      合++;
      /* ★両方 同じでも 紙と 違えば「両方 負け」★（★揃って 間違って いる★） */
      if (Z.真値か(w.正) && !紙と同じ(うち字, w.正)) {
        両方負け++;
        両負け例.push('[' + w.紙 + '] ' + w.式.slice(0, 42) + ' … ★紙 `' + String(w.正).slice(0, 18) + '`★ ／ うち `' + うち字.slice(0, 18) + '` ／ 借り物 `' + 借字.slice(0, 18) + '`');
      }
      continue;
    }
    else if (Z.真値か(w.正)) {
      /* ★紙が 在る＝機械が 分けられる★ */
      const う = 紙と同じ(うち字, w.正), か = 紙と同じ(借字, w.正);
      const 線 = '[' + w.紙 + '] ' + w.式.slice(0, 42) + ' … ★紙 `' + String(w.正).slice(0, 18) + '`★ ／ うち `' + うち字.slice(0, 18) + '` ／ 借り物 `' + 借字.slice(0, 18) + '`';
      if (う && !か) { うち勝ち++; うち勝ち例.push(線); }
      else if (!う && か) { 借り勝ち++; 借り勝ち例.push(線); }
      else { 両方負け++; 両負け例.push(線); }
      continue;
    }
    else 違い.push('[' + w.紙 + '｜材料 ' + (Object.keys(w.材料).length ? Object.keys(w.材料).length + 'マス' : '★無し★') + '] ' + w.式.slice(0, 44) + '\n      うち `' + うち字.slice(0, 30) + '` ／ 借り物 `' + 借字.slice(0, 30) + '`');
  } else if (うちが知る) { 土台だけ++; if (土だけ例.length < 6) 土だけ例.push(w.式.slice(0,40)+' … うち `'+うち字.slice(0,20)+'`'); }
  else if (借が知る) {
    借り物だけ++;
    /* ★★どの 関数が 4,645通りを 作って いるか★★（2026-09-15・指示役1 の 注文）
       ＝★「外して 悪く なる」の 本体は この 数★なので、★関数ごとに 割る★。
         ⇒ ★使われない 関数が 大半なら 外せる／よく 使う 物が 混じるなら まだ★ */
    const 外 = Z.外の関数(w.式) || '（関数では ない）';
    借り物だけの関数[外] = (借り物だけの関数[外] || 0) + 1;
    if (借だけ例.length < 8) 借だけ例.push(w.式.slice(0,40)+' … 借り物 `'+借字.slice(0,20)+'`');
  }
  else { どちらも無++; if (無例.length < 10) 無例.push(w.式.slice(0,40)+' … うち `'+うち字.slice(0,18)+'` ／ 借り物 `'+借字.slice(0,18)+'`'); }
}

const 違数 = 違い.length;
console.log('★締め★');
console.log('  ★両方が 答えた ……… ' + 両方 + '通り★');
console.log('      同じ ' + 合 + ' ／ ★違う ' + (両方 - 合 - 溢れ差) + '★');
console.log('      ★うちは 溢れる／借り物は 溢れない … ' + 溢れ差 + '★（★借り物に 溢れは 在りません★＝負けでは ない）');
console.log('  ★自前だけが 答えた … ' + 土台だけ + '通り★（借り物に 無い 物を 足した 分）');
console.log('  ★★借り物だけが 答えた … ' + 借り物だけ + '通り★★ ＝★これが 借り物を 消せない 理由★');
console.log('  どちらも 答えない … ' + どちらも無 + '通り');
console.log('  ★紙に 材料が 無くて 押さなかった … ' + 材料待ち + '通り★（★偽の 負けを 作らない★）');
console.log('');
console.log('★★紙（実Excel）と 突き合わせて 3通りに 分けた★★（★人が 見るのは 両方負けだけ★）');
console.log('  ★うちの 勝ち ………… ' + うち勝ち + '通り★（借り物を 外すと ★良く なる★）');
console.log('  ★★借り物の 勝ち …… ' + 借り勝ち + '通り★★（★借り物を 外すと 悪く なる＝これだけが 判断に 効く★）');
console.log('  ★両方 負け ………… ' + 両方負け + '通り★（借り物を 外しても 変わらない＝今も 間違って いる）');
console.log('');
console.log('  ★どちらが 答えたか（土台の 数え）★ ' + JSON.stringify(H.答えた数()));
if (違数) {
  console.log('');
  console.log('★両方 答えたのに 違う（★全部★）★＝★どちらが 正しいかは 人が 見る★');
  違い.filter(Boolean).forEach((x) => console.log('  ' + x));
}
console.log('');
console.log('★自前だけが 答えた（例）★'); 土だけ例.forEach((x)=>console.log('  '+x));
console.log('');
console.log('★借り物だけが 答えた（例）★'); 借だけ例.forEach((x)=>console.log('  '+x));
console.log('');
console.log('★どちらも 答えない（例）★'); 無例.forEach((x)=>console.log('  '+x));
if (溢れ例.length) {
  console.log('');
  console.log('★溢れの 左上どうしも 違う 物★＝★ここは 見る 値打ちが 在ります★');
  溢れ例.forEach((x) => console.log('  ' + x));
}
const 並べる = (名, 一覧) => { if (!一覧.length) return; console.log(''); console.log(名); 一覧.forEach((x) => console.log('  ' + x)); };
並べる('★うちの 勝ち（全部）★', うち勝ち例);
並べる('★★借り物の 勝ち（全部）★★＝★借り物を 外すと 悪く なる 行★', 借り勝ち例);
並べる('★両方 負け（全部）★＝★今も 間違って いる／借り物を 外しても 変わらない★', 両負け例);

/* ★★借り物だけが 答える 4,645通りを 関数ごとに 割る★★（外す 代金の 本体） */
{
  const 並 = Object.keys(借り物だけの関数).sort((a, b) => 借り物だけの関数[b] - 借り物だけの関数[a]);
  console.log('');
  console.log('★★借り物だけが 答える ' + 借り物だけ + '通り＝関数ごと★★（' + 並.length + '個の 関数）');
  console.log('  ★これが「外して 悪く なる」の 本体★＝外した 瞬間に #NAME? に なる');
  let 積 = 0;
  並.forEach((k, i) => {
    積 += 借り物だけの関数[k];
    if (i < 40) console.log('  ' + String(借り物だけの関数[k]).padStart(5) + '本  ' + k
      + '   （ここまで ' + 積 + '本／' + Math.round(積 / 借り物だけ * 100) + '%）');
  });
  if (並.length > 40) console.log('  … 残り ' + (並.length - 40) + '個');
}
