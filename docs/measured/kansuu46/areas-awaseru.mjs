/* areas-awaseru.mjs — ★AREAS を ★本番の 道★で 押して 実測と 突き合わせる★（2026-09-08）
 *
 *  ★★物差しは 1本に します★★
 *    2026-09-08 に ★同じ 事を 数える 物を 3つ★ 作って
 *    ★40／29／12 と 3つ 違う 数★が 出ました。
 *    ⇒★訳★ ①部品を 直に 押す ②本番の 道で 押す ③名前の 一覧を 渡す／渡さない
 *      … が 混ざっていた
 *    ⇒★★食い違う 物差しは 1本の 物差しより 悪い★★（2026-09-07 に 決めた）
 *    ⇒ この 1本に まとめました。★本番の 道（convertFormula → engine）で 押します★
 *
 *  ★答え★ 実Excel に 4回 打たせた `golden-areas*-2026-09-07.tsv`
 *
 *  ★★名前（なまえ）の 事★★
 *    ここ（node）では ★名前の 一覧が 無い★ので 名前を 使う 式は 出しません。
 *    ⇒★画面では 名前は ★先に 参照に 開かれる★★（convertFormula の 頭で 名前の箱.開く）
 *    ⇒ だから ★お客さんの 道では ここに 名前は 来ません★
 *    ⇒★部品だけは 名前の 一覧を 渡して 別に 押します★（下の ③）
 *
 *  使い方: node docs/measured/kansuu46/areas-awaseru.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..', '..', '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));

const A = require_(path.join(ROOT, 'lib/formula-areas.js'));
const HFns = require_(path.join(ROOT, 'hyperformula.full.min.js'));
const EF = require_(path.join(ROOT, 'exally-formula.js'));
EF.registerExallyFunctions(HFns);
const hf = HFns.HyperFormula.buildEmpty({ licenseKey: 'gpl-v3' });
const SID = hf.getSheetId(hf.addSheet('Sheet1'));
EF.initExallyFormula(hf);

const 赤の名 = (t) => ({
  NA: '#N/A', DIV_BY_ZERO: '#DIV/0!', VALUE: '#VALUE!', NUM: '#NUM!',
  NAME: '#NAME?', REF: '#REF!', CYCLE: '#CYCLE!', ERROR: '#ERROR!',
}[t] || ('#' + t));

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
  const 後 = EF.convertFormula(式);
  const 表 = 土台();
  表[0][10] = 後;
  hf.setSheetContent(SID, 表);
  const v = hf.getCellValue({ sheet: SID, row: 0, col: 10 });
  return { 値: (v && v.type) ? 赤の名(v.type) : v, 後: 後 };
}

const 金 = [];
for (const f of ['golden-areas-2026-09-07.tsv', 'golden-areas2-2026-09-07.tsv',
  'golden-areas3-2026-09-07.tsv', 'golden-areas4-2026-09-07.tsv']) {
  const t = fs.readFileSync(path.join(ここ, f), 'utf-8');
  for (const l of t.split('\n')) {
    if (!l || l.startsWith('#')) continue;
    const p = l.split('\t');
    if (/^=AREAS\(/.test(p[1] || '')) 金.push({ 式: p[1], 答: p[2] });
  }
}

const 紙 = [];
const 言う = (s) => { 紙.push(s); console.log(s); };
言う('# ★AREAS を 本番の 道で 押して 実測と 突き合わせた★（' + new Date().toISOString().slice(0, 10) + '）');
言う('');
言う('★押し方★ convertFormula → engine（★お客さんと 同じ 道★）');
言う('★実測★ ' + 金.length + '本（4回・Excel 16.0 build 20326）');
言う('');

let 合 = 0, 聞けない = 0;
const まだ = [], 外れ = [];
for (const g of 金) {
  if (/受け取らない|打てない/.test(g.答)) { 聞けない++; continue; }
  const r = 押す(g.式);
  if (String(r.値) === '#NAME?') { まだ.push({ 式: g.式, 答: g.答 }); continue; }
  const よい = /^-?[0-9.]+$/.test(g.答)
    ? Math.abs(Number(r.値) - Number(g.答)) <= 1e-9
    : String(r.値) === g.答;
  if (よい) 合++;
  else 外れ.push(g.式 + '   正 ' + g.答 + ' ／ 出 ' + r.値 + '（' + r.後 + '）');
}

言う('★① 本番の 道で 押した★');
言う('  ★合った ……………… ' + 合 + '本★');
言う('  ★違う ……………… ★' + 外れ.length + '本★');
言う('  ★まだ（#NAME?）… ' + まだ.length + '本★');
言う('  ★実Excel が 式として 受け取らない … ' + 聞けない + '本★');
for (const x of 外れ) 言う('    ' + x);

/* ★★『まだ』の 中身を 1つずつ 読む★★
   ⇒ 2026-09-08 に ★2つの 間違いが「出さない」側に 倒れていた★
     （名前が シートの 数え方を 汚した／シート名の 中の 空白で 割った）
   ⇒★『出さない』に 倒れる 間違いは 赤に ならない＝★中身を 読むしか 無い★★ */
言う('');
言う('★② 「まだ」の 中身（★数だけで 済ませない★）★');
const 分け = { '名前（画面では 先に 開かれる）': [], '関数入り': [], '重ならない 重なり': [],
  'こぼれの 印（A1#）': [], '★その他★': [] };
for (const m of まだ) {
  const 中 = m.式.replace(/^=AREAS\(/, '').replace(/\)$/, '');
  if (/ひとつ|とびとび|みっつ|なまえ/.test(m.式)) 分け['名前（画面では 先に 開かれる）'].push(m);
  else if (m.答 === '#NULL!') 分け['重ならない 重なり'].push(m);
  else if (/#\s*$/.test(中)) 分け['こぼれの 印（A1#）'].push(m);   /* ★A1# … 実Excel も #REF!★ */
  else if (/[()]/.test(中)) 分け['関数入り'].push(m);
  else 分け['★その他★'].push(m);
}
for (const k of Object.keys(分け)) 言う('  ' + k + ' … ' + 分け[k].length + '本');
言う('  ⇒★「その他」が 0本で ない時は ★出せるはずの 物を 出していない★★');
for (const m of 分け['★その他★']) 言う('    ★' + m.式 + '   （実Excel は ' + m.答 + '）★');

/* ★③名前は 部品に 一覧を 渡して 別に 押す★ */
言う('');
言う('★③ 名前（画面では 先に 参照に 開かれる／ここでは 一覧を 渡して 押す）★');
const 名表 = { 'ひとつ': 1, 'とびとび': 2, 'みっつ': 3, 'なまえ': 1 };
let 名合 = 0;
const 名外 = [];
for (const g of 金) {
  if (!/ひとつ|とびとび|みっつ|なまえ/.test(g.式)) continue;
  if (/受け取らない|打てない/.test(g.答)) continue;
  const 中 = g.式.replace(/^=AREAS\(/, '').replace(/\)$/, '');
  const r = A.区画を数える(中, 名表);
  const 出 = (typeof r === 'number') ? String(r)
    : (r && r.誤り) ? ('#' + r.誤り + '!') : '(出さない)';
  if (出 === g.答) 名合++; else 名外.push(g.式 + '   正 ' + g.答 + ' ／ 出 ' + 出);
}
言う('  ★合った … ' + 名合 + '本 ／ 違う … ' + 名外.length + '本★');
for (const x of 名外) 言う('    ' + x);

fs.writeFileSync(path.join(ここ, 'areas-awaseru.txt'), 紙.join('\n') + '\n', { encoding: 'utf-8' });
console.log('\n★書いた … docs/measured/kansuu46/areas-awaseru.txt★');
process.exit((外れ.length || 名外.length || 分け['★その他★'].length) ? 1 : 0);
