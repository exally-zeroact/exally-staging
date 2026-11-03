/* osu-jitsubutsu-mitame.mjs — ★差し替える 前に「何マス 変わるか」を 数える★（2026-09-15）
 *
 *  ★★なぜ★★（指示役1 2026-09-15）
 *    `js/book-open.js:378` の `XLSX.SSF.format(...)`（借り物 SheetJS）を
 *    ★自前の 台（`lib/shoshiki.js`）に 差し替える★前に ★何マス 変わるか★を 知る。
 *    ⇒★壊してから 数えない★
 *
 *  ★★実Excel は 要りません★★
 *    ★くらべる 相手は どちらも JS★（SheetJS の SSF ／ うちの 台）。
 *    ⇒ COM も 写しも 要らず、★本を 読むだけ★。
 *    （1回目は PowerShell で 実Excel に 聞こうと しましたが、
 *      `.Text` も `.NumberFormatLocal` も ★かたまりでは null★＝1マスずつ 触る事に なり、
 *      ★137,907マス ぶんの COM★に なる。★そもそも 相手は 実Excel では ない★ので やめた）
 *
 *  ★★大事な 見つけ物★★
 *    `:378` は ★いつも 通る 道では ありません★＝
 *      `cell.d` が ★空か 数★の 時だけ 通る ★逃げ道★です。
 *      ふだんは ★`c.w`（ファイルに 入って いる 字）★を そのまま 出して います。
 *    ⇒★差し替えで 変わるのは「逃げ道を 通る マス」だけ★
 *    ⇒★その 数も 出します★（★「67,542マス 変わる」では ありません★）
 *
 *  ★★出す 物★★（指示役1 の 決め）
 *    ★数と 書式の 字だけ★／★会社名・金額・人の 名前は 1文字も★／★マスの 中身も 出さない★
 *    ★違った マスは「どの 板の 何行何列」まで★（★中身は 書かない★）
 *
 *  使い方: node docs/measured/osu-jitsubutsu-mitame.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const XLSX = require_(path.join(ROOT, 'lib/xlsx.full.min.js'));
const 台 = require_(path.join(ROOT, 'lib/shoshiki.js'));

const 本 = 'C:\\Users\\zeroa\\kyukyu-0907\\代行計算表2026.xlsb';
if (!fs.existsSync(本)) { console.error('★本体が 無い★ ' + 本); process.exit(2); }
const 前 = fs.statSync(本);

/* ★`js/book-open.js` の 2つの 下請けを そのまま 写す★
   ★写した 事を 書いて 置きます★＝★差し替える 時に 2つとも 台へ 寄せます★ */
const WD = ['日', '月', '火', '水', '木', '金', '土'];
/* ★★本当の 道は `book.html:3466`★★（2026-09-15 に 見つけた）
   `js/book-open.js:378` は ★逃げ道★（`c.w` が 無い 時だけ）。
   ★画面が ふだん 通るのは `fmtForDisplay`★で、そこは
   ★書式を 掛ける 前に 15桁に 丸めて います★（`_十五桁`／2026-09-11 の 直し）。
   ★`:378` には その 直しが 在りません★＝★同じ 事を する 道が 2本／片方だけ 直って いる★
   ⇒★台に 寄せれば 1本に なります★ */
function 十五桁(n) {                            /* book.html:3440 と 同じ */
  if (typeof n !== 'number' || !isFinite(n) || n === 0) return n;
  const v = Number(n.toPrecision(15));
  return isFinite(v) ? v : n;
}
function withWeekday(fmt, serial) {            /* book-open.js:217 と 同じ */
  if (!/a{3,4}/.test(String(fmt))) return fmt;
  const d = new Date(Date.UTC(1899, 11, 30) + Math.floor(serial) * 86400000);
  const w = WD[d.getUTCDay()];
  return String(fmt).replace(/a{4}/g, '"' + w + '曜日"').replace(/a{3}/g, '"' + w + '"');
}

const wb = XLSX.read(fs.readFileSync(本), { type: 'buffer', cellNF: true, cellText: true });

let 数のマス = 0, 逃げ道 = 0, 同じ = 0, 違う = 0, 台が出せない = 0, SSFが出せない = 0;
const 書式ごと = {};          /* 書式 → {全, 違, 出せない} */
const 場所 = [];              /* ★どの 板の 何行何列★（中身は 書かない） */
const 違いの形 = {};          /* ★数字を 9 に 潰した 形★ → 何マス（★中身は 出さない★） */

for (const 名 of wb.SheetNames) {
  const ws = wb.Sheets[名];
  if (!ws || !ws['!ref']) continue;
  const 範 = XLSX.utils.decode_range(ws['!ref']);
  for (let r = 範.s.r; r <= 範.e.r; r++) {
    for (let c = 範.s.c; c <= 範.e.c; c++) {
      const セル = ws[XLSX.utils.encode_cell({ r, c })];
      if (!セル || typeof セル.v !== 'number' || !セル.z) continue;
      数のマス++;
      /* ★逃げ道を 通るか★＝`cell.d` が 空か 数（＝`c.w` が 無い） */
      const 逃げるか = (セル.w === undefined || セル.w === '');
      if (逃げるか) 逃げ道++;

      const z = String(セル.z);
      if (!書式ごと[z]) 書式ごと[z] = { 全: 0, 違: 0, 出せない: 0, 逃: 0 };
      書式ごと[z].全++;
      if (逃げるか) 書式ごと[z].逃++;

      let a = null;
      try { a = XLSX.SSF.format(withWeekday(z, セル.v), 十五桁(セル.v)); } catch (e) { a = null; }
      if (a === undefined || a === '') a = null;
      if (a === null) SSFが出せない++;

      const b = 台.当てる(セル.v, z);
      if (!b.出せる) { 台が出せない++; 書式ごと[z].出せない++; }

      const うちの字 = b.出せる ? b.字 : null;
      if (a === うちの字) { 同じ++; continue; }
      違う++;
      書式ごと[z].違++;
      /* ★★中身を 出さずに 違いの 形だけ 見る★★
         ＝★数字を 全部 `9` に 潰します★（金額は 1文字も 出さない／形は 分かる） */
      const 伏 = (x) => (x === null ? '(出せない)' : String(x).replace(/[0-9]/g, '9'));
      const 形 = 伏(a) + ' ／ ' + 伏(うちの字);
      違いの形[形] = (違いの形[形] || 0) + 1;
      if (場所.length < 25) 場所.push('    板' + (wb.SheetNames.indexOf(名) + 1)
        + ' 行' + (r + 1) + ' 列' + (c + 1) + '  書式 `' + z + '`'
        + (逃げるか ? ' ★逃げ道★' : ''));
    }
  }
}

const 後 = fs.statSync(本);
const 触ったか = (前.mtimeMs === 後.mtimeMs && 前.size === 後.size)
  ? '★動いて いません★' : '★★動いた＝すぐ 止める★★';

console.log('# ★SheetJS の 字と うちの 台の 字を 全マス 突き合わせた★（2026-09-15）');
console.log('#   ★実Excel は 使って いません★＝★くらべる 相手は どちらも JS★');
console.log('#   ★本体は 読むだけ★ … ' + 後.size + ' バイト … ' + 触ったか);
console.log('#   ★出すのは 数と 書式の 字だけ★（会社名・金額・人の 名前は 1文字も 出して いません）');
console.log('');
console.log('★数の マス（書式つき）★ … ' + 数のマス);
console.log('  ★その うち 逃げ道（`:378` を 通る）★ … ' + 逃げ道
  + '（' + (数のマス ? Math.round(逃げ道 / 数のマス * 1000) / 10 : 0) + '%）');
console.log('  ★★これが「差し替えで 変わりうる マス」の 上限★★');
console.log('');
console.log('★同じ 字★ … ' + 同じ + ' ／ ★違う 字★ … ' + 違う
  + '（' + (数のマス ? Math.round(違う / 数のマス * 1000) / 10 : 0) + '%）');
console.log('  SSF が 出せない … ' + SSFが出せない + ' ／ 台が 出せない … ' + 台が出せない);
console.log('');
console.log('★書式ごと（違いが 多い順・全部）★');
const 並 = Object.keys(書式ごと).sort((x, y) => 書式ごと[y].違 - 書式ごと[x].違);
for (const z of 並) {
  const v = 書式ごと[z];
  console.log('  ' + String(v.全).padStart(6) + 'マス  違い ' + String(v.違).padStart(6)
    + '  逃げ道 ' + String(v.逃).padStart(6)
    + '  台が出せない ' + String(v.出せない).padStart(6) + '   `' + z + '`');
}
console.log('');
console.log('★★違いの 形★★（★数字は 全部 `9` に 潰して 在ります＝金額は 1文字も 出て いません★）');
console.log('   SheetJS ／ うちの 台');
Object.keys(違いの形).sort((x, y) => 違いの形[y] - 違いの形[x]).slice(0, 20)
  .forEach((k) => console.log('  ' + String(違いの形[k]).padStart(5) + 'マス  ' + k));
console.log('');
console.log('★違った マス（先頭 25・★中身は 書いて いません★）★');
場所.forEach((x) => console.log(x));
