/* shiki-convert.mjs — ★CONVERT の 単位表を 実Excel から 取る 為の 式★（2026-09-07）
 *
 *  ★★なぜ 実Excel から 取るのか★★
 *    単位の 換算は ★私が 覚えている 数を 書いたら そこで 終わり★です。
 *    ⇒★1ポンド＝0.45359237kg を 0.4536 と 書いても 誰も 気づきません★
 *    ⇒★実際 本番が そう なっていました★（4桁で 丸めていた）
 *    ⇒★★だから 1つずつ 実Excel に 聞きます★★
 *
 *  ★取り方★
 *    ①同じ 種類の 中で ★基準の 単位★を 決め、`=CONVERT(1, 単位, 基準)` を 打たせる
 *      ⇒ 返って きた 数が ★そのまま 換算の 係数★
 *    ②接頭辞（k・M・m…）は ★メートルに 付けて★ 打たせる
 *    ③温度は 掛け算では ない ので ★2点★（0 と 100）で 打たせる
 *    ④★通らない 組み合わせ★も 打たせる（種類を またぐ／知らない 単位）
 *
 *  使い方: node docs/measured/kansuu46/shiki-convert.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ここ = path.dirname(fileURLToPath(import.meta.url));

/* ★種類ごとの 単位★（Excel の 説明に 載っている 物を 並べた／★係数は 書かない★） */
const 種類 = {
  重さ: ['g', 'sg', 'lbm', 'u', 'ozm', 'grain', 'cwt', 'shweight', 'uk_cwt', 'lcwt', 'hweight',
    'stone', 'ton', 'uk_ton', 'LTON', 'brton'],
  長さ: ['m', 'mi', 'Nmi', 'in', 'ft', 'yd', 'ang', 'ell', 'ly', 'parsec', 'pc', 'survey_mi',
    'Picapt', 'Pica', 'pica'],
  時間: ['yr', 'day', 'd', 'hr', 'mn', 'min', 'sec', 's'],
  圧力: ['Pa', 'p', 'atm', 'at', 'mmHg', 'psi', 'Torr'],
  力: ['N', 'dyn', 'dy', 'lbf', 'pond'],
  エネルギー: ['J', 'e', 'c', 'cal', 'eV', 'ev', 'HPh', 'hh', 'Wh', 'wh', 'flb', 'BTU', 'btu'],
  仕事率: ['HP', 'h', 'PS', 'W', 'w'],
  磁気: ['T', 'ga'],
  体積: ['tsp', 'tspm', 'tbs', 'oz', 'cup', 'pt', 'us_pt', 'uk_pt', 'qt', 'uk_qt', 'gal', 'uk_gal',
    'l', 'L', 'lt', 'ang3', 'barrel', 'bushel', 'ft3', 'in3', 'ly3', 'm3', 'mi3', 'yd3', 'Nmi3',
    'Picapt3', 'Pica3', 'GRT', 'regton'],
  面積: ['m2', 'mi2', 'Nmi2', 'in2', 'ft2', 'yd2', 'ang2', 'Picapt2', 'Morgen', 'ar',
    'uk_acre', 'us_acre', 'ly2', 'ha', 'Pica2'],
  情報: ['bit', 'byte'],
  速さ: ['admkn', 'kn', 'm/h', 'm/hr', 'mph', 'm/s', 'm/sec'],
};
const 基準 = {
  重さ: 'g', 長さ: 'm', 時間: 'sec', 圧力: 'Pa', 力: 'N', エネルギー: 'J', 仕事率: 'W',
  磁気: 'T', 体積: 'l', 面積: 'm2', 情報: 'bit', 速さ: 'm/s',
};
/* ★^ を 使う 書き方も 通るか★（体積・面積） */
const 別書き = ['ang^3', 'ft^3', 'in^3', 'ly^3', 'm^3', 'mi^3', 'yd^3', 'Nmi^3', 'Picapt^3', 'Pica^3',
  'm^2', 'mi^2', 'Nmi^2', 'in^2', 'ft^2', 'yd^2', 'ang^2', 'Picapt^2', 'ly^2', 'Pica^2'];

/* ★接頭辞★（10進 と 2進） */
const 頭 = ['Y', 'Z', 'E', 'P', 'T', 'G', 'M', 'k', 'h', 'e', 'd', 'c', 'm', 'u', 'n', 'p', 'f', 'a', 'z', 'y',
  'da', 'Yi', 'Zi', 'Ei', 'Pi', 'Ti', 'Gi', 'Mi', 'ki'];

const 行 = [];
const 足す = (式) => 行.push('CONVERT\t=' + 式);

for (const k of Object.keys(種類)) {
  for (const u of 種類[k]) 足す('CONVERT(1,"' + u + '","' + 基準[k] + '")');
}
for (const u of 別書き) {
  const b = /\^3$/.test(u) ? 'l' : 'm2';
  足す('CONVERT(1,"' + u + '","' + b + '")');
}
/* ★接頭辞★ … メートルと ビットの 両方（2進の 頭は ビットにしか 付かない はず） */
for (const p of 頭) {
  足す('CONVERT(1,"' + p + 'm","m")');
  足す('CONVERT(1,"' + p + 'bit","bit")');
}
/* ★温度★＝掛け算では ない ので 2点 */
for (const t of ['C', 'cel', 'F', 'fah', 'K', 'kel', 'Rank', 'Reau']) {
  足す('CONVERT(0,"' + t + '","C")');
  足す('CONVERT(100,"' + t + '","C")');
  足す('CONVERT(0,"C","' + t + '")');
  足す('CONVERT(100,"C","' + t + '")');
}
/* ★通らない 組み合わせ★ */
足す('CONVERT(1,"m","g")');
足す('CONVERT(1,"あ","cm")');
足す('CONVERT(1,"m","")');
足す('CONVERT(1,"M","m")');
足す('CONVERT(1,"km","M")');
足す('CONVERT(1,"kg","g")');
足す('CONVERT(1,"kibyte","byte")');
足す('CONVERT(1,"Mibyte","byte")');
足す('CONVERT(1,"kbyte","byte")');

fs.writeFileSync(path.join(ここ, 'cases-convert.txt'), 行.join('\n') + '\n', { encoding: 'utf-8' });
console.log('★式 ' + 行.length + '本★ … cases-convert.txt に 書いた');
