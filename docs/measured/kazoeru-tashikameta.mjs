/* kazoeru-tashikameta.mjs — ★『名前が 通った』の うち ★答えを 確かめた 事が 在る★のは 何本か★
 *（2026-09-07）
 *
 *  ★★なぜ 要るか★★
 *    `tests/kansuu-kabaa.test.mjs` の 数（491個）は
 *    ★引数の 型紙 17通りの うち 1つが #NAME? を 返さなかった★だけの 数です。
 *    ⇒★答えが 合っているかは 1度も 見ていません★
 *    ⇒ そこで ★実Excel の 答えと 突き合わせた 事が 在るか★で 分けます。
 *
 *  ★★数え方（★甘い 向きも 先に 書きます★）★★
 *    ①★紙★ … 実Excel の 答えを 書いた 物（docs/measured/**、tests/fixtures/**）に
 *              その 関数が `FOO(` の 形で 出てくるか
 *    ②★試験★ … tests/*.mjs に 出てくるか（★押した 数の 台帳は 除く★）
 *    ★甘い①★ ★ついでに 出てきただけ★を 数えます
 *      ⇒ SUM / IF / INDEX / MAX / AND / T / N / CHOOSE は
 *        ★他の 関数の 試験の 式に 出てくるだけ★＝★その 関数を 確かめた 訳では ない★
 *      ⇒★あ）い）は ★多め（甘い）★に 出ます／う）は ★少なめ★に 出ます★
 *    ★甘い②★ 試験の 中で 名前を 組み立てている 物は 拾えません
 *    ⇒★★どちらにしても『1本 押しただけ』が 大多数★という 向きは 変わりません★★
 *
 *  ★★道具は 1つに します★★（2026-09-07 に 踏んだ）
 *    同じ 事を 数える 道具を 2つ 作ったら ★5個 食い違いました★（62 と 67）。
 *    ⇒★★食い違う 2つの 物差しは、1つの 物差しより 悪い★★
 *    ⇒ この 1本に まとめました。
 *
 *  使い方: node docs/measured/kazoeru-tashikameta.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

/* ★手元の 絶対の 道を 焼き込まない★（焼き込むと ★手元は 緑・CI だけ 赤★） */
const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..', '..');

const 読む = (p) => fs.readFileSync(p, 'utf-8').split('\n')
  .map((s) => s.trim()).filter((s) => s && !s.startsWith('#'));

const 全部 = 読む(path.join(ROOT, 'docs/measured/excel-functions-2026-09-06.txt'));
const 通らない = 読む(path.join(ROOT, 'docs/measured/exally-missing-2026-09-07.txt'));
const 通った = 全部.filter((f) => !通らない.includes(f));

/* ★台帳そのものは 数えない★＝名前が 載っているだけで「確かめた」に なってしまう */
const 除く = /excel-functions-|exally-missing-|kansuu\.md$|ugoku-tana-/;

function 集める(dirs, 拡張) {
  const 出 = [];
  const 掘る = (d) => {
    if (!fs.existsSync(d)) return;
    for (const e of fs.readdirSync(d, { withFileTypes: true })) {
      const p = path.join(d, e.name);
      if (e.isDirectory()) 掘る(p);
      else if (拡張.some((x) => e.name.endsWith(x)) && !除く.test(p.replace(/\\/g, '/'))) 出.push(p);
    }
  };
  for (const d of dirs) 掘る(path.join(ROOT, d));
  return 出;
}

const 読み込む = (ps) => ps.map((p) => {
  try { return fs.readFileSync(p, 'utf-8'); } catch (e) { return ''; }
});

/* ★★『答えの 紙』を 名前で 決める★★（2026-09-07 に 3回 数え直した 末）
 *  ⇒ はじめ `.txt` も 拾っていたら ★この 道具自身の 出力★を「答えの 紙」に 数えていた
 *    （`osu-wakeru.txt` には 関数名が ずらりと 並ぶ）
 *  ⇒★★物差しが 自分の 出した 物を 食べていた★★＝数が 67 → 113 と 動いた
 *  ⇒ だから ★実Excel に 打たせた 物だけ★に 絞る
 *    ・`docs/measured/**\/golden-*.tsv` … 実Excel の 答え（この repo の 決まり）
 *    ・`tests/fixtures/*.json` ……………… 実Excel が 作った 見本
 */
const 金の紙か = (p) => /\/golden-/.test(p.split(path.sep).join('/'));
const 紙 = 読み込む([
  ...集める(['docs/measured'], ['.tsv']).filter(金の紙か),
  ...集める(['tests/fixtures'], ['.json']),
]);
/* ★押した 数の 台帳は 試験から 外す★＝そこには 519個 全部の 名前が 出てくる */
const 試験 = 読み込む(集める(['tests'], ['.mjs'])
  .filter((p) => !/kansuu-kabaa|kansuu-tana|ugokanai-osu|hairanai/.test(p.replace(/\\/g, '/'))));

const 出るか = (f, 束) => {
  const re = new RegExp('(^|[^A-Z0-9._])' + f.replace(/\./g, '\\.') + '\\s*\\(');
  return 束.some((t) => re.test(t));
};

const 紙に = 通った.filter((f) => 出るか(f, 紙));
const 紙か試験に = 通った.filter((f) => 紙に.includes(f) || 出るか(f, 試験));
const 一度も = 通った.filter((f) => !紙か試験に.includes(f));

const 行 = [];
const 言う = (s) => { 行.push(s); console.log(s); };
言う('# ★『名前が 通った』の うち 答えを 確かめた 事が 在るのは 何本か★（'
  + new Date().toISOString().slice(0, 10) + '）');
言う('');
言う('★見た 紙 … ' + 紙.length + '本 ／ 見た 試験 … ' + 試験.length + '本★');
言う('');
言う('★実Excel が 知っている ………………………… ' + 全部.length + '個');
言う('★名前が 通った ………………………………… ' + 通った.length + '個');
言う('★名前も 通らない ……………………………… ' + 通らない.length + '個');
言う('');
言う('★あ）実Excel の 答えの 紙に 出てくる … ' + 紙に.length + '個'
  + '  ★（甘い＝本当は これより 少ない）★');
言う('★い）紙 または 試験に 出てくる ………… ' + 紙か試験に.length + '個'
  + '  ★（甘い＝本当は これより 少ない）★');
言う('★う）どちらにも 1度も 出てこない ……… ★' + 一度も.length + '個★'
  + '  ★（甘い＝本当は これより 多い）★');
言う('  （' + (一度も.length / 通った.length * 100).toFixed(1) + '%）');
言う('');
言う('★★「一部しか 動かない 物」は ★う）の 中に 居ます★（外には 居ません）★★');
言う('');
言う('★う）の 名簿★');
for (let i = 0; i < 一度も.length; i += 12) 言う('  ' + 一度も.slice(i, i + 12).join(' '));

fs.writeFileSync(path.join(ここ, 'ugoku-tana-mikakunin.txt'), 行.join('\n') + '\n', { encoding: 'utf-8' });
console.log('\n★書いた … docs/measured/ugoku-tana-mikakunin.txt★');
