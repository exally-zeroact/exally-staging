/* _tsukurimono.mjs — ★実物の本の字を 作り物に 替えてから 試験に 渡す★（2026-10-10）
 *
 *  ★なぜ★ 公開 repo に 働く人の 実名が 在った（司さんの決め 10-10＝本物として扱い 作り物に 替える）。
 *    試験は 作り物の 名前で 書く。司さんの 実物の 本（この 機械にだけ 在る）は 実名の 列を 持つので、
 *    読み込んだ 直後に 本の 字を 作り物に 替えて、試験と 同じ 字の 世界に そろえる。
 *  ★本物の 名前は repo に 置かない★＝対応表は repo の 外（~/.tsukurimono/okikae-hyou.json・作り＝ダイコメの席）。
 *    表が 無い 機械（CI など）では null ＝ 呼ぶ側は「未測定」と 言う（緑に しない）。
 */
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

export function 対応表を読む() {
  try {
    const j = JSON.parse(fs.readFileSync(path.join(os.homedir(), '.tsukurimono', 'okikae-hyou.json'), 'utf8'));
    if (!j || !j.people || !Object.keys(j.people).length) return null;
    return { people: j.people, companies: j.companies || {} };
  } catch (e) { return null; }
}

/** sheets＝[{ name, data: { 'r,c': { v, f, d } } }] の 字を 作り物に 替える（その場で）。替えた 字の 数を 返す。
 *  ・本物の フルネーム（人・社）は どこに 在っても 作り物に
 *  ・セルの 字が まるごと 本物の 名前の 頭（2字以上）で、当たる 人が 1人だけ の時は、その人の 作り物の フルネームに
 *    （見出しが 名字だけの 列＝名字が 同じ 人が 2人 居る時は 決められないので 替えない） */
export function 本の字を作り物に(sheets, 表) {
  const 全名 = Object.entries(Object.assign({}, 表.companies, 表.people)).sort((a, b) => b[0].length - a[0].length);
  const 人 = Object.entries(表.people);
  let 数 = 0;
  const 替える = (s) => {
    if (typeof s !== 'string' || !s) return s;
    let t = s;
    for (const [本, 作] of 全名) if (t.indexOf(本) >= 0) t = t.split(本).join(作);
    if (t === s && s.length >= 2) {
      const 候補 = 人.filter(([本]) => 本.length > s.length && 本.startsWith(s));
      if (候補.length === 1) t = 候補[0][1];
    }
    if (t !== s) 数++;
    return t;
  };
  for (const sh of sheets) {
    for (const k of Object.keys(sh.data || {})) {
      const c = sh.data[k];
      if (!c) continue;
      c.v = 替える(c.v); c.f = 替える(c.f); c.d = 替える(c.d);
    }
  }
  return 数;
}
