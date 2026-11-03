/* xlsb-jitai.test.mjs — ★.xlsb の マスごとの 字体を 自分で 読む★（2026-09-11）
 *
 *  ★★なぜ 自分で 読むか★★
 *    借り物（SheetJS）は ★.xlsb の マスごとの 字体を くれません★
 *      実測 … 司さんの 実物 ★35,760マス中 0マス★（`c.s` が 付かない／xlsx なら 付く）
 *      字体の 表（31本）と マスの 形（139本）は 在るのに ★結び付ける 番号が 来ない★
 *    ⇒ 実Excel が 9pt で 書いて 在る 所を うちは 既定（11pt）で 描き、
 *      ★入りきらず `######`★ に なって いました（給料表 120マス）
 *
 *  ★★借り物の 中は 1文字も 読んで いません★★
 *    使うのは ★この repo が 前から 持って いる 物★だけ
 *      `lib/zip-surgeon.js`（袋を 開ける）／`lib/xlsb-edit.js`（記録を 歩く）
 *
 *  ★★当て推量で 決めて いません★★
 *    ・字体の 記録＝★番号43★（数が 借り物の 言う「字体31本」と 一致）
 *    ・マスの 形＝★番号47★（142本＝借り物の 言う 139本 ＋ 3）
 *    ・★形の 番号は 3つ ずらす★
 *        司さんの 実物 838マスで … ずれ0 は ★94本★しか 合わない ／ ずれ3 は ★838本 全部★
 *    ・大きさは ★twip（pt × 20）★
 *
 *  ★★まだ 画面には 繋いで いません★★（2026-09-11）
 *    繋いだら ★ファイルが 開かなく なりました★（120秒 待っても 開かない）。
 *    原因を 掴めて いないので ★戻しました★＝★開かない 物は 出しません★。
 *    ⇒ ここでは ★部品が 正しく 読める事★だけを 見張ります。
 *
 *  ★見本★ tests/fixtures/jitai-2026-09-11.xlsb（★実Excel に 作らせました★）
 *    いちまいめ … 9pt / 11pt / 12pt / 18pt / 8pt ／ にまいめ … 16pt
 *    ★司さんの 実物は 使いません★（取引先の 名前・金額）
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const 見本 = path.join(ROOT, 'tests/fixtures/jitai-2026-09-11.xlsb');

let pass = 0, fail = 0;
const T = (n, f) => {
  try { f(); pass++; console.log('  ✓ ' + n); }
  catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); }
};

console.log('');
console.log('[xlsb-jitai] ★.xlsb の マスごとの 字体を 自分で 読む★');

const Z = require_(path.join(ROOT, 'lib/zip-surgeon.js'));
const XE = require_(path.join(ROOT, 'lib/xlsb-edit.js'));
const JT = require_(path.join(ROOT, 'lib/xlsb-jitai.js'));

let 袋 = null, styles = null, 対応 = null;
const 出す = [];

async function 支度() {
  const buf = fs.readFileSync(見本);
  袋 = Z.read(new Uint8Array(buf));
  const とる = async (名) => {
    const list = 袋.entries || [];
    let e = null;
    for (const x of list) if (x.name === 名) { e = x; break; }
    if (!e) return null;
    return e.method === 0 ? e.raw : new Uint8Array(await Z.inflateRaw(e.raw));
  };
  const st = await とる('xl/styles.bin');
  styles = st ? XE.parse(st) : null;
  const wbb = await とる('xl/workbook.bin');
  const rels = await とる('xl/_rels/workbook.bin.rels');
  if (wbb && rels) {
    let 字 = '';
    for (let i = 0; i < rels.length; i++) 字 += String.fromCharCode(rels[i]);
    対応 = JT.板とファイル(XE.parse(wbb).recs, 字);
  }
  if (対応) {
    for (const 名 of Object.keys(対応)) {
      const sb = await とる(対応[名]);
      if (!sb) continue;
      const sp = XE.parse(sb);
      if (!sp.ok) continue;
      出す.push({ 名: 名, 表: JT.板の字体(styles.recs, sp.recs) });
    }
  }
}
await 支度();

T('★見本が 読めて いる（空振りして いない）★', () => {
  if (!fs.existsSync(見本)) throw new Error('★見本が 無い★');
  if (!styles || !styles.ok) throw new Error('★styles.bin を 歩けない★');
  if (!対応) throw new Error('★板と ファイルの 対応が 取れない★');
});

T('★★板の 名前と ファイルが 繋がる（★番号で 当てない★）★★', () => {
  const 名 = Object.keys(対応);
  if (名.length < 2) throw new Error('★' + 名.length + '枚しか 取れない★');
  for (const n of 名) {
    if (!/^xl\/worksheets\/sheet\d+\.bin$/.test(対応[n])) {
      throw new Error('★' + n + ' の 行き先が おかしい … ' + 対応[n] + '★');
    }
  }
  console.log('      … ' + 名.length + '枚（' + 名.join('／') + '）');
});

T('★★実Excel に 打たせた 大きさと 合う★★', () => {
  /* ★実Excel が 言った 事★（見本を 作った 時に 一緒に 測った） */
  const 正 = { 'いちまいめ': { '0,0': 9, '1,0': 11, '2,0': 12, '3,0': 18, '4,0': 8 },
               'にまいめ': { '0,0': 16 } };
  const 違い = [];
  for (const s of 出す) {
    const 期待 = 正[s.名];
    if (!期待) continue;
    if (!s.表) { 違い.push(s.名 + ' の 字体が 1つも 取れない'); continue; }
    for (const k in 期待) {
      const f = s.表[k];
      if (!f) { 違い.push(s.名 + ' ' + k + ' が 取れない'); continue; }
      if (f.pt !== 期待[k]) 違い.push(s.名 + ' ' + k + ' 実Excel ' + 期待[k] + 'pt ／ うち ' + f.pt + 'pt');
    }
  }
  if (違い.length) throw new Error('★' + 違い.length + '本 違う★  ' + 違い.join(' ／ '));
  console.log('      … 9/11/12/18/8pt ／ 別の 板の 16pt も 合う');
});

T('★★書体の 名前も 取れる★★', () => {
  const s = 出す.find((x) => x.名 === 'いちまいめ');
  const f = s && s.表 && s.表['0,0'];
  if (!f) throw new Error('★取れない★');
  if (!f.名 || f.名.length < 2) throw new Error('★書体の 名前が 空★（実Excel は 游ゴシック）');
});

T('★★形の 表は 印で 切る（★数で 当てない★）★★', () => {
  /* ★★2026-09-11 ここで つまずきました★★
     最初は「3つ ずらす」と して いました＝★司さんの ファイルにだけ 合う 数★。
     見本（別の ファイル）では ★6本 全部 外れ★、この 見張りが 捕まえました。
     ⇒★印（617 … 47 … 618）で 切ります★＝どの ファイルでも 合う
     実測 … 司さんの実物 ①3本 ②139本 ／ 見本 ①1本 ②8本 */
  if (JT.形のずれ !== undefined) {
    throw new Error('★「ずらす 数」が 戻って います★＝1つの ファイルにだけ 合う 決め方です');
  }
  if (JT.形の始まり !== 617 || JT.形の終わり !== 618) {
    throw new Error('★印が 変わって います（' + JT.形の始まり + '/' + JT.形の終わり + '）★');
  }
});

console.log('');
console.log('  ' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
