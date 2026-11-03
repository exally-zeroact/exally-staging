/* xirr-mon.test.mjs — ★XIRR が 実Excel と 同じ 条件で 同じ 誤りを 返すか★（2026-09-09）
 *
 *  ★★何が 起きて いたか★★
 *    実Excel が ★#NUM!★ を 返す 組で、うちは ★数を 返して いました★
 *      =XIRR(全部プラス)  … 実Excel #NUM! ／ うち ★375686054239746.3★
 *      =XIRR(日付が逆順)  … 実Excel #NUM! ／ うち ★-0.5002981524324954★
 *    ⇒★誤りに ならず 数が 出る＝お客さんは 気づけない★（★お金の 利回り★）
 *    ⇒ 375686054239746.3 は ★桁が おかしいのに「変」と 思えない 顔★を して います
 *
 *  ★★これは「機能を 減らす」では ありません＝実Excel と 同じに する★★
 *    ★断らない 所は 断りません★（空の マス・同じ 日付・0 を 含む は ★値を 返す★）
 *
 *  ★★どこから 断るかは ★私が 決めて いません★★★
 *    実Excel に 打たせました … `docs/measured/golden-xirr-sakaime-2026-09-09.tsv`（77本）
 *
 *  ★★何を 守るか（1つずつ 名指し）★★
 *    ①★断る 条件★ … 符号が 片方だけ／日付が 昇順でない／日付が 負／
 *                    範囲の 大きさが 違う／1件だけ／字が 混ざる／見当が 0 より 小さい
 *    ②★断らない 条件★ … 空の マス／同じ 日付／0 を 含む（★出来る 物を 止めない★）
 *    ③★見当で 誤りが 変わらない★（★入口の 門であって 繰り返し 計算の 失敗では ない★）
 *    ④★見当が 大きくても 落ち着く★（★前は NaN に なって いた★）
 *
 *  ★★見て いない 範囲（★書かない 見張りは「全部 守った」と 読まれる★）★★
 *    ・★MIRR は 見て いません★＝★向きが 逆の 穴が 2本 在ります★（別件）
 *        `=MIRR(全部マイナス,0.1,0.12)` 実Excel ★−1★ ／ うち ★#DIV/0!★
 *        `=MIRR(A1:A3,-1,0.12)`        実Excel ★0.17132403714770583★ ／ うち ★#DIV/0!★
 *        ⇒★実Excel が 答えるのに うちが 断って いる＝出来る 物を 止めて いる★
 *    ・★NPV は 測って 穴が 0本でした★（★触って いません★）
 *    ・★答えの 細かい 桁は 見て いません★（★幅 1e-6★＝実Excel 自身が 見当で 9.65e-8 ばらつく）
 *    ・★XIRR と XNPV は JS層／IRR は エンジン側（EX.IRR）★＝★本番と 同じ 道で 押します★
 *
 *  使い方: node tests/xirr-mon.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..');
const 直に走った = process.argv[1] && pathToFileURL(process.argv[1]).href === import.meta.url;
const 自己試験 = 直に走った && process.argv.includes('--self-test');

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };

const require_ = createRequire(path.join(ROOT, 'package.json'));
const HFns = require_(path.join(ROOT, 'hyperformula.full.min.js'));
const EF = require_(path.join(ROOT, 'exally-formula.js'));
EF.registerExallyFunctions(HFns);
const hf = HFns.HyperFormula.buildEmpty({
  licenseKey: 'gpl-v3', useArrayArithmetic: true, smartRounding: false,
});
const SID = hf.getSheetId(hf.addSheet('S'));
EF.initExallyFormula(hf);

/* ★A列＝お金 ／ B列＝日付★（★組ごとに 敷き直す★） */
function 押す(値ら, 日ら, 見当) {
  const 板 = [];
  const n = Math.max(値ら.length, 日ら.length);
  for (let i = 0; i < n + 2; i++) 板.push([null, null, null]);
  for (let i = 0; i < n; i++) {
    板[i][0] = i < 値ら.length ? 値ら[i] : null;
    板[i][1] = i < 日ら.length ? 日ら[i] : null;
  }
  hf.setSheetContent(SID, 板);
  const 式 = '=XIRR(A1:A' + 値ら.length + ',B1:B' + 日ら.length
    + (見当 === undefined ? '' : ',' + 見当) + ')';
  return { 式, 答: 本番の道(板, 式) };
}
/* ★★本番と 同じ 道★★ ①JS層 → ②convertFormula → エンジン */
function 本番の道(板, 式) {
  hf.setSheetContent(SID, 板.map((r) => r.slice()));
  const js = EF._jsComputeFormula(0, 式);
  if (js !== null && js !== undefined) return String(js);
  const 盤 = 板.map((r) => r.slice());
  while (盤.length <= 式の行) 盤.push(new Array(盤[0].length).fill(null));
  盤[式の行][0] = EF.convertFormula(式);
  hf.setSheetContent(SID, 盤);
  const v = hf.getCellValue({ sheet: SID, row: 式の行, col: 0 });
  if (v && v.type) return '#' + v.type;
  return v === null || v === undefined ? '(空)' : String(v);
}
/* ★★本番と 同じ 道で 押す（2026-09-09 に 足した）★★
   ★1回目は JS層だけ★を 押して いました。
   ⇒ IRR を ★エンジン側（EX.IRR）★に 移したら ★null が 返り 見張りが 赤★に
   ⇒★本番は ①JS層 → ②convertFormula → エンジン★＝★同じ 道で 押す★
   （★見張りが 見る 道を 本番と 揃える★＝今日 3回 踏んだ 型） */
const 式の行 = 30;
const 近い = (う, 正) => {
  const a = Number(う), b = Number(正);
  if (!isFinite(a) || !isFinite(b)) return false;
  return Math.abs((a - b) / b) <= 1e-6;   /* ★実測で 決めた 幅★ */
};

console.log('\n[xirr-mon] ★XIRR が 実Excel と 同じ 条件で 同じ 誤りを 返すか★');

T('★★断る 条件（実Excel の 実測どおり）★★', () => {
  const 組 = [
    [[100, 200, 300], [45292, 45383, 45474], undefined, '#NUM!', '★全部 プラス★'],
    [[-100, -200, -300], [45292, 45383, 45474], undefined, '#NUM!', '★全部 マイナス★'],
    [[0, 0, 0], [45292, 45383, 45474], undefined, '#NUM!', '★全部 0★'],
    [[-1000, 600, 700], [45474, 45383, 45292], undefined, '#NUM!', '★日付が 逆順★'],
    [[-1000, 600, 700], [45383, 45292, 45474], undefined, '#NUM!', '★日付が バラバラ★'],
    [[-1000, 600, 700], [-10, 45383, 45474], undefined, '#NUM!', '★日付が 負★'],
    [[-1000, 600, 700], [45292, 45383], undefined, '#NUM!', '★範囲の 大きさが 違う★'],
    [[-1000], [45292], undefined, '#N/A', '★1件だけ（#NUM! では ない）★'],
    [[-1000, 'あ', 1100], [45292, 45383, 45474], undefined, '#VALUE!', '★字が 混ざる★'],
    [[-1000, 600, 700], [45292, 45383, 45474], -0.01, '#NUM!', '★見当が −0.01★'],
    [[-1000, 600, 700], [45292, 45383, 45474], -0.9, '#NUM!', '★見当が −0.9★'],
  ];
  for (const [v, d, g, 正, 札] of 組) {
    const r = 押す(v, d, g);
    if (r.答 !== 正) throw new Error('★' + 札 + '（' + r.式 + '）… うち ' + r.答 + ' ／ 実Excel ' + 正 + '★');
  }
  console.log('      … ' + 組.length + '通り とも 実Excel と 同じ 誤り');
});

T('★★断らない 条件（★出来る 物を 止めない★）★★', () => {
  const 組 = [
    [[-1000, 600, 700], [45292, 45383, 45474], undefined, 1.00119332075119, '普通'],
    [[-1000, 1100], [45292, 45657], undefined, 0.09999999403953552, '★一番 短い（2件）★'],
    [[-1000, 0, 1100], [45292, 45383, 45474], undefined, 0.21063382029533387, '★0 を 含む★'],
    [[-1000, null, 1100], [45292, 45383, 45474], undefined, 0.21063382029533387, '★空の マス（＝0 と 見る）★'],
    [[-1000, 600, 700], [45292, 45292, 45474], undefined, 2.071931099891663, '★同じ 日付が 2つ★'],
    [[-1000, 600, 700], [45292, 45383, 45474], 0, 1.0011933300781244, '★見当が 0（境目・通る）★'],
  ];
  for (const [v, d, g, 正, 札] of 組) {
    const r = 押す(v, d, g);
    if (typeof r.答 === 'string' && r.答[0] === '#') {
      throw new Error('★' + 札 + '（' + r.式 + '）を 断って いる（' + r.答 + '）★'
        + '\n      ⇒★実Excel は ' + 正 + ' を 返します＝★出来る 物を 止めた★');
    }
    if (!近い(r.答, 正)) throw new Error('★' + 札 + '… うち ' + r.答 + ' ／ 実Excel ' + 正 + '★');
  }
  console.log('      … ' + 組.length + '通り とも 値を 返し、実Excel と 合う');
});

T('★★見当を 変えても 誤りは 変わらない（入口の 門で あって 計算の 失敗では ない）★★', () => {
  for (const g of [undefined, 0.1, 0, 10, 1000000]) {
    const r = 押す([100, 200, 300], [45292, 45383, 45474], g);
    if (r.答 !== '#NUM!') throw new Error('★見当 ' + g + ' で ' + r.答 + '★（#NUM! のはず）');
  }
  console.log('      … 全部プラスは 見当 5通り とも #NUM!');
});

T('★★見当が 大きくても 落ち着く（前は NaN に なって いた）★★', () => {
  const 組 = [
    [10, 1.0011933278292418], [100, 1.0011933161877096],
    [10000, 1.0011933136411244], [1000000, 1.0011933255782424],
  ];
  for (const [g, 正] of 組) {
    const r = 押す([-1000, 600, 700], [45292, 45383, 45474], g);
    if (String(r.答) === 'NaN') throw new Error('★見当 ' + g + ' で NaN★＝★発散して いる★');
    if (!近い(r.答, 正)) throw new Error('★見当 ' + g + ' … うち ' + r.答 + ' ／ 実Excel ' + 正 + '★');
  }
  console.log('      … 見当 10／100／10000／1000000 とも 落ち着く');
});

T('★門と 計算が ★同じ 組★を 見て いる（口裏が 合って いる）★', () => {
  /* ★空の マスで 踏んだ★ … 門は 通したのに 計算側が 空を 落として 日付と ずれた */
  const 空あり = 押す([-1000, null, 1100], [45292, 45383, 45474], undefined);
  const ゼロあり = 押す([-1000, 0, 1100], [45292, 45383, 45474], undefined);
  if (String(空あり.答) !== String(ゼロあり.答)) {
    throw new Error('★空の マスと 0 で 答えが 違う★（' + 空あり.答 + ' ／ ' + ゼロあり.答 + '）'
      + '\n      ⇒★門は 空を 0 と 見るのに 計算が 落として いる＝口裏が 合って いない★');
  }
  console.log('      … 空の マスと 0 が 同じ 答え（' + 空あり.答 + '）');
});

/* ══ ★IRR と XNPV も 同じ 家族（2026-09-09 に 足した）★ ══════════
   ★XIRR を 直した 後、同じ 家族を 測ったら ★穴が 16本★ 在りました★
   ⇒ IRR 6本＋XNPV 8本 は ★同じ 直し方（実Excel と 同じ 誤りを 返す）★＝ここに 入れる
   ⇒ MIRR 2本は ★向きが 逆（実Excel が 答えるのに うちが 断る）★＝★別件★ */
T('★★IRR … 実Excel と 同じ 条件で 同じ 誤り★★', () => {
  const 組 = [
    [[100, 200, 300], undefined, '#NUM!', '★全部 プラス（前は 16693338463349.688）★'],
    [[-100, -200, -300], undefined, '#NUM!', '★全部 マイナス★'],
    [[0, 0, 0], undefined, '#NUM!', '★全部 0★'],
    [[-1000], undefined, '#NUM!', '★1件だけ（★XIRR は #N/A・家族でも 違う★）★'],
  ];
  for (const [v, g, 正, 札] of 組) {
    const 板 = v.map((x) => [x, null]);
    板.push([null, null]);
    const 式 = '=IRR(A1:A' + v.length + (g === undefined ? '' : ',' + g) + ')';
    const r = 本番の道(板, 式);
    if (r !== 正) throw new Error('★' + 札 + '（' + 式 + '）… うち ' + r + ' ／ 実Excel ' + 正 + '★');
  }
  console.log('      … ' + 組.length + '通り とも 実Excel と 同じ 誤り');
});

T('★★IRR … 見当が 何でも 同じ 答え（前は 発散して いた）★★', () => {
  /* ★実測 … −0.99〜100 の 11通りで 全部 0.1888…★ */
  const 正 = 0.18881944173074183;
  for (const g of [undefined, -0.99, -0.5, 0, 0.1, 1, 10, 100]) {
    const 式 = '=IRR(A1:A3' + (g === undefined ? '' : ',' + g) + ')';
    const r = 本番の道([[-1000, null], [600, null], [700, null], [null, null]], 式);
    if (String(r) === 'NaN') throw new Error('★見当 ' + g + ' で NaN★');
    if (!近い(r, 正)) throw new Error('★見当 ' + g + ' … うち ' + r + ' ／ 実Excel ' + 正 + '★'
      + ' ⇒★前は 見当 10 で -2168772783.613375 ／ 100 で -336432261411301570★');
  }
  console.log('      … 見当 8通り とも 0.1888…（★前は 大きい 見当で 発散★）');
});

T('★★XNPV … 利率の 境目は 0（★XIRR の 見当とは 逆★）★★', () => {
  const 敷く = () => hf.setSheetContent(SID, [[-1000, 45292], [600, 45383], [700, 45474], [null, null]]);
  /* ★0 以下は 断る★（刻んで 詰めた） */
  for (const r2 of [-2, -1, -0.5, -0.01, -0.001, 0]) {
    敷く();
    const v = EF._jsComputeFormula(0, '=XNPV(' + r2 + ',A1:A3,B1:B3)');
    if (v !== '#NUM!') throw new Error('★利率 ' + r2 + ' … うち ' + v + ' ／ 実Excel #NUM!★');
  }
  /* ★0 より 大きければ 通す★（★出来る 物を 止めない★） */
  敷く();
  const v1 = EF._jsComputeFormula(0, '=XNPV(0.0001,A1:A3,B1:B3)');
  if (!近い(v1, 299.9501405358503)) throw new Error('★利率 0.0001 … うち ' + v1 + ' ／ 実Excel 299.9501405358503★');
  敷く();
  const v2 = EF._jsComputeFormula(0, '=XNPV(0.1,A1:A3,B1:B3)');
  if (!近い(v2, 253.4216596369007)) throw new Error('★利率 0.1 … うち ' + v2 + ' ／ 実Excel 253.4216596369007★');
  console.log('      … 0 以下 6通り 断る ／ 0.0001 と 0.1 は 通る（★境目は 0★）');
});

T('★★XNPV … 空・字・日付の 順・大きさ は #NUM!（★XIRR と 違う★）★★', () => {
  const 組 = [
    [[-1000, null, 1100], [45292, 45383, 45474], '★空が 混ざる（★XIRR は 通す★）★'],
    [[-1000, 'あ', 1100], [45292, 45383, 45474], '★字が 混ざる（★XIRR は #VALUE!★）★'],
    [[-1000, 600, 700], [45474, 45383, 45292], '★日付が 逆順★'],
  ];
  for (const [v, d, 札] of 組) {
    const 板 = [];
    for (let i = 0; i < v.length; i++) 板.push([v[i], d[i]]);
    板.push([null, null]);
    hf.setSheetContent(SID, 板);
    const r = EF._jsComputeFormula(0, '=XNPV(0.1,A1:A3,B1:B3)');
    if (r !== '#NUM!') throw new Error('★' + 札 + ' … うち ' + r + ' ／ 実Excel #NUM!★');
  }
  console.log('      … 3通り とも #NUM!（★同じ 家族でも XIRR とは 違う★）');
});

T('★★NPV は 穴が 無かった（★測って 無かった事も 成果★）★★', () => {
  /* ★69本 中 NPV は 0本 穴が 在りませんでした★＝★触りません★
     ⇒ ここでは ★触って いない事★を 記録します（★直した ふりを しない★） */
  const s = fs.readFileSync(path.join(ROOT, 'exally-formula.js'), 'utf-8');
  if (/mNpv|NPV:1/.test(s)) {
    throw new Error('★NPV を JS層で 受けて いる★＝★穴が 無いのに 触って いる★');
  }
  console.log('      … NPV は JS層で 受けて いない（★穴 0本＝触らない★）');
});

T('★紙（実測）と 試験の 材料が 同じ★', () => {
  const p = path.join(ROOT, 'docs/measured/golden-xirr-sakaime-2026-09-09.tsv');
  if (!fs.existsSync(p)) throw new Error('★紙が 無い★ … ' + p);
  const 紙 = {};
  for (const l of fs.readFileSync(p, 'utf-8').split('\n')) {
    if (!l.startsWith('#材料')) continue;
    const c = l.split('\t');
    if (c.length >= 3) 紙[c[1].trim()] = c[2].trim();
  }
  /* ★A/B＝全部プラス ／ E/F＝普通 ／ M/N＝日付が 逆順★ */
  const 見る = [['A1', '100'], ['A3', '300'], ['E1', '-1000'], ['E3', '700'],
    ['M1', '-1000'], ['N1', '45474'], ['N3', '45292']];
  const 違う = 見る.filter(([k, v]) => 紙[k] !== v);
  if (違う.length) {
    throw new Error('★紙と 違う … ' + 違う.map(([k, v]) => k + '（紙 ' + 紙[k] + ' ／ 試験 ' + v + '）').join(' / ') + '★');
  }
  console.log('      … 紙と 同じ（' + 見る.length + 'マスを 名指しで 見た）');
});

/* ══ ★自己試験（★壊すのは 写し★＝ファイルは 1バイトも 触らない）★ ══ */
if (自己試験) {
  console.log('\n★自己試験（★壊すのは 写し★＝ファイルは 1バイトも 触らない）★');

  T('★★門が コードに 在る（空振りして いない）★★', () => {
    const s = fs.readFileSync(path.join(ROOT, 'exally-formula.js'), 'utf-8');
    if (!/function _xirrの門/.test(s)) throw new Error('★_xirrの門 が 無い★');
    if (!/function _xirrの組/.test(s)) throw new Error('★_xirrの組 が 無い★＝口裏を 合わせる 所が 無い');
    if (!/_xirrの門\(生値,生日,見当\)/.test(s)) throw new Error('★門を 呼んで いない★');
    console.log('      … 門・組・呼び出し の 3つとも 在る');
  });

  T('★★門を 外したら 数が 出て しまう（＝門が 効いて いる 証拠）★★', () => {
    /* ★門を 通さずに 直に 計算したら どうなるか★を 見る */
    const 直 = EF._jsXirr([100, 200, 300], [45292, 45383, 45474]);
    if (String(直) === 'NaN') { console.log('      … 門なしでは NaN（どちらにせよ ★値では ない★）'); return; }
    if (typeof 直 === 'string' && 直[0] === '#') throw new Error('★計算側が 断って いる＝門の 手柄では ない★');
    console.log('      … 門なしでは ' + 直 + ' が 出る（★実Excel は #NUM!★）');
  });

  T('★見当の 境目が 0 だと 書いて 在る（★−1 では ない★）★', () => {
    const s = fs.readFileSync(path.join(ROOT, 'exally-formula.js'), 'utf-8');
    if (/見当<=-1/.test(s)) throw new Error('★−1 の ままに なって いる★＝実測は −0.01 でも #NUM!');
    if (!/見当<0/.test(s)) throw new Error('★見当の 境目が 0 に なって いない★');
    if (!/刻んで 打ちました/.test(s)) throw new Error('★刻んで 測った 断りが 消えた★');
    console.log('      … 境目は 0（刻んで 測った 断りも 残って いる）');
  });
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (fail) process.exit(1);
