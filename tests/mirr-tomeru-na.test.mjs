/* mirr-tomeru-na.test.mjs — ★MIRR が 実Excel の 答える 所で 断って いないか★（2026-09-09）
 *
 *  ★★これは「断る」直しでは ありません＝「止めて いたのを 動かす」直しです★★
 *    ★エンジン（借り物）が 実Excel の 答える 所で #DIV/0! を 返して いました★
 *      =MIRR(全部マイナス,0.1,0.12) 実Excel ★−1★ ／ うち ★#DIV/0!★
 *      =MIRR(A1:A3,-1,0.12)        実Excel ★0.17132403714770583★ ／ うち ★#DIV/0!★
 *    ⇒★★出来る 物を 止めて いた＝方針（Excel の 最上級・全部 つける）に 真っ向から 反する★★
 *    ⇒ 同じ 日に 直した XIRR/IRR/XNPV は ★逆向き★（数を 返す 所を 断る）＝★別の PR★
 *
 *  ★★式は 実Excel の 定義から 自分で 書きました（★借り物は 読んで いません★）★★
 *    MIRR ＝ ( −(正の 流れを 回す利率で 最後まで 育てた 合計)
 *              ÷ (負の 流れを 借りる利率で 今に 引き戻した 合計) )^(1/(n−1)) − 1
 *    ⇒★手で 検算して 実Excel と 6通り 一致★してから 入れました
 *
 *  ★★何を 守るか（1つずつ 名指し）★★
 *    ①★実Excel が 答える 所で 断らない★（全部マイナス／借りる利率 −1）
 *    ②★実Excel が 断る 所は 断る★（全部プラス／全部0／1件だけ／回す利率 −1）
 *    ③★空・字は 落として 計算する★（★断らない★）
 *    ④★借りる 利率は 答えに 効かない★（−1／−0.5／−0.01／0／0.01 で 同じ）
 *
 *  ★★見て いない 範囲（★書かない 見張りは「全部 守った」と 読まれる★）★★
 *    ・★負の 流れが 2つ 以上 在る 組は 測って いません★
 *      （測った 組は 負が 1件・0期＝だから 借りる 利率が 効かなかった）
 *      ⇒★「借りる 利率は 効かない」は ★この 組での 話★です★
 *    ・IRR / XIRR / XNPV は ★別の 見張り★（tests/xirr-mon.test.mjs）
 *    ・答えの 細かい 桁（幅 1e-9）
 *
 *  使い方: node tests/mirr-tomeru-na.test.mjs [--self-test]
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

const 式の行 = 30;
/* ★★本番と 同じ 道★★ ①JS層 → ②convertFormula → エンジン */
function 押す(値ら, 借, 回) {
  const 板 = 値ら.map((x) => [x, null]);
  板.push([null, null]);
  const 式 = '=MIRR(A1:A' + 値ら.length + ',' + 借 + ',' + 回 + ')';
  hf.setSheetContent(SID, 板.map((r) => r.slice()));
  const js = EF._jsComputeFormula(0, 式);
  if (js !== null && js !== undefined) return { 式, 答: String(js) };
  const 盤 = 板.map((r) => r.slice());
  while (盤.length <= 式の行) 盤.push([null, null]);
  盤[式の行][0] = EF.convertFormula(式);
  hf.setSheetContent(SID, 盤);
  const v = hf.getCellValue({ sheet: SID, row: 式の行, col: 0 });
  if (v && v.type) return { 式, 答: '#' + v.type };
  return { 式, 答: v === null || v === undefined ? '(空)' : String(v) };
}
const 近い = (う, 正) => {
  const a = Number(う), b = Number(正);
  if (!isFinite(a) || !isFinite(b)) return false;
  if (b === 0) return Math.abs(a) < 1e-12;
  return Math.abs((a - b) / b) <= 1e-9;
};

console.log('\n[mirr-tomeru-na] ★MIRR が 実Excel の 答える 所で 断って いないか★');

T('★★実Excel が 答える 所で 断って いない（★一番 大事★）★★', () => {
  const 組 = [
    [[-100, -200, -300], '0.1', '0.12', -1, '★全部 マイナス（前は #DIV/0!）★'],
    [[-1000, 600, 700], '-1', '0.12', 0.17132403714770583, '★借りる 利率 −1（前は #DIV/0!）★'],
  ];
  for (const [v, 借, 回, 正, 札] of 組) {
    const r = 押す(v, 借, 回);
    if (typeof r.答 === 'string' && r.答[0] === '#') {
      throw new Error('★' + 札 + '（' + r.式 + '）を 断って いる（' + r.答 + '）★'
        + ' ⇒★実Excel は ' + 正 + ' を 返します＝★出来る 物を 止めた★');
    }
    if (!近い(r.答, 正)) throw new Error('★' + 札 + ' … うち ' + r.答 + ' ／ 実Excel ' + 正 + '★');
  }
  console.log('      … ' + 組.length + '通り とも 値を 返す（★前は どちらも #DIV/0!★）');
});

T('★実Excel が 断る 所は 断る★', () => {
  const 組 = [
    [[100, 200, 300], '0.1', '0.12', '★全部 プラス★'],
    [[0, 0, 0], '0.1', '0.12', '★全部 0★'],
    [[-1000], '0.1', '0.12', '★1件だけ★'],
    [[-1000, 600, 700], '0.1', '-1', '★回す 利率 −1（1+利率 が 0）★'],
  ];
  for (const [v, 借, 回, 札] of 組) {
    const r = 押す(v, 借, 回);
    if (r.答 !== '#DIV/0!') throw new Error('★' + 札 + '（' + r.式 + '）… うち ' + r.答 + ' ／ 実Excel #DIV/0!★');
  }
  console.log('      … ' + 組.length + '通り とも #DIV/0!');
});

T('★空・字は 落として 計算する（★断らない★）★', () => {
  for (const [v, 札] of [[[-1000, null, 1100], '★空が 混ざる★'], [[-1000, 'あ', 1100], '★字が 混ざる★']]) {
    const r = 押す(v, '0.1', '0.12');
    if (!近い(r.答, 0.10000000000000009)) {
      throw new Error('★' + 札 + ' … うち ' + r.答 + ' ／ 実Excel 0.10000000000000009★');
    }
  }
  console.log('      … 空も 字も 落として 計算（実Excel と 同じ）');
});

T('★借りる 利率は 答えに 効かない（★この 組では★）★', () => {
  /* ★負の 流れが 1件・0期だから 効きません★
     ⇒★負が 2つ 以上の 組は 測って いません＝この 話は この 組限り★ */
  const 正 = 0.17132403714770583;
  for (const 借 of ['-1', '-0.5', '-0.01', '0', '0.01']) {
    const r = 押す([-1000, 600, 700], 借, '0.12');
    if (!近い(r.答, 正)) throw new Error('★借りる 利率 ' + 借 + ' … うち ' + r.答 + ' ／ 実Excel ' + 正 + '★');
  }
  console.log('      … 5通り とも 同じ 答え（★測った 組は 負が 1件・0期★）');
});

T('★回す 利率で 答えが 変わる（★効いて いる★）★', () => {
  const 組 = [['-0.5', 0], ['-0.01', 0.13754120804478998], ['0', 0.14017542509913805],
    ['0.01', 0.14280357017293221], ['0.12', 0.17132403714770583]];
  for (const [回, 正] of 組) {
    const r = 押す([-1000, 600, 700], '0.1', 回);
    if (!近い(r.答, 正)) throw new Error('★回す 利率 ' + 回 + ' … うち ' + r.答 + ' ／ 実Excel ' + 正 + '★');
  }
  console.log('      … 5通り とも 実Excel と 合う（★答えが ちゃんと 動く★）');
});

T('★紙（実測）と 試験の 答えが 同じ★', () => {
  const p = path.join(ROOT, 'docs/measured/golden-okane-4kansuu-2026-09-09.tsv');
  if (!fs.existsSync(p)) throw new Error('★紙が 無い★ … ' + p);
  const 紙 = {};
  for (const l of fs.readFileSync(p, 'utf-8').split('\n')) {
    if (l.startsWith('#') || !l.trim()) continue;
    const c = l.split('\t');
    if (c.length >= 3 && c[0].trim() === 'MIRR') 紙[c[1].trim()] = c[2].trim();
  }
  const 見る = [['=MIRR(E1:E3,0.1,0.12)', '-1'],
    ['=MIRR(A1:A3,-1,0.12)', '0.17132403714770583'],
    ['=MIRR(C1:C3,0.1,0.12)', '#DIV/0!'],
    ['=MIRR(A1:A3,0.1,-1)', '#DIV/0!']];
  const 違う = 見る.filter(([k, v]) => 紙[k] !== v);
  if (違う.length) {
    throw new Error('★紙と 違う … ' + 違う.map(([k, v]) => k + '（紙 ' + 紙[k] + ' ／ 試験 ' + v + '）').join(' / ') + '★');
  }
  console.log('      … 紙と 同じ（' + 見る.length + '本を 名指しで 見た）');
});

/* ══ ★自己試験（★壊すのは 写し★＝ファイルは 1バイトも 触らない）★ ══ */
if (自己試験) {
  console.log('\n★自己試験（★壊すのは 写し★＝ファイルは 1バイトも 触らない）★');

  T('★★段が コードに 在る（空振りして いない）★★', () => {
    const s = fs.readFileSync(path.join(ROOT, 'exally-formula.js'), 'utf-8');
    if (!/var mMirr=fOrig\.match/.test(s)) throw new Error('★MIRR の 段が 無い★');
    if (!/\bMIRR:1\b/.test(s)) throw new Error('★関所に MIRR が 無い＝★私の 段まで 届かない★');
    console.log('      … 段と 関所の 両方に 在る');
  });

  T('★★関所から 外したら エンジンが 断って しまう（＝段が 効いて いる 証拠）★★', () => {
    /* ★エンジンに 直に 聞く★＝JS層を 通さない */
    const 盤 = [[-100, null], [-200, null], [-300, null], [null, null]];
    while (盤.length <= 式の行) 盤.push([null, null]);
    盤[式の行][0] = EF.convertFormula('=MIRR(A1:A3,0.1,0.12)');
    hf.setSheetContent(SID, 盤);
    const v = hf.getCellValue({ sheet: SID, row: 式の行, col: 0 });
    const 字 = (v && v.type) ? ('#' + v.type) : String(v);
    if (!/DIV/.test(字)) {
      throw new Error('★エンジンが 断って いない（' + 字 + '）★＝★この 直しは 要らなかった 事に なる★');
    }
    console.log('      … エンジンに 直に 聞くと ' + 字 + '（★実Excel は −1★）');
  });

  T('★借り物の 中を 読んで いない（式は 定義から 自分で 書いた）★', () => {
    const s = fs.readFileSync(path.join(ROOT, 'exally-formula.js'), 'utf-8');
    if (!/実Excel の 定義から 自分で 書きました/.test(s)) {
      throw new Error('★「定義から 自分で 書いた」の 断りが 消えた★');
    }
    if (!/手で 検算して 実Excel と 6通り 一致/.test(s)) {
      throw new Error('★手で 検算した 断りが 消えた★');
    }
    console.log('      … 断り 2つとも 残って いる');
  });
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (fail) process.exit(1);
