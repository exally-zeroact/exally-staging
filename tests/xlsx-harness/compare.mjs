// compare.mjs — Exally(本番経路) と Excelの真値(golden) を突き合わせて、判定・レポート・赤/緑を出す。
//
//   node tests/xlsx-harness/compare.mjs                 … 検証してレポートを書く(赤なら exit 1)
//   node tests/xlsx-harness/compare.mjs --self-test     … ★わざと壊して、ちゃんと赤くなるかを3通り確かめる
//   node tests/xlsx-harness/compare.mjs --update-snapshot … 経路スナップショットを作り直す(内容を報告してから使う)
//
//   判定は4値: 一致 / 不一致(既知=台帳にあり) / 不一致(新規) / 未検証(その版のgoldenが無い)
//     ・新規の不一致 → 赤。既知は緑だが、レポートに必ず全部並ぶ(黙って消さない)。
//     ・goldenが無い版は「未検証」。★未検証を緑とは呼ばない(件数を別に数えて出す)。
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { loadCases, runProductionPath, runRawHF } from './run-exally.mjs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const GOLDEN_DIR = path.join(__dirname, 'golden');
const KNOWN_PATH = path.join(__dirname, 'known-diffs.json');
const SNAP_PATH = path.join(__dirname, 'route-snapshot.json');
const REPORT_PATH = path.join(__dirname, 'report.md');

/* ══ 値の正規化と比較 ═══════════════════════════════════════════
 *  ★Excel側の型で比べ方を変える。
 *    文字列同士を数値に読み替えて比べると "0007" と 7 が同じ物になってしまうため、
 *    Excelが文字列(t:'s')なら文字列として厳密一致で見る。
 */
const NUM_TOL = 1e-9;
function excelDisplay(g) {
  if (!g) return '(なし)';
  if (g.t === 'b') return g.v ? 'TRUE' : 'FALSE';
  if (g.t === 'z') return '(空)';
  return String(g.v);
}
function exallyDisplay(a) {
  if (a === undefined) return '(未計算)';
  if (a === null) return '(null)';
  return String(a);
}
function matches(golden, actualRaw) {
  if (golden === undefined) return { ok: false, why: 'goldenなし' };
  if (actualRaw === undefined || actualRaw === null) return { ok: false, why: 'Exally側が値を返していない' };
  const actual = String(actualRaw);
  switch (golden.t) {
    case 'n': {
      const n = Number(actual);
      if (actual.trim() === '' || Number.isNaN(n)) return { ok: false, why: '数値でない' };
      const g = Number(golden.v);
      const diff = Math.abs(n - g);
      const ok = diff <= NUM_TOL || diff <= Math.abs(g) * NUM_TOL;
      return { ok, why: ok ? '' : '数値が違う' };
    }
    case 's': {
      const ok = actual === String(golden.v);           // ★数値に読み替えない
      return { ok, why: ok ? '' : '文字列が違う' };
    }
    case 'b': {
      const ok = actual.toUpperCase() === (golden.v ? 'TRUE' : 'FALSE');
      return { ok, why: ok ? '' : '論理値が違う' };
    }
    case 'e': {
      const ok = actual === String(golden.v);
      return { ok, why: ok ? '' : 'エラーの種類が違う' };
    }
    case 'z': {
      const ok = actual === '' || actual === '0';
      return { ok, why: ok ? '' : '空セルにならない' };
    }
    default: return { ok: false, why: 'goldenの型が不明: ' + golden.t };
  }
}
function rawHfDisplay(r) {
  if (!r) return '(なし)';
  if (r.v === null) return '(null)';
  if (typeof r.v === 'boolean') return r.v ? 'TRUE' : 'FALSE';
  return String(r.v);
}

/* ══ known-diffs の検査 ═══════════════════════════════════════ */
function checkKnownDiffs(known, todayStr) {
  const errs = [];
  const byId = {};
  for (const d of known.diffs) {
    if (!d.id) { errs.push('idが無い項目がある'); continue; }
    if (byId[d.id]) errs.push(`既知台帳にIDの重複: ${d.id}`);
    byId[d.id] = d;
    if (!['A', 'B', 'C'].includes(d.class)) errs.push(`${d.id}: class は A/B/C のいずれか (今: ${d.class})`);
    if (d.class === 'A') {
      if (!d.due) errs.push(`★${d.id}: 区分A は期限(due)が必須。空欄は認めない`);
      else if (!/^\d{4}-\d{2}-\d{2}$/.test(d.due)) errs.push(`${d.id}: due の形式は YYYY-MM-DD (今: ${d.due})`);
      else if (d.due < todayStr) errs.push(`★${d.id}: 期限切れ (due ${d.due} < 今日 ${todayStr})。直すか期限を引き直すこと`);
    }
  }
  return { errs, byId };
}

/* ══ 揮発性(TODAY/NOW) ═══════════════════════════════════════ */
function todaySerial(now = new Date()) {
  const utc = Date.UTC(now.getFullYear(), now.getMonth(), now.getDate());
  return Math.round((utc - Date.UTC(1899, 11, 30)) / 86400000);
}
function judgeVolatile(cs, actual, now) {
  const a = actual === null || actual === undefined ? '' : String(actual);
  if (cs.volatileCheck === 'todaySerial') {
    const want = String(todaySerial(now));
    return { ok: a === want, want, got: a, note: '実行時点の日付シリアルと一致するか' };
  }
  if (cs.volatileCheck === 'zero') {
    return { ok: Number(a) === 0, want: '0', got: a, note: 'NOWの整数部がTODAYと一致するか' };
  }
  return { ok: false, want: '(未定義)', got: a, note: 'volatileCheck が未指定' };
}

/* ══ 判定本体(純関数。self-test はここへ細工したデータを渡す) ═══ */
export function evaluate({ cases, golden, libre, exally, raw, known, snapshot, spy, entry, split, inputFidelity, now }) {
  const todayStr = new Date(now).toISOString().slice(0, 10);
  const { errs: knownErrs, byId: knownById } = checkKnownDiffs(known, todayStr);

  const rows = [];
  const counts = { 一致: 0, '不一致(既知)': 0, '不一致(新規)': 0, 未検証: 0, 揮発性: 0 };
  const volatileRows = [];

  for (const cs of cases) {
    const act = exally?.[cs.id]?.v;
    if (cs.volatile) {
      const v = judgeVolatile(cs, act, new Date(now));
      counts.揮発性++;
      volatileRows.push({ ...cs, ...v });
      continue;
    }
    const g = golden.cases[cs.id];
    const lo = libre ? libre.cases[cs.id] : undefined;
    let verdict, cls = '', note = '';
    if (!g) { verdict = '未検証'; counts.未検証++; }
    else {
      const m = matches(g, act);
      if (m.ok) { verdict = '一致'; counts.一致++; }
      else {
        const k = knownById[cs.id];
        if (k) { verdict = '不一致(既知)'; cls = k.class; note = `${k.note}${k.due ? ' / 期限 ' + k.due : ''}`; counts['不一致(既知)']++; }
        else { verdict = '不一致(新規)'; note = m.why; counts['不一致(新規)']++; }
      }
    }
    //  ★劣化検知: 生HFなら合っていたのに、独自層を通したら合わなくなった＝独自層が結果を悪くしている。
    //    実際に TEXT の日付書式で起きた。独自層に関数を足す時はここが鳴らないことを必ず確認する。
    const rawOk = g ? matches(g, raw?.[cs.id]?.v).ok : false;
    const prodOk = verdict === '一致';
    rows.push({
      ...cs, verdict, cls, note,
      exally: exallyDisplay(act),
      excel: excelDisplay(g),
      libre: lo ? excelDisplay(lo) : '未検証',
      rawhf: rawHfDisplay(raw?.[cs.id]),
      route: exally?.[cs.id]?.route || {},
      degraded: rawOk && !prodOk
    });
  }

  /* ── 経路の固定 ── */
  const routeErrs = [];
  //  ★実際に計算したケースだけで見る。計算していないケース(goldenが無くて追加されただけ等)を
  //    混ぜると、経路が変わっていないのに錠が鳴ってしまう。
  const computed = rows.filter(r => exally?.[r.id] !== undefined);
  const jsAnswered = computed.filter(r => r.route?.jsAnswered).map(r => r.id).sort();
  const rawDiffers = computed.filter(r => r.rawhf !== r.exally).map(r => r.id).sort();
  if (snapshot) {
    const cmp = (name, got, want) => {
      const g = JSON.stringify(got), w = JSON.stringify(want);
      if (g !== w) {
        const added = got.filter(x => !want.includes(x)), gone = want.filter(x => !got.includes(x));
        routeErrs.push(`★経路が変わった(${name}): 増えた=[${added.join(',')}] 消えた=[${gone.join(',')}]`);
      }
    };
    cmp('独自層が答えたケース', jsAnswered, snapshot.jsAnswered);
    cmp('生HFと本番経路で答えが違うケース', rawDiffers, snapshot.rawHfDiffers);
    if (entry && (entry.jsSetCount !== snapshot.entry.jsSetCount || entry.entryPoints !== snapshot.entry.entryPoints)) {
      routeErrs.push(`★独自層の入口の数が変わった: 今=${JSON.stringify(entry)} 期待=${JSON.stringify(snapshot.entry)}`);
    }
    if (rawDiffers.length === 0) routeErrs.push('★生HFと本番経路で答えが1件も違わない=うちの関数が1つも効いていない疑い');
    if (spy && spy.js === 0) routeErrs.push('★_jsComputeFormula が一度も呼ばれていない=本番経路を通っていない');
  }
  //  ★1つの関数は1箇所でだけ定義する: HFプラグインと _jsSet の両方に居たら赤
  if (split) {
    if (!entry?.pluginRegistered) routeErrs.push('★HFの関数プラグインが登録されていない(登録は buildEmpty より前に済んでいる必要がある)');
    const both = split.plugin.filter(n => split.jsSet.includes(n));
    if (both.length) routeErrs.push(`★同じ関数が2箇所で定義されている(プラグインと_jsSetの両方): ${both.join(',')}`);
  }

  /* ── ★独自層が生HFより悪くしていないか(常設チェック) ──
   *  独自層に関数を足すと、HFなら合っていた物を壊すことがある(TEXTの日付書式で実際に起きた)。
   *  台帳に degradation:true で登録されている物だけ既知として通し、それ以外は赤。 */
  const degradedRows = rows.filter(r => r.degraded);
  const degradedErrs = degradedRows
    .filter(r => !(knownById[r.id] && knownById[r.id].degradation))
    .map(r => `★独自層が生HFより悪くしている: ${r.id} 生HF=${r.rawhf} → 本番経路=${r.exally} (Excel真値=${r.excel})`);

  /* ── 揮発性の失敗 ── */
  const volatileErrs = volatileRows.filter(v => !v.ok).map(v => `★揮発性の確認に失敗: ${v.id} 期待=${v.want} 実際=${v.got}`);

  /* ── 入力の型(別枠。台帳に載っていれば既知) ── */
  const inputRows = Object.entries(inputFidelity || {}).map(([id, v]) => {
    const g = golden.inputProbes?.[id];
    const k = knownById[id];
    let verdict;
    if (!g) verdict = '未検証';
    else if (matches(g, v.got).ok) verdict = '一致';
    else verdict = k ? '不一致(既知)' : '不一致(新規)';
    return { id, ...v, excel: excelDisplay(g), excelType: g?.t || '-', verdict, cls: k?.class || '', note: k?.note || '' };
  });
  const inputNew = inputRows.filter(r => r.verdict === '不一致(新規)').length;

  const hardErrs = [...knownErrs, ...routeErrs, ...volatileErrs, ...degradedErrs];
  const exitCode = (counts['不一致(新規)'] > 0 || inputNew > 0 || hardErrs.length > 0) ? 1 : 0;
  return { rows, counts, volatileRows, inputRows, hardErrs, knownErrs, routeErrs, volatileErrs, degradedErrs, degradedRows, exitCode, jsAnswered, rawDiffers };
}

/* ══ レポート ═══════════════════════════════════════════════ */
function buildReport(res, meta) {
  const L = [];
  const esc = s => String(s).replace(/\|/g, '\\|');
  L.push('# 版対応 検証ハーネス レポート');
  L.push('');
  L.push(`- 生成日: ${meta.date}`);
  L.push(`- 真値: **${meta.golden.product} ${meta.golden.version}** (${meta.golden.platform} / ${meta.golden.updateChannel?.includes('492350f6') ? 'Current Channel' : meta.golden.updateChannel})`);
  L.push(`- ロケール: UI=${meta.golden.uiLanguageId} / 国=${meta.golden.countrySetting} / 小数点='${meta.golden.decimalSep}' 桁区切り='${meta.golden.thousandsSep}' / 日付システム=${meta.golden.date1904 ? '1904' : '1900'}`);
  L.push(`- 計算経路: **book.html の setCellFormula(本番と同じ)**。生の HyperFormula ではない。`);
  L.push('');
  L.push('> **★これは「今この瞬間の本番グリッドのお金バグ」である。** 下の「不一致(既知)」のうち money_impact の付いた項目は、');
  L.push('> 金額の端数・表示に直接効く。ハーネスを先に作る判断のため **このコミットでは直していない**（区分Aとして期限付きで台帳に登録）。');
  L.push('');
  L.push('## 集計');
  L.push('');
  L.push('| 判定 | 件数 |');
  L.push('|---|---|');
  for (const [k, v] of Object.entries(res.counts)) L.push(`| ${k} | ${v} |`);
  L.push(`| **合計** | ${Object.values(res.counts).reduce((a, b) => a + b, 0)} |`);
  L.push('');
  L.push('※ 「未検証」は緑ではない。その版の真値がまだ無い、という意味。');
  L.push('');

  L.push('## 版ごとの状態');
  L.push('');
  L.push('| 版 | 状態 |');
  L.push('|---|---|');
  L.push(`| Excel 365 (${meta.golden.version}) | **真値**（このリポジトリの基準） |`);
  L.push(`| LibreOffice | ${meta.libre ? `参考値あり (${meta.libre.version})` : '**未検証**（goldenが無い。CIの別ジョブで生成する）'} |`);
  L.push('| Excel 2016 / 2019 / Mac | **未検証**（実機が無い。golden/RECIPE.md の手順でその環境で1回走らせれば埋まる） |');
  L.push('');

  if (res.volatileRows.length) {
    L.push('## 揮発性の関数（別扱い）');
    L.push('');
    L.push('TODAY / NOW は毎回答えが変わるため **golden突合の対象外**。固定値と比べると必ず腐るので、実行時点との一致だけを見る。');
    L.push('');
    L.push('| ケース | 式 | 見方 | 期待 | 実際 | 判定 |');
    L.push('|---|---|---|---|---|---|');
    for (const v of res.volatileRows) L.push(`| ${v.id} | \`${esc(v.f)}\` | ${v.note} | ${v.want} | ${v.got} | ${v.ok ? 'OK' : '★NG'} |`);
    L.push('');
  }

  const news = res.rows.filter(r => r.verdict === '不一致(新規)');
  L.push('## 不一致（新規）＝赤');
  L.push('');
  if (!news.length) L.push('なし。');
  else {
    L.push('| 関数 | ケース | 式 | Exally | Excel真値 | 理由 |');
    L.push('|---|---|---|---|---|---|');
    for (const r of news) L.push(`| ${r.func} | ${r.id} | \`${esc(r.f)}\` | ${esc(r.exally)} | ${esc(r.excel)} | ${r.note} |`);
  }
  L.push('');

  const knowns = res.rows.filter(r => r.verdict === '不一致(既知)');
  L.push('## 不一致（既知＝台帳にあり・緑だが必ず全件出す）');
  L.push('');
  if (!knowns.length) L.push('なし。');
  else {
    L.push('| 区分 | 関数 | ケース | 式 | Exally | Excel真値 | 中身と期限 |');
    L.push('|---|---|---|---|---|---|---|');
    for (const r of knowns) L.push(`| ${r.cls} | ${r.func} | ${r.id} | \`${esc(r.f)}\` | ${esc(r.exally)} | ${esc(r.excel)} | ${esc(r.note)} |`);
  }
  L.push('');

  L.push('## 入力の型が保たれるか（別枠）');
  L.push('');
  L.push('セルに打ち込んだ文字をどう解釈するか。関数の検証とは別の話なので、混ぜずにここで見る。');
  L.push('期待値は想像ではなく、**実Excelの標準書式セルに同じ文字を打ち込んだ実測値**。');
  L.push('');
  L.push('| ケース | 打ち込んだ値 | Exally | Excel(標準書式セル) | Excel型 | 判定 | 中身 |');
  L.push('|---|---|---|---|---|---|---|');
  for (const r of res.inputRows) L.push(`| ${r.id} | \`${esc(r.raw)}\` | ${esc(r.got)} | ${esc(r.excel)} | ${r.excelType} | ${r.verdict} | ${esc(r.note)} |`);
  L.push('');

  L.push('## ★独自層が生HFより悪くしていないか');
  L.push('');
  L.push('独自層に関数を足すと、HyperFormula なら合っていた物を壊すことがある（TEXTの日付書式で実際に起きた）。');
  L.push('**独自層に関数を足す時は、必ずここが増えていないことを確認してから足す。** 台帳に載っていない劣化は赤。');
  L.push('');
  if (!res.degradedRows.length) L.push('劣化なし。');
  else {
    L.push('| ケース | 生HF | 本番経路(独自層あり) | Excel真値 | 状態 |');
    L.push('|---|---|---|---|---|');
    for (const r of res.degradedRows) {
      const k = r.cls ? '既知(台帳)' : '★新規';
      L.push(`| ${r.id} | ${esc(r.rawhf)} | ${esc(r.exally)} | ${esc(r.excel)} | ${k} |`);
    }
  }
  L.push('');
  L.push('## 経路の固定（将来 生HF に落ちたら気付くための錠）');
  L.push('');
  L.push(`- 独自層(_jsComputeFormula)が答えたケース: **${res.jsAnswered.length}件**`);
  L.push(`- 生HFと本番経路で答えが違うケース: **${res.rawDiffers.length}件** … この差が消えたら「素通りに落ちた」ということ`);
  L.push(`- 独自層の入口: ${JSON.stringify(meta.entry)}（1つだけであること）`);
  if (res.routeErrs.length) { L.push(''); res.routeErrs.forEach(e => L.push(`- ${e}`)); }
  L.push('');

  L.push('## 全ケース');
  L.push('');
  L.push('| 関数 | ケース | 式 | Exally(本番経路) | Excel365(真値) | LibreOffice | 生HF | 判定 | 区分 |');
  L.push('|---|---|---|---|---|---|---|---|---|');
  for (const r of res.rows) {
    L.push(`| ${r.func} | ${r.id} | \`${esc(r.f)}\` | ${esc(r.exally)} | ${esc(r.excel)} | ${esc(r.libre)} | ${esc(r.rawhf)} | ${r.verdict} | ${r.cls} |`);
  }
  L.push('');
  return L.join('\n');
}

/* ══ 実行 ═══════════════════════════════════════════════════ */
function loadGolden() {
  const files = fs.readdirSync(GOLDEN_DIR).filter(f => f.endsWith('.json') && !f.startsWith('pending-'));
  const excel = files.find(f => f.startsWith('excel-'));
  const libreFile = files.find(f => f.startsWith('libreoffice-'));
  if (!excel) throw new Error('golden/excel-*.json が無い。真値が無いので検証できない(tools/golden-excel.ps1 -Init)');
  return {
    golden: JSON.parse(fs.readFileSync(path.join(GOLDEN_DIR, excel), 'utf8')),
    libre: libreFile ? JSON.parse(fs.readFileSync(path.join(GOLDEN_DIR, libreFile), 'utf8')) : null
  };
}

async function main() {
  const args = process.argv.slice(2);
  const selfTest = args.includes('--self-test');
  const updateSnap = args.includes('--update-snapshot');

  const { inputs, cases } = loadCases();
  const { golden, libre } = loadGolden();
  const known = JSON.parse(fs.readFileSync(KNOWN_PATH, 'utf8'));
  const snapshot = fs.existsSync(SNAP_PATH) ? JSON.parse(fs.readFileSync(SNAP_PATH, 'utf8')) : null;

  const prod = await runProductionPath(inputs, cases);
  if (prod.skipped) { console.log('SKIP: ' + prod.skipped); process.exit(1); }
  const raw = runRawHF(inputs, cases);
  const now = Date.now();
  const base = { cases, golden, libre, exally: prod.out, raw, known, snapshot, spy: prod.spyTotals, entry: prod.entry, split: prod.split, inputFidelity: prod.inputFidelity, now };

  if (updateSnap) {
    const r = evaluate({ ...base, snapshot: null });
    const snap = {
      note: '★これは錠。book.html が変わって独自層を通らなくなったら、この一覧が変わって赤くなる。中身を確認してから更新すること。',
      entry: prod.entry, jsAnswered: r.jsAnswered, rawHfDiffers: r.rawDiffers
    };
    fs.writeFileSync(SNAP_PATH, JSON.stringify(snap, null, 1) + '\n');
    console.log(`route-snapshot.json を更新: 独自層が答えた ${r.jsAnswered.length}件 / 生HFと違う ${r.rawDiffers.length}件`);
    return 0;
  }

  if (selfTest) return runSelfTest(base);

  const res = evaluate(base);
  const meta = { date: new Date(now).toISOString().slice(0, 10), golden: golden.meta, libre: libre?.meta, entry: prod.entry };
  fs.writeFileSync(REPORT_PATH, buildReport(res, meta));

  console.log('── 版対応 検証ハーネス ─────────────────────────');
  for (const [k, v] of Object.entries(res.counts)) console.log(`  ${k.padEnd(14)} ${v}`);
  console.log(`  レポート: tests/xlsx-harness/report.md`);
  if (prod.domErrors?.length) console.log(`  (book.html 読み込み時のDOM例外: ${prod.domErrors.length}件 — 計算経路には影響なし)`);
  for (const e of res.hardErrs) console.log('  ' + e);
  const news = res.rows.filter(r => r.verdict === '不一致(新規)');
  for (const r of news) console.log(`  ★新規不一致: ${r.id}  Exally=${r.exally}  Excel=${r.excel}`);
  console.log(res.exitCode ? '★赤(exit 1)' : '緑(新規の不一致なし。既知は台帳で全件明示)');
  return res.exitCode;
}

/* ══ --self-test: わざと壊して赤くなるか ══════════════════════ */
function runSelfTest(base) {
  const clone = o => JSON.parse(JSON.stringify(o));
  let ng = 0;
  const line = s => console.log(s);
  line('══ --self-test: わざと壊して、ちゃんと気付くかを見る ══');

  // 0) まず素の状態が緑であること
  const r0 = evaluate(base);
  base.rows0 = r0.rows;
  line(`\n[0] 素の状態      : exit=${r0.exitCode} 新規不一致=${r0.counts['不一致(新規)']} 既知=${r0.counts['不一致(既知)']} 未検証=${r0.counts.未検証}`);
  if (r0.exitCode !== 0) { line('    ★NG: 壊す前から赤い'); ng++; } else line('    OK: 壊す前は緑');

  // 1) golden を1件わざと書き換える → 新規不一致で赤
  const g1 = clone(base.golden);
  const victim = base.cases.find(c => !c.volatile && g1.cases[c.id] && g1.cases[c.id].t === 'n').id;
  g1.cases[victim].v = Number(g1.cases[victim].v) + 12345;
  const r1 = evaluate({ ...base, golden: g1 });
  const hit1 = r1.rows.find(r => r.id === victim);
  line(`\n[1] goldenを1件改ざん (${victim} → ${g1.cases[victim].v})`);
  line(`    exit=${r1.exitCode} 新規不一致=${r1.counts['不一致(新規)']} 判定=${hit1?.verdict}`);
  if (r1.exitCode === 1 && r1.counts['不一致(新規)'] === 1 && hit1?.verdict === '不一致(新規)') line('    OK: 赤くなった');
  else { line('    ★NG: 赤くならない'); ng++; }

  // 2) known-diffs から1件消す → 既知が新規不一致に昇格して赤
  const knownA = base.known.diffs.find(d => base.cases.some(c => c.id === d.id));
  const k2 = clone(base.known);
  k2.diffs = k2.diffs.filter(d => d.id !== knownA.id);
  const r2 = evaluate({ ...base, known: k2 });
  const hit2 = r2.rows.find(r => r.id === knownA.id);
  line(`\n[2] known-diffs から1件削除 (${knownA.id})`);
  line(`    exit=${r2.exitCode} 新規不一致=${r2.counts['不一致(新規)']} 判定=${hit2?.verdict}`);
  if (r2.exitCode === 1 && hit2?.verdict === '不一致(新規)') line('    OK: 既知が新規不一致に昇格して赤くなった');
  else { line('    ★NG: 昇格しない'); ng++; }

  // 3) goldenの無いケースを1件足す → 「未検証」と出て、一致に混ざらず、緑のまま
  const c3 = base.cases.concat([{ id: 'SELFTEST_NO_GOLDEN', func: 'SUM', f: '=SUM(E1:E2)', group: 'self-test' }]);
  const r3 = evaluate({ ...base, cases: c3 });
  const hit3 = r3.rows.find(r => r.id === 'SELFTEST_NO_GOLDEN');
  line(`\n[3] goldenが無いケースを1件追加 (SELFTEST_NO_GOLDEN)`);
  line(`    exit=${r3.exitCode} 未検証=${r3.counts.未検証}(素の状態=${r0.counts.未検証}) 一致=${r3.counts.一致}(素の状態=${r0.counts.一致}) 判定=${hit3?.verdict}`);
  if (r3.exitCode === 0 && hit3?.verdict === '未検証' && r3.counts.未検証 === r0.counts.未検証 + 1 && r3.counts.一致 === r0.counts.一致) {
    line('    OK: 「未検証」として別に数えられ、一致には混ざらず、緑のまま');
  } else { line('    ★NG: 未検証の扱いがおかしい'); ng++; }

  // 4) 区分Aの期限を空にする → 台帳の検査で赤
  const k4 = clone(base.known);
  const a4 = k4.diffs.find(d => d.class === 'A');
  delete a4.due;
  const r4 = evaluate({ ...base, known: k4 });
  line(`\n[おまけ] 区分Aの期限(due)を空にする (${a4.id})`);
  line(`    exit=${r4.exitCode} / ${r4.knownErrs[0] || '(検出できず)'}`);
  if (r4.exitCode === 1 && r4.knownErrs.some(e => e.includes(a4.id))) line('    OK: 期限が空だと赤くなった');
  else { line('    ★NG: 期限の空欄を見逃した'); ng++; }

  // 5) ★独自層が生HFより悪くしている状態を作る → 台帳に載っていても赤
  const victim5 = base.rows0.find(r => r.verdict === '一致' && r.rawhf === r.exally && !r.volatile);
  const e5 = clone(base.exally);
  e5[victim5.id].v = 'こわした値';
  const k5 = clone(base.known);
  k5.diffs.push({ id: victim5.id, func: victim5.func, root: 'self-test', class: 'C', note: 'self-test(劣化印は付けない)' });
  const r5 = evaluate({ ...base, exally: e5, known: k5 });
  const hit5 = r5.rows.find(r => r.id === victim5.id);
  line(`\n[5] 生HFなら合う物を独自層で壊す (${victim5.id} → こわした値)。台帳には載せるが劣化印は付けない`);
  line(`    exit=${r5.exitCode} 新規不一致=${r5.counts['不一致(新規)']} 判定=${hit5?.verdict} / ${r5.degradedErrs[0] || '(検出できず)'}`);
  if (r5.exitCode === 1 && r5.counts['不一致(新規)'] === 0 && r5.degradedErrs.some(e => e.includes(victim5.id))) {
    line('    OK: 台帳に載っていても「独自層が悪くしている」として赤くなった');
  } else { line('    ★NG: 劣化を見逃した'); ng++; }

  // 6) 同じ関数が HFプラグインと _jsSet の両方に居る状態を作る → 赤
  const sp6 = { plugin: base.split.plugin.slice(), jsSet: base.split.jsSet.slice() };
  sp6.jsSet.push(sp6.plugin[0]);
  const r6 = evaluate({ ...base, split: sp6 });
  line(`\n[6] 同じ関数を2箇所に登録した状態にする (${sp6.plugin[0]} を _jsSet にも入れる)`);
  line(`    exit=${r6.exitCode} / ${r6.routeErrs.find(e => e.includes('2箇所')) || '(検出できず)'}`);
  if (r6.exitCode === 1 && r6.routeErrs.some(e => e.includes('2箇所'))) line('    OK: 二重定義を赤にできた');
  else { line('    ★NG: 二重定義を見逃した'); ng++; }

  line(`\n══ self-test: ${ng ? '★' + ng + '件 失敗' : '6通りとも期待どおり'} ══`);
  return ng ? 1 : 0;
}

if (process.argv[1] && pathToFileURL(process.argv[1]).href === import.meta.url) {
  process.exit(await main());
}
