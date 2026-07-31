// nesting-audit.mjs — ★「独自層は式の一番外側しか横取りしない」の被害範囲を実測する。
//
//   やり方: 同じ式を2通りで計算して比べる。
//     (外側)  =XLOOKUP(...)                       … 独自層が横取りできる形
//     (入れ子) =IF(1=1,XLOOKUP(...),XLOOKUP(...))  … 外側がIFなので独自層は横取りできない
//   Excelでは IF(1=1,X,X) は X と同じ。だから答えが変わるなら、それは
//   「独自層の救済が入れ子だと効いていない」＝表示だけ正しくて中身が違う、という事。
//
//   node tests/xlsx-harness/nesting-audit.mjs            … 既存ケースを全部この形で測る
//   node tests/xlsx-harness/nesting-audit.mjs --probe    … ケースが無い独自層関数も代表式で測る
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { loadCases, runProductionPath } from './run-exally.mjs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..', '..');

export function overrideNames() {
  const src = fs.readFileSync(path.join(ROOT, 'exally-formula.js'), 'utf8');
  const m = src.match(/var _jsSet = \{([\s\S]*?)\};/);
  if(!m) throw new Error('_jsSet が見つからない(独自層の対象表)');
  return [...m[1].matchAll(/([A-Z][A-Z0-9.]*)\s*:\s*1/g)].map(x => x[1]);
}
export function pluginNames() {
  const src = fs.readFileSync(path.join(ROOT, 'exally-formula.js'), 'utf8');
  const m = src.match(/var _PLUGIN_FUNCS\s*=\s*\[([\s\S]*?)\];/);
  if(!m) return [];
  return [...m[1].matchAll(/'([A-Z][A-Z0-9.]*)'/g)].map(x => x[1]);
}

// 一番外側の関数名(独自層が見ているのと同じ判定)
export function outerFunc(f) {
  const m = String(f).replace(/^=/, '').trim().match(/^([A-Z][A-Z0-9.]*)\s*\(/i);
  return m ? m[1].toUpperCase().split('.')[0] : null;
}
// IF(1=1,X,X) で包む = 値も型も変えずに「外側が別の関数」の形にする
export function wrap(f) {
  const body = String(f).replace(/^=/, '');
  return `=IF(1=1,${body},${body})`;
}

/* ケースの無い独自層関数のための代表式(入力セルは cases/_inputs.json と同じ配置) */
const PROBES = {
  NUMBERVALUE: '=NUMBERVALUE("1.5")', RANK: '=RANK(A1,E1:E6,0)', PERCENTILE: '=PERCENTILE(E1:E6,0.5)',
  QUARTILE: '=QUARTILE(E1:E6,2)', MODE: '=MODE(D1:D6)', TRIMMEAN: '=TRIMMEAN(E1:E6,0.2)',
  PERCENTRANK: '=PERCENTRANK(E1:E6,300)', KURT: '=KURT(E1:E6)', INTERCEPT: '=INTERCEPT(E1:E6,C1:C6)',
  FORECAST: '=FORECAST(30,E1:E6,C1:C6)', IRR: '=IRR(A1:A6)', DATEVALUE: '=DATEVALUE("2026-07-31")',
  INDIRECT: '=INDIRECT(B5)', OFFSET: '=OFFSET(E1,2,0)', XLOOKUP: '=XLOOKUP(20,C1:C6,E1:E6)',
  XMATCH: '=XMATCH(20,C1:C6)', LOOKUP: '=LOOKUP(20,C1:C6,E1:E6)', CONCAT: '=CONCAT(D1:D3)',
  FIXED: '=FIXED(A5,2)', DOLLAR: '=DOLLAR(A1,0)', YEN: '=YEN(A1,0)', N: '=N(A1)', TYPE: '=TYPE(A1)',
  ENCODEURL: '=ENCODEURL("a b")', GESTEP: '=GESTEP(A1,500)', TEXTBEFORE: '=TEXTBEFORE(B5,"-")',
  TEXTAFTER: '=TEXTAFTER(B5,"-")', VALUETOTEXT: '=VALUETOTEXT(A1)',
  AGGREGATE: '=AGGREGATE(9,0,E1:E6)', PERMUT: '=PERMUT(5,2)', PERMUTATIONA: '=PERMUTATIONA(5,2)',
  FREQUENCY: '=SUM(FREQUENCY(E1:E6,C1:C6))', MDETERM: '=MDETERM(C1:D2)',
  LENB: '=LENB(B1)', LEFTB: '=LEFTB(B1,2)', RIGHTB: '=RIGHTB(B1,2)', MIDB: '=MIDB(B1,1,2)',
  ASC: '=ASC(B3)', JIS: '=JIS(B3)',
  DSUM: '=DSUM(C1:E6,3,D1:D2)', DCOUNT: '=DCOUNT(C1:E6,3,D1:D2)'
};

async function main() {
  const withProbe = process.argv.includes('--probe');
  const { inputs, cases } = loadCases();
  const ov = overrideNames();
  const plug = pluginNames();

  // 測る式を組み立てる
  const items = [];
  for (const c of cases) {
    if (c.volatile) continue;
    items.push({ id: c.id, func: c.func, f: c.f, kind: 'case' });
  }
  if (withProbe) {
    const covered = new Set(cases.map(c => outerFunc(c.f)).filter(Boolean));
    for (const [name, f] of Object.entries(PROBES)) {
      if (covered.has(name)) continue;
      items.push({ id: 'PROBE_' + name, func: name, f, kind: 'probe' });
    }
  }

  // 外側の形 と 入れ子の形 を両方まとめて1回で計算する
  const pairs = [];
  for (const it of items) {
    pairs.push({ ...it, id: it.id + '::outer', f: it.f, of: it.id });
    pairs.push({ ...it, id: it.id + '::nested', f: wrap(it.f), of: it.id });
  }
  const res = await runProductionPath(inputs, pairs);
  if (res.skipped) { console.log('SKIP: ' + res.skipped); process.exit(1); }

  const rows = [];
  for (const it of items) {
    const a = res.out[it.id + '::outer']?.v;
    const b = res.out[it.id + '::nested']?.v;
    const outer = outerFunc(it.f);
    const inOverride = outer && ov.includes(outer);
    const inPlugin = outer && plug.includes(outer);
    rows.push({ ...it, outer, inOverride, inPlugin, top: a === null || a === undefined ? '(null)' : String(a), nested: b === null || b === undefined ? '(null)' : String(b), same: String(a) === String(b) });
  }

  const broken = rows.filter(r => !r.same);
  const brokenOv = broken.filter(r => r.inOverride);
  const byFunc = {};
  for (const r of broken) byFunc[r.outer || '(なし)'] = (byFunc[r.outer || '(なし)'] || 0) + 1;

  console.log('══ 入れ子で答えが変わるか(独自層の救済が消えるか) ══');
  console.log(`  独自層の対象表(_jsSet): ${ov.length}個 / HFプラグイン登録: ${plug.length}個`);
  console.log(`  測った式: ${rows.length}本 (ケース${rows.filter(r=>r.kind==='case').length} / 代表式${rows.filter(r=>r.kind==='probe').length})`);
  console.log(`  ★入れ子にすると答えが変わる: ${broken.length}本 (うち独自層の対象: ${brokenOv.length}本)`);
  console.log(`  関数別: ${Object.entries(byFunc).map(([k,v])=>k+'×'+v).join(' ') || 'なし'}`);
  if (broken.length) {
    console.log('\n  ケース             外側の答え        入れ子の答え');
    console.log('  ' + '-'.repeat(64));
    for (const r of broken.slice(0, 40)) {
      console.log('  ' + String(r.id).replace(/::.*$/,'').padEnd(20) + String(r.top).slice(0,16).padEnd(18) + String(r.nested).slice(0,20));
    }
    if (broken.length > 40) console.log(`  ...他 ${broken.length-40} 本`);
  }

  // ★baseline を超えたら赤(黙って増やせないようにする)
  const basePath = path.join(__dirname, 'nesting-baseline.json');
  let exitCode = 0;
  if (process.argv.includes('--check')) {
    if (!fs.existsSync(basePath)) { console.log('  ★nesting-baseline.json が無い。--update-baseline で作る'); exitCode = 1; }
    else {
      const base = JSON.parse(fs.readFileSync(basePath, 'utf8'));
      if (broken.length > base.max) {
        console.log(`  ★入れ子で壊れる式が増えた: ${broken.length} > baseline ${base.max}。独自層(_jsSet)に足した関数はHFプラグインへ移すこと`);
        exitCode = 1;
      } else if (broken.length < base.max) {
        console.log(`  減った(${broken.length} < baseline ${base.max})。--update-baseline で締め直すこと`);
      } else {
        console.log('  baseline どおり(増えていない)');
      }
    }
  }
  if (process.argv.includes('--update-baseline')) {
    fs.writeFileSync(basePath, JSON.stringify({
      note: '★入れ子で答えが変わる式の上限。増えたらCIが赤になる。減らしたら締め直す。',
      max: broken.length, measuredAt: '(実行時)', funcs: Object.keys(byFunc).sort()
    }, null, 1));
    console.log(`  nesting-baseline.json を ${broken.length} で更新`);
  }

  const outPath = path.join(__dirname, 'nesting-audit.json');
  fs.writeFileSync(outPath, JSON.stringify({
    overrideCount: ov.length, pluginCount: plug.length, measured: rows.length,
    brokenCount: broken.length, byFunc, broken: broken.map(r => ({ id: r.id, func: r.outer, f: r.f, top: r.top, nested: r.nested }))
  }, null, 1) + '\n');
  console.log(`\n  詳細: tests/xlsx-harness/nesting-audit.json`);
  return exitCode;
}

if (process.argv[1] && pathToFileURL(process.argv[1]).href === import.meta.url) {
  process.exit(await main());
}
