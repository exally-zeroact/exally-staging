/* kikan.test.mjs — ★期間（何月分・◯日からの分）★
 * =============================================================================
 * ★借り物の正本（2026-08-27）★
 *   締め期間の決め方は ★Timeally が正本★（timeally/lib/tc-calc.js の period）。
 *   ★測り方も 借りる★＝timeally/tests/close-period.test.mjs と同じ
 *     ★締め日1〜31 × 月4種 ＝ 124通り★ を そのまま測る。
 *     （★見本を選んで測らない★＝「10/20/25/末日の4つ」を選んだ理由が無い、が向こうの経緯）
 *
 * ★ずれの見張り★
 *   同じ機械に timeally が在る時は ★向こうの period と 1文字ずつ 比べる★。
 *   無い機械では ★未測定★と出す（0件・異常なしにしない）。
 *
 * 使い方: node tests/kikan.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const SELF = process.argv.includes('--self-test');
const OVERRIDE = process.env.EXALLY_KIKAN_OVERRIDE ? JSON.parse(process.env.EXALLY_KIKAN_OVERRIDE) : {};
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const K = require_(OVERRIDE['lib/kikan.js'] || path.join(ROOT, 'lib/kikan.js'));

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ok   ' + n); } catch (e) { fail++; console.log('  NG   ' + n + '\n       ' + (e && e.message)); } };
const ok = (c, m) => { if (!c) throw new Error(m || 'expected truthy'); };
const eq = (a, b, m) => { if (a !== b) throw new Error((m ? m + ': ' : '') + '期待=' + JSON.stringify(b) + ' 実際=' + JSON.stringify(a)); };

console.log('');
console.log('[kikan] ★何月分・◯日からの分（締め期間は Timeally が正本）★');

/* ══ ★124通り（借りた測り方 そのまま）★ ══ */
const daysIn = (y, m) => new Date(Date.UTC(y, m, 0)).getUTCDate();
const iso = (y, m, d) => y + '-' + String(m).padStart(2, '0') + '-' + String(d).padStart(2, '0');
const spanDays = (from, to) =>
  Math.round((Date.parse(to + 'T00:00:00Z') - Date.parse(from + 'T00:00:00Z')) / 86400000) + 1;

const MONTHS = [
  ['2026-02', '2月（平年28日）'],
  ['2024-02', '2月（うるう年29日）'],
  ['2026-04', '30日の月'],
  ['2026-08', '31日の月'],
];

let n = 0;
const bad = [];
for (const [ym, name] of MONTHS) {
  const y = +ym.slice(0, 4), m = +ym.slice(5, 7);
  const last = daysIn(y, m);
  const pm = m === 1 ? 12 : m - 1, py = m === 1 ? y - 1 : y;
  const prevLast = daysIn(py, pm);
  for (let cd = 1; cd <= 31; cd++) {
    n++;
    const p = K.period(ym, cd);
    const wantTo = iso(y, m, Math.min(cd, last));
    const wantFrom = cd >= 31 ? iso(y, m, 1) : iso(py, pm, Math.min(cd + 1, prevLast));
    const why = name + '・締め日' + cd + ' … ';
    if (p.to !== wantTo) bad.push(why + '終わりが ' + p.to + '（' + wantTo + ' のはず）');
    if (p.from !== wantFrom) bad.push(why + '始まりが ' + p.from + '（' + wantFrom + ' のはず）');
    if (spanDays(p.from, p.to) < 27) bad.push(why + '日数が ' + spanDays(p.from, p.to) + '日しかない');
    const cross = p.from.slice(0, 7) !== p.to.slice(0, 7);
    if (cd >= 31 && cross) bad.push(why + '末日締めなのに 月をまたいだ');
    if (cd < 31 && !cross) bad.push(why + '月をまたいでいない');
  }
}
T('★124通り 全部 測った（空振りしていない）★', () => {
  eq(n, 124, '測った数');
  console.log('       … 実測 ' + n + '通り（締め日1〜31 × 月4種）');
});
T('★どの締め日でも 終わり＝締め日／始まり＝締め日の翌日（無い日は末日に寄せる）★', () => {
  eq(bad.length, 0, bad.slice(0, 5).join(' / '));
});

/* ══ ★司さんの言葉（1日から／10日から）★ ══ */
T('★1日からの分＝その月まるごと★', () => {
  const k = K.期間を決める('2026-01', 1);
  eq(k.from, '2026-01-01'); eq(k.to, '2026-01-31'); eq(k.締め日, 31);
});
T('★10日からの分＝前の月の10日〜その月の9日★', () => {
  const k = K.期間を決める('2026-01', 10);
  eq(k.from, '2025-12-10'); eq(k.to, '2026-01-09'); eq(k.締め日, 9);
});
T('★21日からの分（実物の給料表の形）★', () => {
  const k = K.期間を決める('2026-03', 21);
  eq(k.from, '2026-02-21'); eq(k.to, '2026-03-20');
});
T('★年をまたぐ時は 言い方に 年を出す（どの年の12月か 分かるように）★', () => {
  ok(K.期間を決める('2026-01', 10).言い方.indexOf('2025/12/10') > 0,
    '★年が出ていない：' + K.期間を決める('2026-01', 10).言い方);
  eq(K.期間を決める('2026-03', 21).言い方, '3月分（2/21〜3/20）', '★またがない時に 年を出している★');
});
T('★客に見せる字に ★ を書かない★', () => {
  for (const d of [1, 10, 21, 31]) {
    const k = K.期間を決める('2026-01', d);
    ok(k.言い方.indexOf('★') < 0, '★客の字に ★ が出ている：' + k.言い方);
  }
});

/* ══ ★境界（実物で踏む所）★ ══ */
T('★2月に 30日・31日から を頼まれても 穴を開けない★', () => {
  const k = K.期間を決める('2026-03', 31);
  eq(k.from, '2026-02-28', '★2月30日が 消えている★');
  ok(k.to <= '2026-03-31');
});
T('★1月分の「10日から」は 前の年へ 戻る★', () => {
  eq(K.期間を決める('2026-01', 10).from.slice(0, 4), '2025');
});
T('★おかしい入力は null（黙って 今月にしない）★', () => {
  eq(K.期間を決める('', 1), null);
  eq(K.期間を決める('2026-13-01', 1), null, '月の形が違う物を 通している');
  eq(K.期間を決める('2026-01', 0), null);
  eq(K.期間を決める('2026-01', 32), null);
});

/* ══ ★日付（通し番号）★ ══ */
T('★通し番号を 日付に直す（実物の 46023 は 2026-01-01）★', () => {
  eq(K.通し番号を日付に(46023), '2026-01-01');
  eq(K.通し番号を日付に(46043), '2026-01-21');
  eq(K.通し番号を日付に(45658), '2025-01-01');
});
T('★Excelと同じ 1900年の うるう年の嘘★', () => {
  /* ★実測して 見つけた（2026-08-27）★＝よく使う式1本（1899-12-30起点）だと
     1〜59 が 1日 ずれる（1 が 1899-12-31 になる）。 */
  eq(K.通し番号を日付に(1), '1900-01-01');
  eq(K.通し番号を日付に(59), '1900-02-28');
  eq(K.通し番号を日付に(60), '1900-02-29', '★Excelの嘘の日を 黙って ずらしている★');
  eq(K.通し番号を日付に(61), '1900-03-01');
  eq(K.通し番号を日付に(46023), '2026-01-01', '★普段の日付が 壊れた★');
});
T('★時刻つきでも その日として扱う★', () => {
  eq(K.通し番号を日付に(46023.75), '2026-01-01');
  /* ★時刻ぶんを 落とさないと 1900-02-29 が 3/1 に化ける★（実際に 壊して 確かめた） */
  eq(K.通し番号を日付に(60.5), '1900-02-29', '★時刻ぶんを 切り捨てていない★');
  eq(K.通し番号を日付に(59.9), '1900-02-28');
});
T('★日付でない物は null（0や字を 1900年にしない）★', () => {
  eq(K.日付に直す(''), null);
  eq(K.日付に直す('合計'), null);
  eq(K.日付に直す(0), null);
  eq(K.日付に直す(null), null);
});
T('★字で書かれた日付も 読む★', () => {
  eq(K.日付に直す('2026-01-05'), '2026-01-05');
  eq(K.日付に直す('2026/1/5'), '2026-01-05');
});
T('★期間の中か は 両端を含む（時差で変わらない字くらべ）★', () => {
  const k = K.期間を決める('2026-01', 1);
  ok(K.期間の中か('2026-01-01', k), '★始まりの日が 外れている★');
  ok(K.期間の中か('2026-01-31', k), '★終わりの日が 外れている★');
  ok(!K.期間の中か('2025-12-31', k));
  ok(!K.期間の中か('2026-02-01', k));
});

/* ══ ★借り物が ずれていないか（正本と 1文字ずつ 比べる）★ ══ */
const TIMEALLY = 'C:/Users/zeroa/timeally/lib/tc-calc.js';
if (!fs.existsSync(TIMEALLY)) {
  console.log('  ★未測定★ この機械に timeally が無いので 正本と比べられません（0件・異常なしにしない）');
} else {
  T('★正本（Timeally）の period と 1文字も違わない★', () => {
    const 取る = (src) => {
      const i = src.indexOf('function period(ym, closeDay) {');
      ok(i > 0, '★period が 見つからない★');
      let d = 0, j = src.indexOf('{', i);
      for (let k = j; k < src.length; k++) {
        if (src[k] === '{') d++;
        else if (src[k] === '}') { d--; if (!d) { j = k; break; } }
      }
      return src.slice(i, j + 1).replace(/\r\n/g, '\n');
    };
    const 正本 = 取る(fs.readFileSync(TIMEALLY, 'utf8'));
    const うち = 取る(fs.readFileSync(OVERRIDE['lib/kikan.js'] || path.join(ROOT, 'lib/kikan.js'), 'utf8'));
    eq(うち, 正本, '★借り物が ずれた（どちらかが 直された）★');
    console.log('       … 正本 timeally/lib/tc-calc.js と 同じ（' + 正本.split('\n').length + '行）');
  });
}

console.log('');
console.log('  ' + pass + ' 緑 / ' + fail + ' 赤');

if (SELF) {
  const { spawnSync } = await import('node:child_process');
  const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'exally-kikan-'));
  console.log('');
  console.log('[self-test] わざと壊して 赤くなるかを数える（★repo は読むだけ★）');
  const BREAKS = [
    ['★末日締めを その月の1日から にしない★',
      (s) => s.replace("return { ym: ym, from: ym + '-01', to: ym + '-' + pad2(daysInMonth(y, m)) };",
                       "return { ym: ym, from: ym + '-02', to: ym + '-' + pad2(daysInMonth(y, m)) };")],
    ['★無い日を 末日に寄せない（2月30日で 穴を開ける）★',
      (s) => s.replace('var fromD = Math.min(cd + 1, daysInMonth(py, pm));', 'var fromD = cd + 1;')],
    ['★終わりの日を 寄せない★',
      (s) => s.replace('var toD = Math.min(cd, daysInMonth(y, m));', 'var toD = cd;')],
    ['★1日からを 締め日1 にする（前の月から になってしまう）★',
      (s) => s.replace('return d === 1 ? 31 : d - 1;', 'return d;')],
    ['★おかしい締め日を 通す★',
      (s) => s.replace('if (!(d >= 1 && d <= 31)) return null;', '')],
    ['★月の形を 見ない★',
      (s) => s.replace("if (!/^\\d{4}-\\d{2}$/.test(String(ym || ''))) return null;", '')],
    ['★通し番号の起点を 1日 ずらす★',
      (s) => s.replace('Date.UTC(1899, 11, 30)', 'Date.UTC(1899, 11, 31)')],
    ['★時刻つきを 切り捨てない★',
      (s) => s.replace('var i = Math.floor(v);', 'var i = v;')],
    ['★0 を 日付にする★',
      (s) => s.replace('if (!isFinite(v) || v < 1) return null;', 'if (!isFinite(v)) return null;')],
    ['★期間の端を 含めない★',
      (s) => s.replace('return ymd >= 期間.from && ymd <= 期間.to;', 'return ymd > 期間.from && ymd < 期間.to;')],
    ['★年をまたいでも 年を出さない★',
      (s) => s.replace("var またぐ = p.from.slice(0, 4) !== p.to.slice(0, 4);", 'var またぐ = false;')],
    ['★客の字に ★ を書く★',
      (s) => s.replace("言い方: 月 + '月分（'", "言い方: '★' + 月 + '月分（'")],
    ['★1900年の 1〜59 を 1日 ずらす（よく使う式1本に 戻す）★',
      (s2) => s2.replace('var 起点 = (i < 60) ? Date.UTC(1899, 11, 31) : Date.UTC(1899, 11, 30);', 'var 起点 = Date.UTC(1899, 11, 30);')],
    ['★Excelの嘘の日(60)を 黙って ずらす★',
      (s2) => s2.replace("if (i === 60) return '1900-02-29';", '')],
  ];
  let red = 0;
  for (const [name, brk] of BREAKS) {
    const orig = fs.readFileSync(path.join(ROOT, 'lib/kikan.js'), 'utf8');
    const bad2 = brk(orig);
    if (bad2 === orig) { console.log('  ★置換できず★  ' + name); continue; }
    const f = path.join(TMP, 'kikan.js');
    fs.writeFileSync(f, bad2, 'utf8');
    const env = Object.assign({}, process.env, { EXALLY_KIKAN_OVERRIDE: JSON.stringify({ 'lib/kikan.js': f }) });
    const r = spawnSync(process.execPath, [path.join(__dirname, 'kikan.test.mjs')], { encoding: 'utf8', env });
    if (r.status !== 0) { red++; console.log('  赤くなった  ' + name); }
    else console.log('  ★素通り★  ' + name);
  }
  try { fs.rmSync(TMP, { recursive: true, force: true }); } catch (e) { /* 消せなくても検査は済んでいる */ }
  console.log('');
  console.log('  ' + red + '/' + BREAKS.length + ' 通りで赤くなった');
  process.exit(red === BREAKS.length ? 0 : 1);
}

process.exit(fail ? 1 : 0);
