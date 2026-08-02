/* statutory-freshness.test.mjs — ★確認日が自己失効する仕掛け★
 *
 * なぜ必要か（2026-08-03 の指摘）:
 *   出典URLと確認日を lib に持たせても、★率を変えた人が日付を更新し忘れれば静かに嘘になる。★
 *   「文だけ取り残される」の日付版。手で書いた日付は、手で守っても必ず腐る。
 *
 * やること:
 *   ① 法定の値そのものから【指紋】を作り直す（値の抜き出しは buildStatutoryRows と同じ1本を使う）
 *   ② lib/statutory-meta.js に記録された指紋と違えば ★赤★
 *      →「率を変えたのに確認日を更新していません」と、前の指紋・今の指紋を出す
 *   ③ 確認日が入っている物は、出典URLも必ず入っていること（日付だけの主張を許さない）
 *   ④ 確認日が無い物は、何が未確認かの note があること（黙って空にしない）
 *
 * ★率を1つ変えたら赤になることは、--self-test で毎回確かめる（人の記憶に頼らない）。
 *
 * 使い方: node tests/statutory-freshness.test.mjs
 *         node tests/statutory-freshness.test.mjs --self-test
 */
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));

const SR = require_(path.join(ROOT, 'lib/statutory-rows.js'));
const SM = require_(path.join(ROOT, 'lib/statutory-meta.js'));
const libs = {
  SHH: require_(path.join(ROOT, 'lib/shakaihoken-hyo.js')),
  SAI: require_(path.join(ROOT, 'lib/saitei-chingin.js')),
  KOYO: require_(path.join(ROOT, 'lib/koyo-hoken.js')),
  D: require_(path.join(ROOT, 'lib/shotokuzei-densan.js')),
  H: require_(path.join(ROOT, 'lib/shotokuzei-hei.js')),
  NI: require_(path.join(ROOT, 'lib/shotokuzei-nichi.js')),
  SZ: require_(path.join(ROOT, 'lib/shoyo-zei.js')),
  N: require_(path.join(ROOT, 'lib/nenmatsu.js')),
  WM: require_(path.join(ROOT, 'lib/warimashi.js')),
  SHZ: require_(path.join(ROOT, 'lib/shouhizei-ritsu.js')),
};

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };

/* ★純関数：行とメタから「指紋が合っているか」を返す。self-testで作り物を通せる。 */
export function checkFingerprints(rows, meta) {
  const bad = [];
  for (const r of rows) {
    const key = r.kind + ':' + r.year;
    const m = meta[key];
    if (!m) { bad.push({ key, kind: 'メタが無い', now: SM.fingerprintOf(r.data) }); continue; }
    const now = SM.fingerprintOf(r.data);
    if (m.fingerprint !== now) bad.push({ key, kind: '指紋が違う', was: m.fingerprint, now: now, verified_at: m.verified_at });
  }
  return bad;
}

const rows = SR.buildStatutoryRows(libs);

/* ══ self-test ═══════════════════════════════════════════════════════ */
if (process.argv.includes('--self-test')) {
  console.log('\n[statutory-freshness --self-test] わざと率を変えて赤になるか');

  T('今の実物は指紋が合っている（前提）', () => {
    const bad = checkFingerprints(rows, SM.META);
    if (bad.length) throw new Error('前提が崩れています: ' + JSON.stringify(bad));
  });

  T('★率を1つ変えたら赤になる（雇用保険 一般 0.005 → 0.006）', () => {
    const keep = libs.KOYO.RATES[2026].ippan;
    try {
      libs.KOYO.RATES[2026].ippan = 0.006;
      const bad = checkFingerprints(SR.buildStatutoryRows(libs), SM.META);
      const hit = bad.filter(b => b.key === 'koyo:2026');
      if (!hit.length) throw new Error('率を変えたのに赤になりません＝日付が自己失効しない');
    } finally { libs.KOYO.RATES[2026].ippan = keep; }
  });

  T('★健保を1県だけ変えても赤になる（東京 9.85% → 9.86%）', () => {
    const keep = libs.SHH.KENKO_2026.tokyo;
    try {
      libs.SHH.KENKO_2026.tokyo = 0.0986;
      const bad = checkFingerprints(SR.buildStatutoryRows(libs), SM.META);
      if (!bad.filter(b => b.key === 'shakaihoken:2026').length) throw new Error('1県の変更を拾えていません');
    } finally { libs.SHH.KENKO_2026.tokyo = keep; }
  });

  T('★最低賃金を1県だけ変えても赤になる', () => {
    const keep = libs.SAI.todofuken.tokyo.chingin;
    try {
      libs.SAI.todofuken.tokyo.chingin = 1227;
      const bad = checkFingerprints(SR.buildStatutoryRows(libs), SM.META);
      if (!bad.filter(b => b.key === 'saitei_chingin:2025').length) throw new Error('1県の変更を拾えていません');
    } finally { libs.SAI.todofuken.tokyo.chingin = keep; }
  });

  T('元に戻したら緑に戻る（テストが状態を壊していない）', () => {
    const bad = checkFingerprints(SR.buildStatutoryRows(libs), SM.META);
    if (bad.length) throw new Error('戻っていません: ' + JSON.stringify(bad));
  });

  console.log('\n' + pass + ' passed, ' + fail + ' failed');
  process.exit(fail ? 1 : 0);
}

/* ══ 本番 ═══════════════════════════════════════════════════════════ */
console.log('\n[statutory-freshness] 法定データの出典・確認日・指紋');

const bad = checkFingerprints(rows, SM.META);

T('★値を変えたのに確認日を更新していない、が無い（指紋一致）', () => {
  if (bad.length) {
    throw new Error('法定の値と記録された指紋が合いません:\n'
      + bad.map(b => '   - ' + b.key + '  ' + b.kind + '  前=' + (b.was || '(無し)') + '  今=' + b.now
        + (b.verified_at ? '  記録された確認日=' + b.verified_at : '')).join('\n')
      + '\n   → ★率を変えたのに確認日を更新していません。'
      + '\n     一次情報を実際に開いて突き合わせ、lib/statutory-meta.js の verified_at と fingerprint を打ち直してください。'
      + '\n     ★開いていないなら verified_at:null のまま note に何が未確認かを書くこと（見ていない物に日付を書かない）。');
  }
});

T('確認日がある物は、出典URLも必ずある（日付だけの主張を許さない）', () => {
  for (const k of SM.keys()) {
    const m = SM.META[k];
    if (m.verified_at && !m.source_url) throw new Error(k + ': 確認日はあるのに出典URLが無い');
    if (m.verified_at && !/^\d{4}-\d{2}-\d{2}$/.test(m.verified_at)) throw new Error(k + ': 確認日の形が YYYY-MM-DD でない');
  }
});

T('確認日が無い物は、何が未確認かが書いてある（黙って空にしない）', () => {
  for (const k of SM.keys()) {
    const m = SM.META[k];
    if (!m.verified_at && !(m.note && m.note.length > 5)) throw new Error(k + ': 未確認なのに理由が書いていない');
  }
});

T('メタと行が1対1（増やした行にメタを付け忘れていない・消した行のメタが残っていない）', () => {
  const rowKeys = rows.map(r => r.kind + ':' + r.year).sort();
  const metaKeys = SM.keys();
  const missing = rowKeys.filter(k => metaKeys.indexOf(k) < 0);
  const extra = metaKeys.filter(k => rowKeys.indexOf(k) < 0);
  if (missing.length || extra.length) {
    throw new Error('メタ不整合: 付け忘れ=' + (missing.join(', ') || 'なし') + ' / 余り=' + (extra.join(', ') || 'なし'));
  }
});

T('検査が空振りしていない（行を実際に作れている）', () => {
  if (rows.length < 10) throw new Error('行が少なすぎます: ' + rows.length);
});

const verified = SM.keys().filter(k => SM.META[k].verified_at);
console.log('\n── 実測 ──');
console.log('  法定の行: ' + rows.length + '件 / 確認日あり ' + verified.length + '件 / 未確認 ' + (SM.keys().length - verified.length) + '件');
verified.forEach(k => console.log('   ✔ ' + k + '  ' + SM.META[k].verified_at + '  ' + SM.META[k].source_url));

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
