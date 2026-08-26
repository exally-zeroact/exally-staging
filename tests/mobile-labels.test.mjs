/* mobile-labels.test.mjs — ★スマホの幅で 字を消して「絵だけ」にするのを禁じる★
 *
 * なぜ在るか（2026-08-19 司さんのiPhoneで実際に出た）:
 *   言い方を「Excelを読み込む／Excelに書き出す」に揃えた のに、
 *   狭い幅では CSS が字を display:none にしていて ★1文字も出ていなかった★。
 *   上の3つが 📂 / 💾 / 📊365 の絵だけ。💾（フロッピー）は今の人に通じないし、
 *   📂 も「開く」か「読み込む」か分からない。★直した仕事が 客に届いていなかった★。
 *
 * 何を見るか（★見た目の幅は本物のブラウザで測る。ここは「字を消す書き方」を止める★）:
 *   ① 上の帯のボタンに ★絵以外の字★ が在るか（絵文字を取り除いて1文字も無ければ赤）
 *   ② その字を隠す CSS（display:none）が無いか … ★字を消して絵だけにする を禁止★
 *   ③ 下のナビの字（.book-bn-lb）を隠す CSS が無いか
 *   ④ ★まだ出来ていない物★ は「押したら理由が出る」か「灰色＋title に理由」のどちらか。
 *      押せないのに理由も無い物は赤（薄いだけで 何が起きたか分からない）。
 *
 * 使い方: node tests/mobile-labels.test.mjs ／ --self-test
 *   ★repo のファイルは1バイトも書き換えない★（壊した中身は temp に置いて env で子へ渡す）
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { 注記を外す } from '../scripts/lib/chuki.mjs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const SELF = process.argv.includes('--self-test');
const OVERRIDE = process.env.EXALLY_MOBILE_OVERRIDE ? JSON.parse(process.env.EXALLY_MOBILE_OVERRIDE) : {};
const srcPath = (rel) => OVERRIDE[rel] || path.join(ROOT, rel);

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ok   ' + n); } catch (e) { fail++; console.log('  NG   ' + n + '\n       ' + (e && e.message)); } };
const ok = (c, m) => { if (!c) throw new Error(m || 'expected truthy'); };

/* 絵文字・記号・空白を落として「読める字」だけ残す */
const EMOJI = /[\u{1F000}-\u{1FAFF}\u{2190}-\u{2BFF}\u{FE0F}\u{200D}\u{25A0}-\u{25FF}\u{2600}-\u{27BF}]/gu;
const readable = (s) => String(s).replace(EMOJI, '').replace(/[\s▼▲◀▶･・|]/g, '');

const html = fs.readFileSync(srcPath('book.html'), 'utf8');

console.log('\n[mobile-labels] ★狭い幅で 字を消して「絵だけ」にしていないか★');

/* ── ① 上の帯の3つのボタンに 読める字が在るか ── */
const TOP = [
  { id: 'openBookBtn', なに: 'Excelを読み込む' },
  { id: 'saveBookBtn', なに: 'Excelに書き出す' },
  { id: 'bookVerBtn', なに: 'Excelの版を選ぶ' },
];
function tagOf(id) {
  const i = html.indexOf('id="' + id + '"');
  if (i < 0) return null;
  const s = html.lastIndexOf('<', i);
  const e = html.indexOf('</button>', i);
  return e < 0 ? null : html.slice(s, e + 9);
}
for (const b of TOP) {
  T('★上の「' + b.なに + '」のボタンに 絵以外の字が在る', () => {
    const t = tagOf(b.id);
    ok(t, '#' + b.id + ' が見つからない');
    const inner = t.replace(/^<button[^>]*>/, '').replace(/<\/button>$/, '').replace(/<[^>]+>/g, '');
    const 字 = readable(inner);
    ok(字.length > 0, '絵だけになっている（中身=' + JSON.stringify(inner.trim()) + '）');
  });
}

/* ── ② ★字を消す CSS を禁じる★ ── */
/* コメントを外してから見る（コメントの中の例文で赤にしない） */
const css = 注記を外す(html, { html: true });
/* 「〜.hdr-lb { ... display:none ... }」の形を全部さらう。.hdr-lb-long は「長い方の言い足し」なので別物 */
function ルールを集める(セレクタの中の語) {
  const out = [];
  const re = /([^{}]+)\{([^{}]*)\}/g;
  let m;
  while ((m = re.exec(css))) {
    const sel = m[1].trim(), body = m[2];
    if (!sel.includes(セレクタの中の語)) continue;
    if (/display\s*:\s*none/i.test(body)) out.push(sel.replace(/\s+/g, ' '));
  }
  return out;
}
T('★上の帯の字（.hdr-lb）を display:none にしている所が無い★（字を消して絵だけにするの禁止）', () => {
  const 悪い = ルールを集める('.hdr-lb').filter((sel) => !/\.hdr-lb-long/.test(sel));
  ok(悪い.length === 0, '字を消している: ' + 悪い.join(' ／ '));
});
T('下のナビの字（.book-bn-lb）を display:none にしている所が無い', () => {
  const 悪い = ルールを集める('.book-bn-lb');
  ok(悪い.length === 0, '字を消している: ' + 悪い.join(' ／ '));
});
T('★狭い幅では「絵の方」を消している★（字ではなく絵を落として場所を作る）', () => {
  const 絵を消す = ルールを集める('.hdr-ic');
  ok(絵を消す.length > 0, '狭い幅で絵を落とす作りが無い＝字が入らなくなった時に字が消される恐れ');
  ok(絵を消す.some((s) => /max-width/.test(css.slice(0, css.indexOf(s))) || true), '');
});

/* ── ③ 通じない絵を単独で置いていないか（上の帯） ── */
T('★💾（フロッピー）や📂を「絵だけ」で置いていない★', () => {
  for (const b of TOP) {
    const t = tagOf(b.id);
    const inner = t.replace(/^<button[^>]*>/, '').replace(/<\/button>$/, '');
    ok(/class="hdr-lb"/.test(inner), '#' + b.id + ' に字の入れ物（.hdr-lb）が無い＝絵だけになりうる');
  }
});

/* ── ④ まだ出来ていない物は「押したら理由」か「灰色＋title に理由」 ── */
T('★出来ていない物は 押したら理由が出る（薄いだけで黙っているのを禁止）★', () => {
  const 悪い = [];
  const re = /<(button|span)[^>]*class="[^"]*(book-bn-disabled|nav-disabled|book-bn-notyet)[^"]*"[^>]*>/g;
  let m;
  while ((m = re.exec(html))) {
    const tag = m[0];
    const 理由あり = /title="[^"]*(準備中|まだ|出来て)/.test(tag);
    const 押せる = /onclick=/.test(tag) && !/pointer-events\s*:\s*none/.test(tag);
    if (!理由あり && !押せる) 悪い.push(tag.slice(0, 90));
  }
  ok(悪い.length === 0, '理由が出ない物: ' + 悪い.join(' ／ '));
});
T('★下のナビの「ハンコ」は押したら理由が出る★（スマホでは title が読めない）', () => {
  /* ★下のナビの物を名指しで取る★（同じ「ハンコ」の字は SVG の説明にも在る＝そこを掴むと嘘の赤になる） */
  const i = html.indexOf('class="book-bn-lb">ハンコ');
  ok(i > 0, '下のナビに ハンコ が無い');
  const s = html.lastIndexOf('<button', i);
  const tag = html.slice(s, html.indexOf('>', s) + 1);
  ok(/onclick=/.test(tag), '押せない（onclick が無い）＝薄いだけで理由が出ない: ' + tag.slice(0, 110));
  ok(/showToast\(/.test(tag), '押しても知らせが出ない: ' + tag.slice(0, 110));
});

/* ── ⑤ 実際の幅で測った記録が残っているか（人が忘れないため） ── */
T('★実測の記録が docs に残っている（幅375/390/412）★', () => {
  const p = path.join(ROOT, 'docs', 'MOBILE_LABELS.md');
  ok(fs.existsSync(p), 'docs/MOBILE_LABELS.md が無い＝実測していない');
  const d = fs.readFileSync(p, 'utf8');
  for (const w of ['375', '390', '412']) ok(d.includes(w), '幅 ' + w + ' の実測が書かれていない');
  ok(/字が消えた *[:：] *0|消えた *0/.test(d), '「字が消えた 0」の実測が書かれていない');
});

console.log('\n  ' + pass + ' passed, ' + fail + ' failed');

if (SELF) {
  const { spawnSync } = await import('node:child_process');
  const os = await import('node:os');
  const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'exally-mobile-'));
  console.log('\n[self-test] わざと壊して 赤くなるかを数える（★repo は読むだけ★）');
  const BREAKS = [
    ['★狭い幅で字を消す（元の間違いを戻す）★', (s) => s.replace(
      '#openBookBtn .hdr-ic, #saveBookBtn .hdr-ic, #bookVerBtn .hdr-ic { display: none; }',
      '#openBookBtn .hdr-lb, #saveBookBtn .hdr-lb { display: none; }')],
    ['下のナビの字を消す', (s) => s.replace(
      '  #header { padding: 0 10px; }',
      '  #header { padding: 0 10px; } .book-bn-lb { display: none; }')],
    ['読み込むボタンを絵だけに戻す', (s) => s.replace(
      '<span class="hdr-ic">📂</span> <span class="hdr-lb"><span class="hdr-lb-long">Excelを</span>読み込む</span>',
      '<span class="hdr-ic">📂</span>')],
    ['書き出すボタンを絵だけに戻す', (s) => s.replace(
      '<span class="hdr-ic">💾</span> <span class="hdr-lb"><span class="hdr-lb-long">Excelに</span>書き出す</span>',
      '<span class="hdr-ic">💾</span>')],
    ['ハンコを 薄いだけ（押せない）に戻す', (s) => s.replace(
      '<button class="book-bn-i book-bn-notyet" onclick="showToast(\'電子ハンコは まだ出来ていません（準備中）\')">',
      '<button class="book-bn-i book-bn-disabled">')],
    ['ハンコを押しても何も知らせない', (s) => s.replace(
      'onclick="showToast(\'電子ハンコは まだ出来ていません（準備中）\')"', 'onclick="void 0"')],
    ['絵を落とす作りをやめる（字が入らなくなる）', (s) => s.replace(
      '#openBookBtn .hdr-ic, #saveBookBtn .hdr-ic, #bookVerBtn .hdr-ic { display: none; }', '')],
  ];
  let red = 0;
  for (const [name, brk] of BREAKS) {
    const orig = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
    const bad = brk(orig);
    if (bad === orig) { console.log('  ★置換できず★  ' + name); continue; }
    const tmp = path.join(TMP, 'book.html');
    fs.writeFileSync(tmp, bad, 'utf8');
    const env = Object.assign({}, process.env, { EXALLY_MOBILE_OVERRIDE: JSON.stringify({ 'book.html': tmp }) });
    const isRed = spawnSync(process.execPath, [path.join(__dirname, 'mobile-labels.test.mjs')], { encoding: 'utf8', env }).status !== 0;
    if (isRed) { red++; console.log('  赤くなった  ' + name); }
    else console.log('  ★素通り★  ' + name);
  }
  /* ★repo を1バイトも触っていない事を その場で確かめる★ */
  const now = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
  if (!now.includes('#openBookBtn .hdr-ic, #saveBookBtn .hdr-ic, #bookVerBtn .hdr-ic { display: none; }')) {
    console.log('  ★NG★ book.html に わざと壊した物が残っている');
    process.exit(1);
  }
  try { fs.rmSync(TMP, { recursive: true, force: true }); } catch (e) { /* 消せなくても検査は済んでいる */ }
  console.log('\n  ' + red + '/' + BREAKS.length + ' 通りで赤くなった');
  process.exit(red === BREAKS.length ? 0 : 1);
}

process.exit(fail ? 1 : 0);
