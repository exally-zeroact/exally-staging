/* no-dead-ui.test.mjs — ★出来ていない物を客に見せない／画面を止めない★
 *
 * なぜ必要か（2026-08-16・指示役が本番の配信物を数えて発見）:
 *   複数選択の「AIへ」ボタンが本番に在り、押すと
 *     「AI送信はSTEP6で実装予定（N範囲プレビュー）」
 *   と出ていた。トーストが無い時は ★alert★ で ★画面全体が止まる★。
 *   （白紙の印刷ダイアログで固まったのと同じ型）
 *
 * ★止める事 3つ★
 *   ① 画面を止める窓（alert / confirm / prompt）を配信物に1つも置かない
 *   ② 描き終わった画面に「実装予定 / STEP<数字> / 未実装 / TODO」を1つも出さない
 *      ＝★中の言葉（作る側の段取り）を人に見せない★
 *   ③ ★押しても何も無いボタンを置かない★
 *      onclick に書いた関数が本当に在るかを、画面を起動して1つずつ確かめる
 *      （関数だけ消してボタンが残る、の逆の事故を止める）
 *
 * 使い方: node tests/no-dead-ui.test.mjs
 *         node tests/no-dead-ui.test.mjs --self-test   … わざと戻して赤になるか
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));

/* ★見る物＝客のブラウザへ配る物だけ★
   api/ は Vercel の関数（サーバ側）で、客の画面には出ない。
   その中の「まだ未対応（将来実装予定）」は AI への指示文＝Excel関数の対応状況の説明なので対象外。 */
const SKIP_DIR = /node_modules|[\\/]tests[\\/]|[\\/]scripts[\\/]|[\\/]tools[\\/]|[\\/]api[\\/]|[\\/]docs[\\/]|[\\/]supabase[\\/]/;
const SKIP_FILE = /\.min\.js$|hyperformula\.full|^sw\.js$/;

function deliveredFiles() {
  const out = [];
  const walk = (dir) => {
    for (const name of fs.readdirSync(dir)) {
      const p = path.join(dir, name);
      if (name === '.git' || name === 'node_modules') continue;
      const st = fs.statSync(p);
      if (st.isDirectory()) { if (!SKIP_DIR.test(p + path.sep)) walk(p); continue; }
      if (!/\.(html|js)$/.test(name)) continue;
      if (SKIP_DIR.test(p) || SKIP_FILE.test(name)) continue;
      out.push(path.relative(ROOT, p).replace(/\\/g, '/'));
    }
  };
  walk(ROOT);
  return out.sort();
}

/* 画面を止める窓。`window.alert` のような書き方も拾う（前が . でも数える） */
const BLOCKING = /(?:^|[^\w$])(?:window\s*\.\s*)?(alert|confirm|prompt)\s*\(/g;
/* 中の言葉（作る側の段取り）。★人に見せてはいけない★ */
const INTERNAL_WORDS = /実装予定|未実装|STEP\s*\d|TODO|FIXME|工事中|coming\s*soon/i;

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };

/* ── 画面を起動して、描き終わった状態を見る ── */
let JSDOM;
try { ({ JSDOM } = require('jsdom')); }
catch {
  console.log('\n[no-dead-ui] ★jsdom が入っていません（npm install）。★緑ではありません★');
  process.exit(1);
}

async function bootBook(htmlOverride) {
  const html = htmlOverride || fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
  const dom = new JSDOM(html.replace(/<script[\s\S]*?<\/script>/g, ''), {
    runScripts: 'dangerously', pretendToBeVisual: true, url: 'http://localhost/'
  });
  const win = dom.window, doc = win.document;
  win.fetch = () => Promise.reject(new Error('no net'));
  win.scrollTo = () => {};
  /* ★alert を握って「呼ばれたら記録」する★＝本当に画面が止まる物が残っていないかを実際に見る */
  const blocked = [];
  win.alert = (m) => { blocked.push('alert: ' + m); };
  win.confirm = (m) => { blocked.push('confirm: ' + m); return false; };
  win.prompt = (m) => { blocked.push('prompt: ' + m); return null; };
  const ctx = new Proxy({}, {
    get: (t, k) => (k === 'measureText' ? (() => ({ width: 10 }))
      : k === 'canvas' ? { width: 800, height: 600 }
        : k === 'getImageData' ? (() => ({ data: [] })) : (() => {}))
  });
  win.HTMLCanvasElement.prototype.getContext = () => ctx;
  const inject = (code) => { const el = doc.createElement('script'); el.textContent = code; doc.body.appendChild(el); };
  for (const m of html.matchAll(/<script src="([^"]+)"><\/script>/g)) {
    const src = m[1].split('?')[0];
    if (/^https?:/.test(src)) continue;
    inject(fs.readFileSync(path.join(ROOT, src), 'utf8'));
  }
  for (const m of html.matchAll(/<script(?![^>]*\ssrc=)[^>]*>([\s\S]*?)<\/script>/g)) inject(m[1]);
  if (doc.readyState !== 'complete') await new Promise((r) => win.addEventListener('load', r, { once: true }));
  else win.dispatchEvent(new win.Event('load'));
  return { win, doc, blocked };
}

/* 画面に出ている文字（属性の title / placeholder / aria-label も人が読む）
   ★<script> の中身は画面の文字ではない★ ので外す。外さないと
   コードの中のコメントまで「画面に出ている」と数えてしまう。 */
function visibleText(doc) {
  if (!doc.body) return '';
  const clone = doc.body.cloneNode(true);
  for (const el of clone.querySelectorAll('script,style,template,noscript')) el.remove();
  const parts = [clone.textContent || ''];
  for (const el of clone.querySelectorAll('[title],[placeholder],[aria-label],[alt]')) {
    for (const a of ['title', 'placeholder', 'aria-label', 'alt']) {
      const v = el.getAttribute(a);
      if (v) parts.push(v);
    }
  }
  return parts.join('\n');
}

/* onclick="fn(...)" の fn を取り出す */
function onclickTargets(doc) {
  const out = [];
  for (const el of doc.querySelectorAll('[onclick]')) {
    const code = el.getAttribute('onclick') || '';
    for (const m of code.matchAll(/(?:^|[^\w$.])([A-Za-z_$][\w$]*)\s*\(/g)) out.push({ fn: m[1], el: el.id || el.tagName, code: code.slice(0, 60) });
  }
  return out;
}
const JS_KEYWORDS = new Set(['if', 'for', 'while', 'switch', 'return', 'typeof', 'function', 'catch', 'new', 'delete', 'void']);

/* ═══ 自己テスト：わざと戻して赤になるか ═══ */
if (process.argv.includes('--self-test')) {
  console.log('\n[no-dead-ui --self-test] ★わざと戻して赤になるか');
  const html = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');

  T('① alert を1つ足すと 数え方が拾う', () => {
    const broken = html.replace('<body', "<script>function _x(){ alert('だめ'); }</script><body");
    const n = (broken.match(BLOCKING) || []).length;
    if (n === 0) throw new Error('足した alert を拾えていない＝この数え方は空振り');
  });
  T('② 「実装予定」を画面に出すと 文字の検査が拾う', async () => {
    if (!INTERNAL_WORDS.test('AI送信はSTEP6で実装予定')) throw new Error('中の言葉を拾えていない');
    if (INTERNAL_WORDS.test('開きました（見るだけ）')) throw new Error('普通の文まで拾っている＝誤検知');
  });
  T('③ 押しても何も無いボタンを足すと 拾う', async () => {
    const broken = html.replace('<div id="multi-select-bar">', '<div id="multi-select-bar"><button onclick="thisFunctionDoesNotExist()">こわす</button>');
    const { win, doc } = await bootBook(broken);
    const dead = onclickTargets(doc).filter((t) => !JS_KEYWORDS.has(t.fn) && typeof win[t.fn] !== 'function');
    win.close();
    if (!dead.length) throw new Error('無い関数を呼ぶボタンを拾えていない＝この検査は空振り');
  });
  T('④ 本物の book.html には ちゃんと onclick が在る（数える物が0だと空振り）', async () => {
    const { win, doc } = await bootBook();
    const n = onclickTargets(doc).length;
    win.close();
    if (n < 10) throw new Error('onclick が ' + n + '個しか見つからない＝読めていない');
  });
  console.log('\n' + pass + ' passed, ' + fail + ' failed');
  process.exit(fail ? 1 : 0);
}

console.log('\n[no-dead-ui] 出来ていない物を客に見せていないか／画面を止めていないか');

const files = deliveredFiles();
console.log('  見た配信物: ' + files.length + '本');

/* ★消す前の「本当に消しますか」だけは、置き換える物が出来るまで残す★
   ここに書いていない confirm が増えたら赤。alert / prompt は例外なしで0件。 */
const CONFIRM_ALLOW = {
  'js/hub.js': {
    n: 1, what: '取引先を削除する前の1枚',
    why: '消す操作の歯止め。今 外すと ★確認なしで消える★ ので、置き換える物が出来るまで残す',
    replaceWith: '★Timeally が作っている「画面の中で1回 聞く形」を借り物として持ってくる★（作り直さない）',
    due: '2026-09-30',
  },
  'js/ledger.js': {
    n: 1, what: '台帳の記録を削除する前の1枚',
    why: '同上',
    replaceWith: '同上（借りてきた同じ部品を hub と台帳の両方から使う）',
    due: '2026-09-30',
  },
  'seikyu/js/seikyu-app.js': {
    n: 4, what: '入金の記録を消す／請求書を取り消す／下書きを削除する／角印を消す の4枚',
    why: 'どれも消す操作の歯止め。★請求書セッションの持ち物なので Exally は触らない★',
    replaceWith: '同じ借り物の部品（請求書セッションが差し替える）',
    due: '2026-09-30',
  },
};

T('★① alert / prompt が配信物に0件（例外なし）', () => {
  const hits = [];
  for (const f of files) {
    const src = fs.readFileSync(path.join(ROOT, f), 'utf8');
    for (const m of src.matchAll(BLOCKING)) {
      if (m[1] === 'confirm') continue;
      const line = src.slice(0, m.index).split('\n').length;
      hits.push(f + ':' + line + ' ' + m[0].trim());
    }
  }
  if (hits.length) {
    throw new Error('★' + hits.length + '件★ 画面が止まります。知らせはトースト(notify/showToast)で出すこと\n     '
      + hits.slice(0, 10).join('\n     '));
  }
});

T('★①-b confirm は「消す前の1枚」だけ（増えたら赤・台帳に理由と期限）', () => {
  const found = {};
  for (const f of files) {
    const src = fs.readFileSync(path.join(ROOT, f), 'utf8');
    for (const m of src.matchAll(BLOCKING)) {
      if (m[1] !== 'confirm') continue;
      found[f] = (found[f] || 0) + 1;
    }
  }
  const unexpected = Object.keys(found).filter((f) => !CONFIRM_ALLOW[f]);
  if (unexpected.length) throw new Error('台帳に無い confirm: ' + unexpected.join(', '));
  const shown = [];
  for (const f of Object.keys(CONFIRM_ALLOW)) {
    const e = CONFIRM_ALLOW[f];
    /* ★本番とテスト線で持っている画面が違う★。無いファイルは飛ばす（この台帳は両方で同じ物） */
    if (!fs.existsSync(path.join(ROOT, f))) continue;
    if (!found[f]) throw new Error(f + ' の confirm が0個＝置き換えたなら台帳から外すこと（残したままにしない）');
    if (found[f] !== e.n) throw new Error(f + ' の confirm が ' + found[f] + '個（台帳は ' + e.n + '個）＝増やさない');
    if (new Date(e.due) < new Date('2026-08-16')) throw new Error(f + ': 期限切れ ' + e.due);
    shown.push('     残している confirm: ' + f + ' ×' + e.n + ' … ' + e.what + '（期限 ' + e.due + '）');
  }
  if (!shown.length) throw new Error('台帳の対象ファイルが1つも無い＝この検査は空振り');
  shown.forEach((s) => console.log(s));
});

const { win, doc, blocked } = await bootBook();

T('★② 描き終わった画面に「実装予定 / STEP<数字> / 未実装 / TODO」が0件', () => {
  const text = visibleText(doc);
  const bad = text.split('\n').filter((s) => INTERNAL_WORDS.test(s)).map((s) => s.trim()).filter(Boolean);
  if (bad.length) throw new Error('★中の言葉が人に見えています★\n     ' + bad.slice(0, 8).join('\n     '));
});

/* ★押す物の一覧を先に書いてから押す★（数だけ報告しない）
   ②は「描いた直後の画面」しか見られない。今回の「実装予定」は
   ★ボタンを押した時のトーストにだけ出る★ので、実際に押して出た文字を見る。 */
const PRESS_LIST = ['#multi-select-bar button'];
T('★②-b 帯のボタンを実際に押して、出た知らせに中の言葉が無い', () => {
  if (typeof win.toggleMultiSelect === 'function') win.toggleMultiSelect();
  /* ★「範囲を選んでいない」だと どのボタンも同じ断り文で終わってしまい、
     この検査が素通りする★。実際に範囲を1つ持たせてから押す
     （複数選択モードで2つ目の範囲をタップした時に、アプリ自身が入れる形）。 */
  if (Array.isArray(win.selRanges) && win.sheets && win.sheets[win.activeSheet]) {
    win.selRanges.push({ r1: 0, c1: 0, r2: 1, c2: 1, sheet: win.sheets[win.activeSheet].name });
    if (typeof win.updateMultiSelectBar === 'function') win.updateMultiSelectBar();
  }
  const toastEl = doc.getElementById('toast');
  const shown = [];
  let pressed = 0;
  for (const sel of PRESS_LIST) {
    for (const b of doc.querySelectorAll(sel)) {
      const label = (b.textContent || b.id || '').trim();
      b.dispatchEvent(new win.MouseEvent('click', { bubbles: true }));
      pressed++;
      if (toastEl) shown.push(label + ' → ' + (toastEl.textContent || '').trim());
    }
  }
  if (!pressed) throw new Error('押す物が0個＝この検査は空振り（' + PRESS_LIST.join(', ') + '）');
  console.log('     押した物 ' + pressed + '個: ' + shown.map((s) => JSON.stringify(s)).join(' / '));
  const bad = shown.filter((s) => INTERNAL_WORDS.test(s));
  if (bad.length) throw new Error('★押したら中の言葉が出ました★\n     ' + bad.join('\n     '));
  if (blocked.length) throw new Error('★押したら画面を止める窓が開きました★ ' + blocked.join(' / '));
});

T('★②-c 配信するファイルの中身にも 中の言葉が1文字も無い（数えた人が誤読しない）', () => {
  /* ★注記(コメント)に書き写すのも駄目★。配信物を grep で数えた人には
     「まだ残っている」と読める。実際 2026-08-16 に テスト線の配信で2件と数えられた。
     ★見るのは「出来ていない」と読める言葉だけ★。
     `// STEP2: …` のような ★段落の目印★ は普通の書き方なので数えない
     （それまで止めると、関係のない注記の書き換えを強いる＝直しが太る）。 */
  const hits = [];
  for (const f of files) {
    const src = fs.readFileSync(path.join(ROOT, f), 'utf8');
    for (const m of src.matchAll(/実装予定|未実装|工事中|coming\s*soon/gi)) {
      hits.push(f + ':' + src.slice(0, m.index).split('\n').length + ' ' + m[0]);
    }
  }
  if (hits.length) throw new Error('★' + hits.length + '件★（注記の中でも書き写さない）\n     ' + hits.slice(0, 8).join('\n     '));
});

T('★③ ボタンの onclick が呼ぶ関数は全部 実在する（押しても何も無いボタンが無い）', () => {
  const targets = onclickTargets(doc);
  if (targets.length < 10) throw new Error('onclick を ' + targets.length + '個しか読めていない＝空振り');
  const dead = targets.filter((t) => !JS_KEYWORDS.has(t.fn) && typeof win[t.fn] !== 'function');
  if (dead.length) {
    throw new Error('★' + dead.length + '件★ 無い関数を呼ぶボタン\n     '
      + dead.map((d) => d.el + ' → ' + d.fn + '()  [' + d.code + ']').slice(0, 8).join('\n     '));
  }
  console.log('     （onclick ' + targets.length + '個ぜんぶ 実在する関数を呼んでいる）');
});

T('★④ 起動から描き終わりまでに 画面を止める窓が1度も開かない（実際に握って確かめた）', () => {
  if (blocked.length) throw new Error('★' + blocked.length + '件★ ' + blocked.slice(0, 3).join(' / '));
});

T('★⑤ 出来ていない物は「灰色＋理由」で出す（消したボタンが戻っていない）', () => {
  const html = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
  if (/id="multi-select-ai-btn"/.test(html)) {
    throw new Error('★「AIへ」ボタンが戻っています★。中身を作ってから出すこと');
  }
  //  電子ハンコは「灰色＋理由」の形で残してある＝これは許される形
  if (!/nav-disabled[^>]*title="[^"]*準備中/.test(html)) {
    throw new Error('「灰色＋理由」の見本（電子ハンコ）が消えている＝この形を保つこと');
  }
});

win.close();
console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
