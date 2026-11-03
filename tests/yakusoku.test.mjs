/* yakusoku.test.mjs — ★画面に 書いてある「約束」を 押して 確かめる★ 2026-09-03
 *
 *  ★きっかけ★＝コメントの 窓に「右上に 赤い印が 出ます」と 書いてあるのに ★出なかった★。
 *    ★字は 緑・試験も 緑・お客さんだけが 損を していた★。
 *
 *  ★この 見張りが 見る 物★
 *    ①★印（data-yakusoku）と 台帳（lib/yakusoku-daicho.js）が 両方 揃っているか★
 *       （片方だけ 消したら 赤＝★どちらかを 忘れられない★）
 *    ②★印なし＝未点検★の 数（★増えたら 赤★／いきなり 全部 赤には しない）
 *    ③★見方が「中の数」の 物は ★本当に 押して★ 確かめる★（jsdom）
 *    ④★見方が「描いた物」「人」の 物は 名指しで 出す★
 *       ＝★「未測定」では なく「機械が 毎回は 見ていない」★（2026-09-03 の 決まり）
 *
 *  ★数の 出どころ★
 *    「約束の 形の 文」は ★字で 拾った 数＝当てに ならない★（説明・注意・見出しも 混ざる）。
 *    ★分母では ない★／★見落としが 増えていないかを 見る 為だけ★。
 *
 *  走らせ方: node tests/yakusoku.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const require_ = createRequire(import.meta.url);
const OVERRIDE = process.env.EXALLY_YAKUSOKU_OVERRIDE ? JSON.parse(process.env.EXALLY_YAKUSOKU_OVERRIDE) : {};
const srcPath = (rel) => OVERRIDE[rel] || path.join(ROOT, rel);

let JSDOM;
try { ({ JSDOM } = await import('jsdom')); }
catch { console.log('★jsdomが入っていません。この検証は飛ばせません（SKIPを緑と呼ばない）'); process.exit(1); }

let pass = 0, fail = 0, 未測定 = 0;
const T = (n, fn) => {
  try { fn(); pass++; console.log('  ok   ' + n); }
  catch (e) { fail++; console.log('  NG   ' + n + '\n       ' + (e && e.message)); }
};
const 言う = (よい, 文) => { if (!よい) throw new Error(文); };

const 台帳 = require_(srcPath('lib/yakusoku-daicho.js'));
const html = fs.readFileSync(srcPath('book.html'), 'utf8');

console.log('');
console.log('[yakusoku] ★画面が した 約束を 押して 確かめる★');

/* ── ①印 ⇔ 台帳 ── */
/* ★印の 付き方は 3通り★（どれも 拾う）
     ①札に 直に … data-yakusoku="名前"
     ②JS で 付ける … setAttribute('data-yakusoku', '名前')
     ③窓の副題（#funcSub は 窓ごとに 中身が 入れ替わる）… 窓の副題('字', '名前')
   ★1つでも 拾い漏らすと「台帳に 在るのに 画面に 無い」と 嘘の 赤に なる★
   （2026-09-03 に ②③を 拾い漏らして 実際に 出た） */
const 印たち = [
  ...[...html.matchAll(/data-yakusoku=["']([^"']+)["']/g)].map((m) => m[1]),
  ...[...html.matchAll(/setAttribute\(\s*['"]data-yakusoku['"]\s*,\s*['"]([^'"]+)['"]/g)].map((m) => m[1]),
  /* ★窓の副題 は 中に ( ) が 入る★（cellAddr(...) など）＝★[^)]* では 途中で 切れる★
     （2026-09-03 実測＝name-insert を 拾い漏らした）⇒★行ごとに 見て 最後の '…' を 取る★ */
  ...html.split(String.fromCharCode(10)).filter((l) => l.indexOf('窓の副題(') >= 0)
    .map((l) => { const m = l.match(/,\s*['"]([A-Za-z0-9_-]+)['"]\s*\)/); return m ? m[1] : null; })
    .filter(Boolean),
];
const 印セット = [...new Set(印たち)];
const 台帳名 = 台帳.名前たち();

T('★画面の 印が 台帳に 全部 在る★', () => {
  const 無い = 印セット.filter((n) => 台帳名.indexOf(n) < 0);
  言う(無い.length === 0, '台帳に 無い 印：' + 無い.join('・'));
});
T('★台帳の 名前が 画面に 全部 在る★', () => {
  const 無い = 台帳名.filter((n) => 印セット.indexOf(n) < 0);
  言う(無い.length === 0, '画面に 無い 名前：' + 無い.join('・'));
});
T('★見張っている 約束が 6つ 以上（減らしていない）★', () => {
  言う(印セット.length >= 6, '今 ' + 印セット.length + '個（6を 下回った）');
});
T('★見方は 3つの どれか★', () => {
  const 変 = 台帳.台帳.filter((v) => 台帳.見方の種類.indexOf(v.見方) < 0);
  言う(変.length === 0, '知らない 見方：' + 変.map((v) => v.名 + '/' + v.見方).join('・'));
});

/* ── ②印なし＝未点検（★増えたら 赤★） ── */
const 未点検の上限 = 17;   /* ★2026-09-03 実測★（この 形の 文 21 の うち 印なし 17）。
   ★数が 19 → 17 に なった 訳★＝★書き手あての 覚え書き（コメント）を 数えるのを やめた★
     （お客さんの 画面には 出ない／★直した 理由を 書いただけで 赤に なった★）
   ★下げる のは 自由・上げるのは 赤★／★下げ忘れても 赤★（ラチェット） */
let 約束の形の数 = 0, 未点検 = 0, コメントの中 = false, HTML覚え書きの中 = false;
{
  const 行 = html.split('\n');
  for (let i = 0; i < 行.length; i++) {
    const l = 行[i];
    /* ★書き手あての 覚え書き（コメント）は 約束では ない★（2026-09-03 に 気づいた）
       ＝★お客さんの 画面には 出ない★。前は これも 数えていて、
       ★直した 理由を コメントに 書いただけで 未点検が 増えて 赤に なった★。
       ⇒★行の 頭が * や // や /* の 物は 数えない★
       ⇒★この 直しで 数が 19 → 12 に 変わった（数の 出どころが 変わった＝紙に 書く）★ */
    const 頭 = l.trim();
    /* ★かたまりの コメント（/* … *​/）の 中は 全部 とばす★（2026-09-04 に 直した）
       前は ★行の 頭が * // /* の 物★だけ とばしていた
       ⇒★「⇒★…★」で 始まる 続きの 行が 素通り★＝★書いただけで 未点検が 増えて 赤★に なった
       （★2度目★＝1度目は「行の 頭が *」で 直した／今度は ★中に 居るか どうか★で 見る） */
    if (コメントの中) { if (l.indexOf('*/') >= 0) コメントの中 = false; continue; }
    if (頭.startsWith('//')) continue;
    if (頭.startsWith('/*')) { if (l.indexOf('*/') < 0) コメントの中 = true; continue; }
    if (頭.startsWith('*')) continue;
    /* ★★HTML の 覚え書き（`<!-- ... -->`）も とばす★★（2026-09-18・★3度目★）
         ★訳★ … `book.html` に `<!-- ... -->` で 訳を 書いたら ★未点検が 17 -> 18★
                ＝★お客さんの 画面には 出ません★＝★約束では ありません★
         ★この 道具は 同じ 形で 既に 2回 直して います★
           1度目 ... 行の 頭が `*` `//` `/*`（2026-09-03・19 -> 12）
           2度目 ... かたまりの 中に 居るか（2026-09-04）
           ★3度目★ ... ★HTML の 覚え書き★（2026-09-18）
         ⇒★★「覚え書きの 形」は 1つでは ありません★★
         ⇒★上限を 上げるのでは なく ★数える 所を 直します★★ */
    if (HTML覚え書きの中) { if (l.indexOf('-->') >= 0) HTML覚え書きの中 = false; continue; }
    if (頭.startsWith('<!--')) { if (l.indexOf('-->') < 0) HTML覚え書きの中 = true; continue; }
    if (!台帳.約束の形.test(l)) { 台帳.約束の形.lastIndex = 0; continue; }
    台帳.約束の形.lastIndex = 0;
    約束の形の数++;
    /* ★印は 同じ 行 か その 1つ 上の 行に 在る★（箱に 付ける ので） */
    if (!/data-yakusoku=/.test(l) && !/data-yakusoku=/.test(行[i - 1] || '')) 未点検++;
    if (!/data-yakusoku=/.test(l) && !/data-yakusoku=/.test(行[i - 1] || '') && process.env.YAKU_DUMP) console.log('  [未点検] ' + (i + 1) + ' ' + l.trim().slice(0, 100));
  }
}
T('★印なし＝未点検が ' + 未点検 + '個（' + 未点検の上限 + '以下）★', () => {
  言う(未点検 <= 未点検の上限,
    '未点検が 増えた（' + 未点検 + ' > ' + 未点検の上限 + '）＝★見張っていない 約束を 増やした★');
});
/* ★上限の ラチェット★（2026-09-03・指示役）
   ★増えたら 赤★だけだと ★減らしても 締まらない★＝また 増やせる。
   ⇒★減ったら 上限も 一緒に 下げる★／★下げ忘れたら 赤★＝★戻れない★ */
T('★減った分だけ 上限も 下げてある（戻れない）★', () => {
  言う(未点検 >= 未点検の上限,
    '★上限を ' + 未点検 + ' に 下げて ください★（今 ' + 未点検の上限 + '／実物 ' + 未点検 + '）'
    + '＝★下げないと また 増やせる★');
});
console.log('       （約束の 形の 文 ' + 約束の形の数 + '／印を 付けた ' + 印セット.length
  + '／未点検 ' + 未点検 + '。★形の 数は 字で 拾った 数＝分母では ない★）');

/* ── ③「中の数」の 物は 本当に 押す ── */
const CANVAS_STUB = [
  '(function(){var noop=function(){return noop;};',
  'var ctx=new Proxy({measureText:function(t){return {width:(t||"").length*7};}},',
  '{get:function(t,k){ if(k in t) return t[k];',
  ' return noop;},set:function(t,k,v){ t[k]=v; return true; }});',
  'HTMLCanvasElement.prototype.getContext=function(){return ctx;};})();',
].join('\n');

const dom = new JSDOM(html.replace(/<script[\s\S]*?<\/script>/g, ''), {
  runScripts: 'dangerously', url: 'http://localhost/', pretendToBeVisual: true,
  beforeParse(w) {
    w.fetch = () => Promise.reject(new Error('no net')); w.scrollTo = () => {}; w.alert = () => {};
    w.matchMedia = w.matchMedia || (() => ({ matches: false, addListener() {}, removeListener() {}, addEventListener() {}, removeEventListener() {} }));
    w.requestAnimationFrame = (cb) => setTimeout(() => cb(1), 0); w.cancelAnimationFrame = () => {};
    w.eval(CANVAS_STUB);
  },
});
const win = dom.window, doc = win.document;
const inject = (c) => { const el = doc.createElement('script'); el.textContent = c; doc.body.appendChild(el); };
const srcs = [...html.matchAll(/<script src="([^"]+)"><\/script>/g)].map((m) => m[1].split('?')[0]).filter((s) => !/^https?:/.test(s));
for (const s of srcs) { const p = srcPath(s); if (fs.existsSync(p)) inject(fs.readFileSync(p, 'utf8')); }
for (const m of html.matchAll(/<script(?![^>]*\ssrc=)[^>]*>([\s\S]*?)<\/script>/g)) inject(m[1]);
try { doc.dispatchEvent(new win.Event('DOMContentLoaded', { bubbles: true })); } catch (e) { /* 続ける */ }
try { win.dispatchEvent(new win.Event('load')); } catch (e) { /* 続ける */ }

const 中の数の試し = {
  'hyou-ireru-undo': () => {
    /* ★書いてある 約束＝「違ったら ↩ で 1押しで 戻ります」★
       ⇒★入れる前を 控えてから 押す★（比べる相手は 測る前に 複製する） */
    win.sheets[win.activeSheet].data = {};
    const 前 = JSON.stringify(win.sheets[win.activeSheet].data);
    const 前のundo = win.undoStack.length;
    win.sel(0, 0, 0, 0);
    const 番 = win._AIの表たち.push([['あ', 'い'], ['う', 'え'], ['お', 'か']]) - 1;
    win.AIの表をシートに入れる(番);
    const 積んだ = win.undoStack.length - 前のundo;
    言う(積んだ === 1, '★undo に ' + 積んだ + '個 積んだ＝' + 積んだ + '回 押さないと 戻らない★');
    win.doUndo();
    言う(JSON.stringify(win.sheets[win.activeSheet].data) === 前,
      '★1回 押しても 入れる前に 戻っていない★');
  },
  'func-insert': () => {
    win.sel(0, 0, 0, 0);
    win.関数を入れる('SUM');
    const c = win.sheets[win.activeSheet].data['0,0'];
    const 入った = c ? String(c.f !== undefined && c.f !== '' ? c.f : (c.v || '')) : '';
    言う(入った === '=SUM(', '「=SUM(」まで 入っていない（今 「' + 入った + '」）');
  },
  'symbol-insert': () => {
    win.sheets[win.activeSheet].data = {};
    win.sel(1, 0, 1, 0);
    win.記号を入れる('©');
    const c = win.sheets[win.activeSheet].data['1,0'];
    const 入った = c ? String(c.f !== undefined && c.f !== '' ? c.f : (c.v || '')) : '';
    言う(入った.indexOf('©') >= 0, 'その 字が 入っていない（今 「' + 入った + '」）');
  },
  'name-insert': () => {
    const e = doc.getElementById('funcSub');
    win.窓の副題('押すと A1 の 式に その 名前が 入ります', 'name-insert');
    言う(e.getAttribute('data-yakusoku') === 'name-insert',
      '副題を 入れ替えたのに 印が 付いていない');
    win.窓の副題('関係の 無い 窓の 字');
    言う(!e.hasAttribute('data-yakusoku'),
      '★別の 窓の 字に なったのに 印が 残っている★＝別の 約束を 見張って しまう');
  },
  /* ══ ★★外へ つなぐか 訊く（2026-09-26）★★ ══
       ★★ここで 何を 見るか★★
         この 約束は ★網の 口で 数える 物★ なので
         ★本物の 数は `tests/soto-wo-kiku-webkit.mjs` が 取ります★（台を 立てて 網を 塞いで 数える）。
         ⇒★ここでは 「門の 口が 在るか」と 「初めは 出さないか」を 見ます★
         ⇒★★＝2つを 足して 初めて 約束が 守られて います★★
       ★口が 消えたら ここが 赤に なります★＝★門を 黙って 外せません★ */
  'soto-kiku': () => {
    const P = require_(srcPath('lib/formula-soto-plug.js'));
    言う(typeof P.外へ出す === 'function', '★門の 口（外へ出す）が 在りません★');
    言う(typeof P.外へ出してよいか === 'function', '★門の 口（外へ出してよいか）が 在りません★');
    言う(typeof P.外へ出すを忘れる === 'function', '★門の 口（外へ出すを忘れる）が 在りません★');
    P.外へ出すを忘れる();
    言う(P.外へ出してよいか() === null, '★初めは null（まだ 訊いて いない）で ないと いけません★');
    P.外へ出す(false);
    言う(P.外へ出してよいか() === false, '★「つながない」が 入って いません★');
    P.外へ出す(true);
    言う(P.外へ出してよいか() === true, '★「つないで よい」が 入って いません★');
    P.外へ出すを忘れる();
    言う(P.外へ出してよいか() === null, '★本を 開き直したら また 訊く（null に 戻る）★');
    /* ★4つの 名は lib が 正本＝ここで 書き足さない★ */
    const F = require_(srcPath('lib/formula-soto.js'));
    言う(F.足した名前().length === 4, '★外へ 出る 関数が 4つで ありません★');
  },
};

for (const v of 台帳.台帳) {
  if (v.見方 !== '中の数') continue;
  const f = 中の数の試し[v.名];
  if (!f) {
    console.log('  --   ★' + v.名 + ' … まだ 押していない★（' + v.今 + '）');
    continue;
  }
  T('★' + v.名 + ' … ' + v.何を見る + '★', f);
}

/* ── ④機械が 毎回は 見ていない 物を 名指しで 出す ── */
{
  const 描いた物 = 台帳.台帳.filter((v) => v.見方 === '描いた物').map((v) => v.名 + '（' + v.今 + '）');
  const 人 = 台帳.台帳.filter((v) => v.見方 === '人').map((v) => v.名 + '（' + v.今 + '）');
  const まだ = 台帳.台帳.filter((v) => v.見方 === '中の数' && !中の数の試し[v.名]).map((v) => v.名);
  console.log('');
  console.log('  ★機械が 毎回は 見ていない 物（隠さず 出す）★');
  /* ★未測定は 緑の 数に 入れない★（2026-09-03・指示役）
     ＝★見張りが 見ていない 物を 緑に 見せない★
     ★週1の 回で 埋まったら その 行を「見ている」に 変える★ */
  for (const v of 台帳.台帳) {
    if (v.見方 === '描いた物') {
      /* ★見張り済み が 書いて あれば 未測定では ない★＝★週1の 回が 毎週 見ている★
         （2026-09-03・指示役「埋まったら その 行を『見ている』に 変える」） */
      if (v.見張り済み) { console.log('    ★見ている★ ' + v.名 + ' … ' + v.何を見る + '（' + v.見張り済み + '／今＝' + v.今 + '）'); }
      else { 未測定++; console.log('    ★未測定★ ' + v.名 + ' … ' + v.何を見る + '（★週1の 実ブラウザの 回で 埋める★／今＝' + v.今 + '）'); }
    }
    else if (v.見方 === '人') { 未測定++; console.log('    ★未測定★ ' + v.名 + ' … ' + v.何を見る + '（★人が 見る しかない★／今＝' + v.今 + '）'); }
  }
  for (const n of まだ) { 未測定++; console.log('    ★未測定★ ' + n + ' … 中の数だが ★試しが まだ 無い★'); }
  console.log('    （週1で 見る ' + 描いた物.length + '／人 ' + 人.length + '／試しが 無い ' + まだ.length + '）');
  T('★「まだ 押していない」が 2個 以下（増やしていない）★', () => {
    言う(まだ.length <= 2, 'まだ 押していない 物が 増えた（' + まだ.join('・') + '）');
  });
}

/* ── わざと 壊して 赤に なるか ── */
if (process.argv.includes('--self-test')) {
  console.log('');
  console.log('★わざと 壊して 赤に なるか★');
  const 壊す = [
    ['★印を 1つ 外す（台帳だけ 残る）★',
      (s) => s.replace(' data-yakusoku="func-insert"', '')],
    ['★約束を 1つ 増やして 印を 付けない（未点検が 増える）★',
      (s) => s.replace('<div id="ctx-menu" ontouchstart',
        '<div>ここを 押すと 行が きれいに なります</div>' + String.fromCharCode(10) + '<div id="ctx-menu" ontouchstart')],
    ['★副題を 入れ替えても 印を 外さない（別の 約束を 見張る）★',
      (s) => s.replace('  else e.removeAttribute(\'data-yakusoku\');', '  else { /* 外さない */ }')],
    ['★関数を 入れても セルに 入らない（約束が 嘘に なる）★',
      (s) => s.replace("  setCell(selR1, selC1, '=' + 名 + '(');", '  /* 入れない */')],
    /* ★上限の ラチェット★＝★印を 付けて 未点検が 減ったのに 上限を 下げない★ と 赤 */
    ['★印を 1つ 増やして 上限を 下げない（戻れる 形に する）★',
      (s) => s.replace("（式でした＝答えが変わります）", "（式でした＝答えが変わります）<!-- data-yakusoku=\"tsuika-test\" -->")],
  ];
  const 元 = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
  const tmp = path.join(ROOT, 'tests', '_yakusoku_broken.html');
  const { execFileSync } = await import('node:child_process');
  for (const [名, f] of 壊す) {
    const 壊れ = f(元);
    if (壊れ === 元) { console.log('  ★素通り★  ' + 名 + '（印が 古い＝直せ）'); fail++; continue; }
    fs.writeFileSync(tmp, 壊れ);
    let 赤 = false;
    try {
      execFileSync(process.execPath, [path.join(ROOT, 'tests', 'yakusoku.test.mjs')], {
        env: { ...process.env, EXALLY_YAKUSOKU_OVERRIDE: JSON.stringify({ 'book.html': tmp }) },
        stdio: 'pipe',
      });
    } catch (e) { 赤 = true; }
    console.log((赤 ? '  赤くなった  ' : '  ★素通り★  ') + 名);
    if (!赤) fail++;
  }
  try { fs.unlinkSync(tmp); } catch (e) { /* 無くてよい */ }
}

console.log('');
console.log('yakusoku: ' + pass + ' 緑 / ' + fail + ' 赤 / ★未測定 ' + 未測定 + '件★'
  + '（★未測定は 緑に 数えていません★）');
process.exit(fail ? 1 : 0);
