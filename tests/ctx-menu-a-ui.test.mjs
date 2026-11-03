/* ctx-menu-a-ui.test.mjs — ★本物の book.html を 載せて 右クリックを 実際に 出す★ 2026-09-03
 *
 *  見るのは 2つ（どちらも ★A で 足した 分★）
 *   ①★ハイパーリンクを開く＝リンクが 在る時だけ 押せる／無ければ 灰★
 *      （実Excel も 同じ 見え方＝リンクが 無い セルでは 灰）
 *   ②★一覧が 長い時だけ 1行に する★
 *      ★個数では 決めない★＝窓の 高さで 境目が 動く（実ブラウザ実測 720で19個／560で14個）
 *      ⇒ 決めるのは ★「切り取り」が 転がさずに 見えるか★
 *      ⇒ jsdom は ★高さを 持たない★＝★測れない★ ⇒ ★安全側＝1行★に なる事を 見る
 *        （★測れない時に「一覧のまま」に なったら 赤★＝50個で 切り取りが 消える 事故に 戻る）
 *
 *  ★本物の 見た目は 実ブラウザで 別に 押しています★（絵＝shot/list05.png・list50.png）。
 *  使い方: node tests/ctx-menu-a-ui.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const SELF = process.argv.includes('--self-test');
const OVERRIDE = process.env.EXALLY_CTXA_OVERRIDE ? JSON.parse(process.env.EXALLY_CTXA_OVERRIDE) : {};
const srcPath = (rel) => OVERRIDE[rel] || path.join(ROOT, rel);

let JSDOM;
try { ({ JSDOM } = await import('jsdom')); }
catch { console.log('★jsdomが入っていません。この検証は飛ばせません（SKIPを緑と呼ばない）'); process.exit(1); }

const CANVAS_STUB = [
  '(function(){var noop=function(){return noop;};',
  'var ctx=new Proxy({measureText:function(t){return {width:(t||"").length*7};}},',
  '{get:function(t,k){ if(k in t) return t[k];',
  ' return noop;},set:function(t,k,v){ t[k]=v; return true; }});',
  'HTMLCanvasElement.prototype.getContext=function(){return ctx;};})();',
].join('\n');

const html = fs.readFileSync(srcPath('book.html'), 'utf8');
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

let pass = 0, fail = 0;
const T = (n, fn) => {
  try { fn(); pass++; console.log('  ok   ' + n); }
  catch (e) { fail++; console.log('  NG   ' + n + '\n       ' + (e && e.message)); }
};
const 言う = (よい, 文) => { if (!よい) throw new Error(文); };

console.log('');
console.log('[ctx-menu-a-ui] ★本物の 画面で 右クリックを 出す（A で 足した 分）★');

const 出す = () => { win.showCtxMenu(10, 10); };
const menu = () => doc.getElementById('ctx-menu');

/* ── ①ハイパーリンクを開く＝灰に なるか ── */
T('★リンクが 無い セルでは 灰（ctx-off）★', () => {
  win.sheets[win.activeSheet].links = {};
  win.sel(0, 0, 0, 0);
  出す();
  const el = doc.getElementById('ctx-link-open');
  言う(!!el, '「ハイパーリンクを開く」の 札が 無い');
  言う(el.classList.contains('ctx-off'), '灰に なっていない（リンクが 無いのに 押せる）');
});

T('★リンクが 在る セルでは 押せる（灰が 外れる）★', () => {
  win.sheets[win.activeSheet].links = { '0,0': { 先: 'https://example.com', 字: 'れい' } };
  win.sel(0, 0, 0, 0);
  出す();
  const el = doc.getElementById('ctx-link-open');
  言う(!el.classList.contains('ctx-off'), 'リンクが 在るのに 灰のまま');
  言う(/example\.com/.test(el.title || ''), '行き先が 説明に 出ていない');
});

T('★灰の 札を 押しても 何も 起きない（メニューも 閉じない）★', () => {
  win.sheets[win.activeSheet].links = {};
  win.sel(0, 0, 0, 0);
  出す();
  const el = doc.getElementById('ctx-link-open');
  const 前 = menu().classList.contains('show');
  el.onclick.call(el, { stopPropagation() {} });
  言う(前 === true, 'そもそも 出ていない');
  言う(menu().classList.contains('show') === true, '灰を 押したのに 閉じた');
});

/* ── ②一覧が 長い時 ── */
const 決まりを置く = (n) => {
  const items = [];
  for (let i = 1; i <= n; i++) items.push('えらぶ' + i);
  win.sheets[win.activeSheet].validations = { '0,0': { kind: 'list', items } };
  win.sel(0, 0, 0, 0);
};

T('★測れない時は 安全側＝1行（jsdom は 高さを 持たない）★', () => {
  決まりを置く(50);
  出す();
  言う(win.__validListMode !== '一覧',
    '測れないのに「一覧」のまま（50個 並べたら 切り取りが 消える）＝今 ' + win.__validListMode);
  const one = doc.getElementById('ctx-valid-one');
  言う(!!one, '1行の 札（ctx-valid-one）が 無い');
  言う(/ドロップダウン/.test(one.textContent), '字が 実Excel と 違う');
});

T('★1行を 押すと その場で 一覧に なる（窓は 出さない）★', () => {
  決まりを置く(50);
  出す();
  const one = doc.getElementById('ctx-valid-one');
  one.onclick.call(one, { stopPropagation() {} });
  const 札 = doc.querySelectorAll('[data-valid-pick]');
  言う(札.length === 50, '一覧に ならない（今 ' + 札.length + '個）');
  言う(menu().classList.contains('show'), 'メニューが 閉じた（窓を 出してはいけない）');
});

T('★決まりが 無い セルでは 何も 出さない★', () => {
  win.sheets[win.activeSheet].validations = {};
  win.sel(0, 0, 0, 0);
  出す();
  言う(win.__validListMode === 'なし', '出ないはずの 所で 出ている（今 ' + win.__validListMode + '）');
  言う(doc.querySelectorAll('[data-valid-pick]').length === 0, '選ぶ札が 残っている');
});

T('★切り替えは 出すたび 1回だけ（行ったり来たり しない）★', () => {
  決まりを置く(50);
  出す();
  言う(win.__validListSwitch === null, '切り替えの 手が 残っている（何度も 走る）');
});

/* ── ★短い時は 一覧のまま／長い時は 1行★（★作り物の 高さを 与えて 見る★） ──
   ★本物の 画面では ありません★＝jsdom は 高さを 持たないので ★こちらで 高さを 作る★。
     ・.ctx-item は 1つ 34px（実ブラウザ実測 約34.5px）
     ・#ctx-menu の 箱の 高さは この 試験が 決める
   ⇒★「切り取りが 見えるか」で 切り替わっているか★を 数で 見張れる。
   ★実ブラウザの 数（720で19個／560で14個）は 手で 測って 紙に 書いてある★
     （docs/EXCEL_CELL_CTXMENU_2026-09-03.md）。
   ★CI に 実ブラウザは 在りません★＝そこは ★未測定★と 紙に 書いた。 */
const 高さを作る = (箱の高さ) => {
  const H = 34;
  const proto = win.HTMLElement.prototype;
  Object.defineProperty(proto, 'offsetHeight', {
    configurable: true,
    get() { return this.classList && this.classList.contains('ctx-item') ? H : 0; },
  });
  Object.defineProperty(proto, 'offsetTop', {
    configurable: true,
    get() {
      const m = doc.getElementById('ctx-menu');
      if (!m || !m.contains(this)) return 0;
      const 全 = [...m.querySelectorAll('.ctx-item, .ctx-divider')];
      let y = 0;
      for (const e of 全) {
        if (e === this) return y;
        if (e.style && e.style.display === 'none') continue;
        y += e.classList.contains('ctx-item') ? H : 9;
      }
      return y;
    },
  });
  Object.defineProperty(proto, 'clientHeight', {
    configurable: true,
    get() { return this.id === 'ctx-menu' ? 箱の高さ : 0; },
  });
};
const 高さを戻す = () => {
  const proto = win.HTMLElement.prototype;
  for (const k of ['offsetHeight', 'offsetTop', 'clientHeight']) {
    Object.defineProperty(proto, k, { configurable: true, get() { return 0; } });
  }
};

for (const [箱, 一覧のまま, 一行に] of [[704, 19, 20], [400, 10, 11]]) {
  T('★箱 ' + 箱 + 'px … ' + 一覧のまま + '個は 一覧のまま★', () => {
    高さを作る(箱);
    try {
      決まりを置く(一覧のまま);
      出す();
      言う(win.__validListMode === '一覧',
        '短いのに 1行に した（今 ' + win.__validListMode + '）');
    } finally { 高さを戻す(); }
  });
  T('★箱 ' + 箱 + 'px … ' + 一行に + '個で 1行に なる★', () => {
    高さを作る(箱);
    try {
      決まりを置く(一行に);
      出す();
      言う(win.__validListMode === '1行',
        '長いのに 一覧のまま（今 ' + win.__validListMode + '）＝切り取りが 押せなく なる');
    } finally { 高さを戻す(); }
  });
}

T('★箱が 低いほど 早く 1行に なる（個数 固定では ない）★', () => {
  const 境目 = (箱) => {
    高さを作る(箱);
    try {
      for (let n = 1; n <= 40; n++) {
        決まりを置く(n); 出す();
        if (win.__validListMode !== '一覧') return n - 1;
      }
      return 40;
    } finally { 高さを戻す(); }
  };
  const 高い = 境目(704), 低い = 境目(400);
  言う(低い < 高い, '箱が 低いのに 境目が 同じ／広い（高い=' + 高い + ' 低い=' + 低い + '）');
});

/* ── ★個数で 決めていない事★（jsdom は 高さを 持たないので ★字で 見る★） ──
   ★窓の 高さで 境目が 動く★＝実ブラウザ実測 720で19個／560で14個。
   ⇒★`items.length <= 数` の 形で 決めたら 赤★（小さい 画面で また 壊れる） */
T('★一覧の 切り替えを「個数」で 決めていない★', () => {
  const 素 = fs.readFileSync(srcPath('book.html'), 'utf8');
  const i = 素.indexOf('window.__validListSwitch');
  const j = 素.indexOf('};', i);
  const 中 = 素.slice(i, j);
  言う(i >= 0, '切り替えの 手が 見つからない');
  言う(!/items\.length\s*[<>=]/.test(中),
    '★個数で 決めている★（items.length と 数を 比べている）＝小さい 画面で 壊れる');
  言う(/offsetTop/.test(中) && /clientHeight|箱/.test(中),
    '★「切り取りが 見えるか」で 決めていない★');
});

/* ── わざと 壊して 赤に なるか ── */
/* ★`if (SELF) {` では tests/name-vs-body.test.mjs が 見つけられない★（2026-09-03 実測）
   ＝あの 見張りは 決まった 書き方だけを 探す。★書き方を 合わせる★ */
if (process.argv.includes('--self-test')) {
  console.log('');
  console.log('★わざと 壊して 赤に なるか★');
  const 壊す = [
    ['★灰を 付けない（リンクが 無くても 押せる）★',
      (s) => s.replace("else { _lo.classList.add('ctx-off');", "else { _lo.classList.remove('ctx-off');")],
    ['★測れない時に 一覧のまま に する（安全側を やめる）★',
      (s) => s.replace("if(!kiri || !箱){ 一行にする(); window.__validListMode='1行(測れない)'; return; }",
        "if(!kiri || !箱){ return; }")],
    /* ★jsdom は 高さを 持たない★ので「見える/見えない」その物は 実ブラウザで 見る。
       ここでは ★決め方が 個数に 戻っていないか★を 字で 見張る。 */
    ['★個数で 決める（20個 固定＝小さい 画面で 壊れる）★',
      (s) => s.replace('var 見える = (kiri.offsetTop + kiri.offsetHeight) <= 箱;',
        'var 見える = vr.items.length <= 20;')],
  ];
  const 元 = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
  const tmp = path.join(ROOT, 'tests', '_ctxa_broken.html');
  for (const [名, f] of 壊す) {
    const 壊れ = f(元);
    if (壊れ === 元) { console.log('  ★素通り★  ' + 名 + '（印が 古い＝直せ）'); fail++; continue; }
    fs.writeFileSync(tmp, 壊れ);
    const { execFileSync } = await import('node:child_process');
    let 赤 = false;
    try {
      execFileSync(process.execPath, [path.join(ROOT, 'tests', 'ctx-menu-a-ui.test.mjs')], {
        env: { ...process.env, EXALLY_CTXA_OVERRIDE: JSON.stringify({ 'book.html': tmp }) },
        stdio: 'pipe',
      });
    } catch (e) { 赤 = true; }
    console.log((赤 ? '  赤くなった  ' : '  ★素通り★  ') + 名);
    if (!赤) fail++;
  }
  try { fs.unlinkSync(tmp); } catch (e) { /* 無くてよい */ }
}

console.log('');
console.log('ctx-menu-a-ui: ' + pass + ' 緑 / ' + fail + ' 赤');
process.exit(fail ? 1 : 0);
