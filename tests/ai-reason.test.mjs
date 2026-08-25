/* ai-reason.test.mjs — ★AIに繋がらない時に 理由と 次の一手を出すか★
 *
 *  真値は tests/fixtures/ai-endpoint-golden.json
 *  （★お金を使わずに実配信を叩いて測った★：GETは405、中身なしPOSTは400で
 *    Anthropic に届く前に返る）。
 *
 *  ★指示役が実配信で見つけた事（2026-08-21）★
 *    空の A1 で「AIに解説させる」を押すと
 *    「AIに接続できなかったよ。もう一度試してみてね。」だけが出た。
 *    ★何度 押しても 失敗する（テスト線には AIの窓口が そもそも無い）★。
 *    ★空のセルでも AIを呼んでいた＝お金を使っていた★。
 *
 *  ★この検査では 本物のネットを使わない★（fetch を偷って数える）。
 *
 *  --self-test … わざと壊して「何通りで赤くなるか」を数える（★repo は読むだけ★）
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const SELF = process.argv.includes('--self-test');
const OVERRIDE = process.env.EXALLY_AI_OVERRIDE ? JSON.parse(process.env.EXALLY_AI_OVERRIDE) : {};
const srcPath = (rel) => OVERRIDE[rel] || path.join(ROOT, rel);

let JSDOM;
try { ({ JSDOM } = await import('jsdom')); }
catch { console.log('★jsdomが入っていません。この検証は飛ばせません（SKIPを緑と呼ばない）。npm install してください。'); process.exit(1); }

const GOLD = JSON.parse(fs.readFileSync(path.join(ROOT, 'tests/fixtures/ai-endpoint-golden.json'), 'utf8'));

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ok   ' + n); } catch (e) { fail++; console.log('  NG   ' + n + '\n       ' + (e && e.message)); } };
const ok = (c, m) => { if (!c) throw new Error(m || 'expected truthy'); };
const eq = (a, b, m) => { if (a !== b) throw new Error((m ? m + ': ' : '') + 'expected ' + JSON.stringify(b) + ' got ' + JSON.stringify(a)); };

const CANVAS_STUB = [
  '(function(){var noop=function(){};var ctx=new Proxy({},{get:function(t,k){',
  ' if(k==="measureText")return function(){return{width:40};};',
  ' if(k==="canvas")return{width:900,height:600};',
  ' if(k==="getImageData")return function(){return{data:[]};};',
  ' if(k==="createLinearGradient"||k==="createPattern")return function(){return{addColorStop:noop};};',
  ' return noop;}});HTMLCanvasElement.prototype.getContext=function(){return ctx;};})();',
].join('\n');

/* ★開いた窓と 刷った回数を数える偽物★（本物のブラウザの窓は出さない） */
const WINDOW_SPY = [
  'window.__開いた窓 = [];',
  'window.open = function(){',
  '  var 書いた = "";',
  '  var w = { __刷った: 0, focus: function(){}, print: function(){ this.__刷った++; },',
  '            document: { open:function(){}, close:function(){}, write:function(s){ 書いた += s; } },',
  '            get 中身(){ return 書いた; } };',
  '  window.__開いた窓.push(w); return w;',
  '};',
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
inject(WINDOW_SPY);

console.log('\n[ai-reason] ★AIに繋がらない時に 理由と 次の一手を出すか★');
console.log('  真値 = ' + GOLD._measured_with);

const G4 = GOLD['④繋がらない時の言い方']['後'];
const sh = () => win.sheets[win.activeSheet];
function reset() { sh().data = {}; win.sel(0, 0, 0, 0); }
const 窓の字 = () => doc.getElementById('explain-text').textContent.replace(/\s+/g, ' ').trim();

/* ★AIを呼びに行ったかを 数える偽物★（本物のネットは使わない＝お金を使わない） */
let 呼んだ回数 = 0, 次に返す = null;
win.fetch = function (url, opt) {
  呼んだ回数++;
  if (次に返す === 'ネット切れ') return Promise.reject(new Error('failed to fetch'));
  const r = 次に返す || { status: 200, body: { text: 'これは合計です' } };
  return Promise.resolve({
    status: r.status, ok: r.status >= 200 && r.status < 300,
    json: () => (r.body === undefined ? Promise.reject(new Error('not json')) : Promise.resolve(r.body)),
    text: () => Promise.resolve(JSON.stringify(r.body || '')),
  });
};
const 押す = async (返す, fn) => { 次に返す = 返す; 呼んだ回数 = 0; await fn(); 次に返す = null; };

/* ★待つ検査★ … T() は同期なので async を渡すと ★中で落ちても 緑のまま素通りする★。
   ここは 必ず await する形にして 数える。 */
const AT = async (n, fn) => {
  try { await fn(); pass++; console.log('  ok   ' + n); }
  catch (e) { fail++; console.log('  NG   ' + n + '\n       ' + (e && e.message)); }
};

/* ── 土台 ── */
T('★画面が立ち上がっていて AiReason と 呼び口が在る', () => {
  ok(win.AiReason && typeof win.AiReason.読む === 'function', 'AiReason が無い');
  ok(typeof win._aiFetch === 'function', '呼び口 _aiFetch が無い');
  ok(typeof win.explainCell === 'function', 'explainCell が無い');
});
T('★AIを呼ぶ口は 1本だけ（言い方がばらけない）★', () => {
  const n = (html.match(/fetch\('\/api\/claude'/g) || []).length;
  eq(n, 1, '★呼ぶ口が ' + n + 'か所 在る（前は3か所で 同じ言い訳を出していた）★');
});
T('★古い言い方が コードに残っていない★', () => {
  /* 注記（コメント）の中は 残っていてよい。★動く所★に残っていないかを見る */
  const 動く所 = html.replace(/\/\*[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '');
  const n = (動く所.match(/AIに接続できなかったよ/g) || []).length;
  eq(n, 0, '★「もう一度試してみてね」だけの言い方が ' + n + '件 残っている★');
});

/* ── ① 理由と 次の一手（実配信で測った物と 同じ言葉か）── */
const 組 = [
  ['テスト線 404', { status: 404, 本番か: false }, G4['404 / 405（テスト線）']],
  ['テスト線 405', { status: 405, 本番か: false }, G4['404 / 405（テスト線）']],
  ['本番 405', { status: 405, 本番か: true }, G4['404 / 405（本番）']],
  ['鍵 401', { status: 401, 本番か: true }, G4['401 / 403']],
  ['混雑 429', { status: 429, 本番か: true }, G4['429']],
  ['中身なし 400', { status: 400, 本番か: true }, G4['400']],
  ['AI側 500', { status: 500, 本番か: true }, G4['500以上']],
  ['ネット切れ', { status: null, ネット切れ: true, 本番か: true }, G4['ネット切れ']],
  ['空の返事 200', { status: 200, 中身: '', 本番か: true }, G4['200だが 中身が空']],
];
for (const [名, 入れる, 期待] of 組) {
  T('★' + 名 + ' … 理由と 次の一手を 出す★', () => {
    const r = win.AiReason.読む(入れる);
    eq(r.ok, false, 'つながった事にしている');
    const 待つ = 期待.split('／');
    ok(r.理由.indexOf(待つ[0].trim().replace(/^★|★$/g, '').slice(0, 8)) >= 0 || 待つ[0].indexOf(r.理由.slice(0, 8)) >= 0,
      '理由が違う：' + r.理由);
    ok(r.次.length > 0, '★次にどうすればよいかを 出していない★');
    eq(r.言葉, r.理由 + '\n' + r.次, '出す言葉が 組み立てられていない');
  });
}
T('★ちゃんと返った時は そのまま出す（余計な事を言わない）★', () => {
  const r = win.AiReason.読む({ status: 200, 中身: 'A1は合計です', 本番か: true });
  eq(r.ok, true, 'つながっていないと言っている');
  eq(r.言葉, 'A1は合計です', '中身を書き換えている');
});
T('★押し直しても直らない物には「もう一度」と言わない★', () => {
  /* ★本番の404も 直らない★（窓口が配信されていない＝押し直しても出ない） */
  for (const 入れる of [{ status: 404, 本番か: false }, { status: 404, 本番か: true },
                        { status: 405, 本番か: true }, { status: 401, 本番か: true }, { status: 403, 本番か: true }]) {
    const r = win.AiReason.読む(入れる);
    ok(r.次.indexOf('もう一度 押してみてね') < 0,
      '★直らないのに「もう一度」と言っている★：' + r.次);
  }
});
T('★テスト線と本番で 次の一手が 違う（同じ事しか言わないと 直せない）★', () => {
  const テスト線 = win.AiReason.読む({ status: 404, 本番か: false });
  const 本番 = win.AiReason.読む({ status: 404, 本番か: true });
  ok(テスト線.次 !== 本番.次, '★どちらでも 同じ事しか言っていない★：' + テスト線.次);
  ok(テスト線.次.indexOf('テスト環境') >= 0, 'テスト線で 何をすればよいか 出していない：' + テスト線.次);
  ok(本番.次.indexOf('管理者') >= 0, '本番で 誰に言えばよいか 出していない：' + 本番.次);
});
T('★空白だけのセルは 中身が無い扱い（お金を使わない）★', () => {
  eq(win.AiReason.呼んでよいか(['   ']).呼ぶ, false, '★空白だけで AIを呼んでいる★');
  eq(win.AiReason.呼んでよいか(['　']).呼ぶ, false, '★全角の空白だけで AIを呼んでいる★');
  eq(win.AiReason.呼んでよいか(['', '\n']).呼ぶ, false, '★改行だけで AIを呼んでいる★');
  eq(win.AiReason.呼んでよいか(['0']).呼ぶ, true, '0 は 中身として数える');
});
/* ── ★合言葉（サーバが「なぜ失敗したか」を そのまま返す）★ ── */
const G5 = GOLD['⑤★失敗しても 200 を返していた（2026-08-22 に見つけて直した）★'];
T('★残高切れは 400ではなく 402＋合言葉 zandaka で来る（400のままだと「送る中身が足りません」と嘘を言う）★', () => {
  const 嘘 = win.AiReason.読む({ status: 400, 本番か: true });
  ok(嘘.理由.indexOf('送る中身') >= 0, '400 の読み方が 変わっている：' + 嘘.理由);
  const 本当 = win.AiReason.読む({ status: 402, 理由: 'zandaka', 本番か: true });
  ok(本当.理由.indexOf('足りません') >= 0, '残高切れを 読めていない：' + 本当.理由);
  ok(本当.次.indexOf('もう一度') < 0, '★残高が尽きているのに「もう一度」と言っている★：' + 本当.次);
  ok(本当.次.indexOf('残高') >= 0, '誰に何を頼めばよいか 出していない：' + 本当.次);
});
T('★番号より 合言葉を 先に読む（Anthropic自身は 残高切れを 400 で返す）★', () => {
  /* ★境界★ うちのサーバは 400→402 に分けて返すが、分け損ねて 400 のまま来た時でも
     合言葉が在れば ★正しく「足りません」と言えなければならない★（番号だけ見ると 嘘になる） */
  const r = win.AiReason.読む({ status: 400, 理由: 'zandaka', 本番か: true });
  ok(r.理由.indexOf('足りません') >= 0, '★番号(400)を先に読んで 嘘を言っている★：' + r.理由);
  ok(r.次.indexOf('もう一度') < 0, '★直らないのに「もう一度」と言っている★：' + r.次);
  const t = win.AiReason.読む({ status: 502, 理由: 'jikangire', 本番か: true });
  ok(t.理由.indexOf('時間切れ') >= 0, '★合言葉 jikangire を 読んでいない★：' + t.理由);
});
T('★合言葉の5つが 番号だけの時と 同じ言葉になる（言い方が2通りに割れない）★', () => {
  const 組 = [['kagi', 401], ['komiai', 429], ['jikangire', 504], ['ai_shippai', 502], ['zandaka', 402]];
  for (const [合言葉, 番号] of 組) {
    const a = win.AiReason.読む({ status: 番号, 理由: 合言葉, 本番か: true });
    const b = win.AiReason.読む({ status: 番号, 本番か: true });
    eq(a.言葉, b.言葉, '★' + 合言葉 + ' の言い方が 番号だけの時と 違う★');
    ok(a.ok === false && a.次.length > 0, '次の一手が無い：' + 合言葉);
  }
});
T('★真値(golden)に書いた 番号と合言葉の組み合わせが そのまま在る★', () => {
  const 後 = G5['後'];
  const 待つ = { zandaka: 402, kagi: 401, komiai: 429, jikangire: 504, ai_shippai: 502 };
  let 数 = 0;
  for (const 行 of Object.values(後)) {
    const m = String(行).match(/★(\d{3}) [{]"error":"([a-z_]+)"[}]★/);
    ok(m, '真値の書き方が変わった：' + 行);
    eq(待つ[m[2]], Number(m[1]), '★' + m[2] + ' の番号が 真値と違う★');
    数++;
  }
  eq(数, 5, '★真値の行が 5つでない★');
});
T('★中の言葉を 客に見せない（Vercel・環境変数・APIキーの字を出さない）★', () => {
  const 全部 = [{ status: 401, 本番か: true }, { status: 402, 理由: 'zandaka', 本番か: true },
                { status: 429, 本番か: true }, { status: 504, 本番か: true },
                { status: 502, 本番か: true }, { status: 400, 本番か: true },
                { status: 404, 本番か: false }, { status: 404, 本番か: true },
                { status: null, ネット切れ: true, 本番か: true }];
  for (const 入れる of 全部) {
    const 言葉 = win.AiReason.読む(入れる).言葉;
    for (const 中の字 of ['Vercel', 'Environment', 'ANTHROPIC_API_KEY', 'APIキー', 'undefined']) {
      ok(言葉.indexOf(中の字) < 0, '★中の言葉「' + 中の字 + '」が 客に出ている★：' + 言葉);
    }
  }
});
T('★押し直しても直らない物に「もう一度」と言わない（鍵・残高を足した）★', () => {
  for (const 入れる of [{ status: 401, 本番か: true }, { status: 402, 理由: 'zandaka', 本番か: true },
                        { status: 403, 本番か: true }]) {
    const r = win.AiReason.読む(入れる);
    ok(r.次.indexOf('もう一度') < 0, '★直らないのに「もう一度」と言っている★：' + r.次);
    ok(r.次.indexOf('管理者') >= 0, '★誰に言えばよいかを 出していない★：' + r.次);
  }
});
T('★客に見せる字に ★ を書かない（★は うちの覚え書きの印）★', () => {
  const 全部 = [{ status: 404, 本番か: false }, { status: 404, 本番か: true }, { status: 401, 本番か: true },
                { status: 402, 理由: 'zandaka', 本番か: true }, { status: 429, 本番か: true },
                { status: 504, 本番か: true }, { status: 502, 本番か: true }, { status: 400, 本番か: true },
                { status: null, ネット切れ: true, 本番か: true }, { status: 200, 中身: '', 本番か: true }];
  for (const 入れる of 全部) {
    const 言葉 = win.AiReason.読む(入れる).言葉;
    ok(言葉.indexOf('★') < 0, '★客の字に ★ が出ている★：' + 言葉);
  }
  ok(win.AiReason.呼んでよいか(['']).言葉.indexOf('★') < 0, '★空の時の字に ★ が出ている★');
});
T('★押し直せば直る物には ちゃんと「もう一度」と言う★', () => {
  for (const 入れる of [{ status: 500, 本番か: true }, { status: 200, 中身: '', 本番か: true }]) {
    ok(win.AiReason.読む(入れる).次.indexOf('もう一度') >= 0, '次の一手が無い');
  }
});

/* ── ② ★中身が無い物に AIを呼ばない（お金を使わない）★ ── */
T('★空のセルでは AIを呼ばない★', () => {
  reset();
  win.sel(0, 0, 0, 0);
  呼んだ回数 = 0;
  win.explainCell();
  eq(呼んだ回数, 0, '★空のセルなのに AIを呼んだ（お金を使う）★');
  ok(窓の字().indexOf('何も入っていません') >= 0, '理由を出していない：' + 窓の字());
  ok(窓の字().indexOf('選んでから') >= 0, '次の一手を出していない：' + 窓の字());
});
T('★空の範囲でも AIを呼ばない★', () => {
  reset();
  win.sel(0, 0, 3, 3);
  呼んだ回数 = 0;
  win.explainCell();
  eq(呼んだ回数, 0, '★空の範囲なのに AIを呼んだ★');
});
T('★字が在れば ちゃんと呼ぶ（呼ばなくなっていない）★', () => {
  reset();
  win.setCell(0, 0, '100');
  win.sel(0, 0, 0, 0);
  呼んだ回数 = 0;
  win.explainCell();
  eq(呼んだ回数, 1, '★字が在るのに AIを呼ばない＝機能が死んでいる★');
});
T('★式だけのセルでも 呼ぶ★', () => {
  reset();
  win.setCell(0, 0, '=1+2');
  win.sel(0, 0, 0, 0);
  呼んだ回数 = 0;
  win.explainCell();
  eq(呼んだ回数, 1, '式のセルで 呼ばない');
});
T('★上限や 既定オフを 足していない（勝手に決めない）★', () => {
  for (const 語 of ['maxCalls', 'AI_LIMIT', 'aiDisabled', '上限', '1日', '回まで']) {
    ok(html.indexOf('var ' + 語) < 0, '★使う上限を 勝手に足している：' + 語 + '★');
  }
});

/* ── ③ 呼び口が 理由を返すか（本物の _aiFetch を押す）── */
await AT('★呼び口が 404 を「窓口が無い」と読む（実際に押す）★', async () => {
  await 押す({ status: 404 }, async () => {
    const r = await win._aiFetch({ message: 'x', history: [] });
    eq(r.ok, false, 'つながった事にしている');
    ok(r.言葉.indexOf('窓口') >= 0, '理由が出ていない：' + r.言葉);
  });
});
await AT('★呼び口が ネット切れを 読む（実際に押す）★', async () => {
  await 押す('ネット切れ', async () => {
    const r = await win._aiFetch({ message: 'x', history: [] });
    eq(r.ok, false, 'つながった事にしている');
    ok(r.言葉.indexOf('ネット') >= 0, '理由が出ていない：' + r.言葉);
  });
});
await AT('★呼び口が 200 の中身を そのまま返す（実際に押す）★', async () => {
  await 押す({ status: 200, body: { text: 'これは合計です' } }, async () => {
    const r = await win._aiFetch({ message: 'x', history: [] });
    eq(r.ok, true, 'つながっていないと言っている');
    eq(r.言葉, 'これは合計です', '中身を書き換えている');
  });
});

/* ── ④ チャットも 実際に押す（解説の窓だけでなく）── */
const チャットの字 = () => (doc.getElementById('ai-response-area') || {}).textContent || '';
await AT('★PCのチャットで 繋がらない時 理由と 次の一手が出る（実際に押す）★', async () => {
  doc.getElementById('ai-response-area').innerHTML = '';
  doc.getElementById('ai-input').value = 'これ何？';
  次に返す = { status: 404 };
  await win.sendToAI();
  次に返す = null;
  const t = チャットの字();
  ok(t.indexOf('窓口') >= 0, '★理由が出ていない★：' + t.slice(0, 120));
  ok(t.indexOf('もう一度試してみてね') < 0, '★古い言い方のまま★');
  eq(doc.getElementById('ai-send').disabled, false, '★押せないままになっている★');
});
await AT('★PCのチャットで 繋がった時は 答えを出す★', async () => {
  doc.getElementById('ai-response-area').innerHTML = '';
  doc.getElementById('ai-input').value = 'これ何？';
  次に返す = { status: 200, body: { text: 'これは合計です' } };
  await win.sendToAI();
  次に返す = null;
  ok(チャットの字().indexOf('これは合計です') >= 0, '★答えが出ていない＝チャットを殺した★');
});

await AT('★繋がらなかった時は 会話の記録に 残さない（次の質問に エラー文を送らない）★', async () => {
  win.aiHistory.length = 0;
  doc.getElementById('ai-response-area').innerHTML = '';
  doc.getElementById('ai-input').value = 'これ何？';
  次に返す = { status: 404 };
  await win.sendToAI();
  次に返す = null;
  eq(win.aiHistory.length, 0,
    '★繋がらなかったのに 会話の記録に ' + win.aiHistory.length + '件 残している'
    + '（次の質問で エラー文を AIの答えとして送る）★：' + JSON.stringify(win.aiHistory));
});
await AT('★繋がった時は 会話の記録に 残る（記録を殺していない）★', async () => {
  win.aiHistory.length = 0;
  doc.getElementById('ai-response-area').innerHTML = '';
  doc.getElementById('ai-input').value = 'これ何？';
  次に返す = { status: 200, body: { text: 'これは合計です' } };
  await win.sendToAI();
  次に返す = null;
  eq(win.aiHistory.length, 2, '★会話の記録が 残っていない★');
});

await AT('★残高切れ(402)を 実際に押すと 画面に 理由と 次の一手が出る★', async () => {
  reset();
  win.setCell(0, 0, '100');
  win.sel(0, 0, 0, 0);
  /* ★番号は 400（＝そのままなら「送る中身が足りません」）。合言葉を読んで初めて正しく言える★ */
  await 押す({ status: 400, body: { error: 'zandaka', text: '', tsv: '' } }, async () => {
    win.explainCell();
    await new Promise((r) => setTimeout(r, 30));
  });
  const 字 = 窓の字();
  ok(字.indexOf('足りません') >= 0, '★残高切れの理由が 画面に出ていない★：' + 字);
  ok(字.indexOf('残高を足してほしい') >= 0, '★次の一手が 画面に出ていない★：' + 字);
  ok(字.indexOf('もう一度') < 0, '★直らないのに「もう一度」と 画面に出ている★：' + 字);
});
await AT('★繋がらなかった時は「深掘り」を出さない（押しても また失敗する）★', async () => {
  reset();
  win.setCell(0, 0, '100');
  win.sel(0, 0, 0, 0);
  await 押す({ status: 404 }, async () => {
    win.explainCell();
    await new Promise((r) => setTimeout(r, 30));
  });
  const deep = doc.getElementById('explain-deep');
  ok(deep, '深掘りボタンが 見つからない');
  eq(deep.style.display, 'none', '★繋がっていないのに 深掘りを出している★');
  const copy = [...doc.querySelectorAll('#explain-popup .explain-footer button')]
    .find((b) => b.textContent.indexOf('コピー') >= 0);
  ok(copy && copy.style.display !== 'none', '★コピーまで消している（管理者に貼れなくなる）★');
});
await AT('★空のセルでも「深掘り」を出さない★', async () => {
  reset();
  win.sel(0, 0, 0, 0);
  win.explainCell();
  eq(doc.getElementById('explain-deep').style.display, 'none', '★中身が無いのに 深掘りを出している★');
});
await AT('★繋がった時は 深掘りが 出る（機能を殺していない）★', async () => {
  reset();
  win.setCell(0, 0, '100');
  win.sel(0, 0, 0, 0);
  await 押す({ status: 200, body: { text: 'これは100だよ' } }, async () => {
    win.explainCell();
    await new Promise((r) => setTimeout(r, 30));
  });
  eq(doc.getElementById('explain-deep').style.display, '', '★繋がったのに 深掘りが 消えている★');
});
await AT('★鍵がだめ(401)を 実際に押しても 中の言葉（Vercel…）は 出ない★', async () => {
  reset();
  win.setCell(0, 0, '100');
  win.sel(0, 0, 0, 0);
  await 押す({ status: 401, body: { error: 'kagi', text: '', tsv: '' } }, async () => {
    win.explainCell();
    await new Promise((r) => setTimeout(r, 30));
  });
  const 字 = 窓の字();
  ok(字.indexOf('鍵') >= 0, '★理由が 画面に出ていない★：' + 字);
  for (const 中の字 of ['Vercel', 'Environment', 'ANTHROPIC']) {
    ok(字.indexOf(中の字) < 0, '★中の言葉「' + 中の字 + '」が 画面に出ている★：' + 字);
  }
});

console.log('\n  ' + pass + ' 緑 / ' + fail + ' 赤');

if (SELF) {
  const { spawnSync } = await import('node:child_process');
  const os = await import('node:os');
  const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'exally-ai-'));
  console.log('\n[self-test] わざと壊して 赤くなるかを数える（★repo は読むだけ★）');
  const BREAKS = [
    ['book.html', '★空のセルでも AIを呼ぶ（お金を使う）★',
      (s) => s.replace('  if(!可否.呼ぶ){', '  if(false){')],
    ['book.html', '★中身を見ずに いつも呼ぶ（守りごと外す）★',
      (s) => s.replace('  var 可否 = AiReason.呼んでよいか(送る字);', '  var 可否 = { 呼ぶ: true, 言葉: "" };')],
    ['book.html', '★字が在っても 呼ばない（機能を殺す）★',
      (s) => s.replace('  var 可否 = AiReason.呼んでよいか(送る字);', '  var 可否 = { 呼ぶ: false, 言葉: "だめ" };')],
    ['book.html', '★理由を捨てて 古い言い方に戻す★',
      (s) => s.replace('  showExplainPopup(_explainAddr, _explainFormula, r.言葉, r.ok);',
        "  showExplainPopup(_explainAddr, _explainFormula, 'AIに接続できなかったよ。もう一度試してみてね。');")],
    ['book.html', '★呼ぶ口を もう1本 増やす（言い方がばらける）★',
      (s) => s.replace('async function _sendExplain(prompt){',
        "async function _sendExplainOld(prompt){ await fetch('/api/claude',{method:'POST'}); }\nasync function _sendExplain(prompt){")],
    ['book.html', '★使う上限を 勝手に足す（足すなと言われている）★',
      (s) => s.replace('async function _aiFetch(body){', 'var AI_LIMIT = 10;\nasync function _aiFetch(body){')],
    ['book.html', '★つながっていないのに つながった事にする★',
      (s) => s.replace('    if(!r.ok){ addAIChatMsg(\'ai\', aiText); document.getElementById(\'ai-send\').disabled = false; return; }', '')],
    ['lib/ai-reason.js', '★窓口が無い(404)を 読まない★',
      (s) => s.replace('    if (s === 404 || s === 405) {', '    if (false) {')],
    ['lib/ai-reason.js', '★鍵がだめ(401)を 読まない★',
      (s) => s.replace('    if (s === 401 || s === 403) return 鍵();', '')],
    ['lib/ai-reason.js', '★混んでいる(429)を 読まない★',
      (s) => s.replace('    if (s === 429) return 混雑();', '')],
    ['lib/ai-reason.js', '★AI側の失敗(500)を 読まない★',
      (s) => s.replace('    if (s && s >= 500) return AI側();', '')],
    ['lib/ai-reason.js', '★残高切れ(402)を 読まない★',
      (s) => s.replace('    if (s === 402) return 残高();', '')],
    ['lib/ai-reason.js', '★合言葉を 読まない（番号だけ見る）★',
      (s) => s.replace("    if (合言葉 === 'zandaka') return 残高();", '')],
    ['lib/ai-reason.js', '★残高切れなのに「もう一度」と言う★',
      (s) => s.replace("      '押し直しても 直りません。管理者に「AIの残高を足してほしい」と伝えてね。');",
        "      'しばらくしてから もう一度 押してみてね。');")],
    ['lib/ai-reason.js', '★鍵の話で 中の言葉（Vercel）を 客に見せる★',
      (s) => s.replace("      '押し直しても 直りません。管理者に「AIの鍵を見てほしい」と伝えてね。');",
        "      'VercelのEnvironment VariablesにANTHROPIC_API_KEYを設定してください。');")],
    ['book.html', '★繋がらなくても 深掘りを 見せる★',
      (s) => s.replace("  if(deep) deep.style.display = (繋がったか === false) ? 'none' : '';", '')],
    ['book.html', '★繋がった時まで 深掘りを 消す★',
      (s) => s.replace("  if(deep) deep.style.display = (繋がったか === false) ? 'none' : '';",
        "  if(deep) deep.style.display = 'none';")],
    ['lib/ai-reason.js', '★客の字に ★ を書く★',
      (s) => s.replace("'テスト環境には AIが在りません（練習用の倉庫だけ）。AIは 本番の画面で使ってね。'",
        "'★テスト環境には AIが在りません（練習用の倉庫だけ）。AIは 本番の画面で使ってね。★'")],
    ['book.html', '★失敗した時に 合言葉を 読まない★',
      /* ★2026-08-25：★4 事故止めで この1行に 待ち秒・どれ が足された★
         ⇒ 行まるごとではなく ★合言葉を読む所だけ★を消す（置換できずに素通りしていた） */
      (s) => s.replace("理由 = (e2 && e2.error) || '';", "理由 = '';")],
    ['lib/ai-reason.js', '★ネット切れを 読まない★',
      (s) => s.replace('    if (o.ネット切れ) {', '    if (false) {')],
    ['lib/ai-reason.js', '★空の返事でも つながった事にする★',
      (s) => s.replace('      if (!o.中身) {', '      if (false) {')],
    ['lib/ai-reason.js', '★次にどうすればよいかを 出さない★',
      (s) => s.replace("    return { ok: false, 理由: 理由, 次: 次, 言葉: 理由 + '\\n' + 次 };",
        "    return { ok: false, 理由: 理由, 次: '', 言葉: 理由 };")],
    ['lib/ai-reason.js', '★直らない物にも「もう一度」と言う★',
      (s) => s.replace("        o.本番か ? '管理者に「/api/claude が配信されていない」と伝えてね。'",
        "        o.本番か ? 'もう一度 押してみてね。'")],
    ['lib/ai-reason.js', '★テスト線と本番で 同じ事しか言わない★',
      (s) => s.replace("                 : 'テスト環境には AIが在りません（練習用の倉庫だけ）。AIは 本番の画面で使ってね。');",
        "                 : '管理者に「/api/claude が配信されていない」と伝えてね。');")],
    ['lib/ai-reason.js', '★空でも 呼んでよい事にする★',
      (s) => s.replace('    if (何か在る) return { 呼ぶ: true, 言葉: \'\' };', '    return { 呼ぶ: true, 言葉: \'\' };')],
    ['lib/ai-reason.js', '★中身が在る時に 空白だけを 中身と数える★',
      (s) => s.replace("      return s !== undefined && s !== null && String(s).trim() !== '';",
        "      return s !== undefined && s !== null && String(s) !== '';")],
  ];
  let red = 0;
  for (const [rel, name, brk] of BREAKS) {
    const orig = fs.readFileSync(path.join(ROOT, rel), 'utf8');
    const bad = brk(orig);
    if (bad === orig) { console.log('  ★置換できず★  ' + name); continue; }
    const tmpFile = path.join(TMP, rel.replace(/[\\/]/g, '_'));
    fs.writeFileSync(tmpFile, bad, 'utf8');
    const env = Object.assign({}, process.env, { EXALLY_AI_OVERRIDE: JSON.stringify({ [rel]: tmpFile }) });
    const isRed = spawnSync(process.execPath, [path.join(__dirname, 'ai-reason.test.mjs')], { encoding: 'utf8', env }).status !== 0;
    if (isRed) { red++; console.log('  赤くなった  ' + name); }
    else console.log('  ★素通り★  ' + name);
  }
  /* ★repo を書き換えていない事を 押した後に 数える★ */
  for (const rel of ['book.html', 'lib/ai-reason.js']) {
    /* ★注記(コメント)の中は 見ない★＝「昔こう出していた」と書き残す事は 正しい。
       見るのは ★動く所★だけ（2026-08-22：頭の注記に引っかかって 嘘の赤が出た） */
    const now = fs.readFileSync(path.join(ROOT, rel), 'utf8')
      .replace(/\/\*[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '');
    if (now.includes('AI_LIMIT') || now.includes('_sendExplainOld') || now.includes('if (false) {')
        || now.includes('VercelのEnvironment Variables')) {
      console.log('  ★NG★ ' + rel + ' に わざと壊した物が残っている');
      process.exit(1);
    }
  }
  try { fs.rmSync(TMP, { recursive: true, force: true }); } catch (e) { /* 消せなくても検査は済んでいる */ }
  console.log('\n  ' + red + '/' + BREAKS.length + ' 通りで赤くなった');
  process.exit(red === BREAKS.length ? 0 : 1);
}

process.exit(fail ? 1 : 0);
