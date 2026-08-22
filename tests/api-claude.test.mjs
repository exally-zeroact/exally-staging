/* api-claude.test.mjs — チャットのサーバ側が返す【基準数値】を実数で固定する。
 *
 * なぜ必要か（2026-08-02の事故から）:
 *   ・api/claude.js は、法定の料率を kyuyo/lib/ から読んで、システムプロンプトに埋め込んでいる。
 *   ・読み先が消えた時は 500 で【派手に】落ちたので気づけた（refs-resolve.test.mjs が根を止めた）。
 *   ・しかしこの手の埋め込みは、libのAPIが変わった時に **NaN や undefined が静かに入る** 方が怖い。
 *     画面は普通に出るし、AIも普通に喋る。ただ「健康保険料率: NaN%」と客に言うだけ。
 *   ★だから数値そのものを機械が見る。
 *
 * 判定:
 *   ① 対象月を固定した時に、実際の官公値がそのまま出る（実数リテラルで固定）
 *   ② 年度が変わる境界（社保=3月起算 / 労働保険=4月起算）で、ちゃんと切り替わる
 *   ③ ★「今日」で組み立てても NaN / undefined / Infinity が混ざらない
 *
 * 使い方: node tests/api-claude.test.mjs
 */
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(import.meta.url);

if (!process.env.ANTHROPIC_API_KEY) process.env.ANTHROPIC_API_KEY = 'test-dummy-key';
const handler = require_(path.join(ROOT, 'api/claude.js'));
const build = handler.__buildStatutoryPrompt;

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };
const has = (s, sub) => { if (s.indexOf(sub) < 0) throw new Error('出てこない: ' + sub + '\n--- 実際 ---\n' + s); };

console.log('\n[api-claude] チャットが客に言う「基準数値」が本物か');

T('★令和8年度(2026-08)の実数 — 健保(東京)4.925% / 厚年9.15% / 雇用0.50% / 消費税10%・8%', () => {
  const s = build('2026-08');
  has(s, '健康保険料率（東京）: 4.925%');   // 協会けんぽ R8 東京 9.85% の折半
  has(s, '令和8年度');
  has(s, '厚生年金保険料率: 9.15%');        // 18.3% の折半（全国一律）
  has(s, '雇用保険料率: 0.50%');            // R8 一般の事業 5.0/1000（R7の5.5から引下げ）
  has(s, '消費税: 10%（標準）/ 8%（軽減）');
});

T('★令和7年度(2025-08)の実数 — 健保(東京)4.955% / 雇用0.55%', () => {
  const s = build('2025-08');
  has(s, '健康保険料率（東京）: 4.955%');   // 協会けんぽ R7 東京 9.91% の折半
  has(s, '令和7年度');
  has(s, '雇用保険料率: 0.55%');            // R7 一般の事業 5.5/1000
});

T('社保年度は3月起算で切り替わる（2026-02は令和7・2026-03は令和8）', () => {
  has(build('2026-02'), '令和7年度');
  has(build('2026-03'), '令和8年度');
});

T('労働保険年度は4月起算で切り替わる（2026-03は0.55%・2026-04は0.50%）', () => {
  has(build('2026-03'), '雇用保険料率: 0.55%');
  has(build('2026-04'), '雇用保険料率: 0.50%');
});

T('★「今日」で組み立てても NaN / undefined が混ざらない', () => {
  const s = build();
  for (const bad of ['NaN', 'undefined', 'Infinity', 'null']) {
    if (s.indexOf(bad) >= 0) throw new Error('プロンプトに ' + bad + ' が入っています:\n' + s);
  }
  if (!/健康保険料率（東京）: \d+\.\d{3}%/.test(s)) throw new Error('健保料率が数字になっていません:\n' + s);
  if (!/雇用保険料率: \d+\.\d{2}%/.test(s)) throw new Error('雇用保険料率が数字になっていません:\n' + s);
});

T('検査が空振りしていない（テスト用の窓が実際に生えている）', () => {
  if (typeof build !== 'function') throw new Error('__buildStatutoryPrompt が無い');
  if (typeof handler !== 'function') throw new Error('api/claude.js が関数を export していない（Vercelが呼べない）');
});

/* ── ★失敗した時に 何を返すか（2026-08-22）★ ──────────────────────────
 *  前は ★どんな失敗でも status 200★ ＋ 言い訳を text に入れて返していた。
 *  画面は 200 を「つながった」と読むので ★言い訳が AIの答えとして 吹き出しに出る★。
 *  ここは ★偽のAIを差し込んで ハンドラを本当に呼び★、返る番号と合言葉を数える。
 *  （★本物のネットは 1回も使わない＝お金を使わない★） */
const 分ける = handler.__失敗を分ける;
const 失敗 = (o) => Object.assign(new Error(o.message || 'err'), o);
const 受け皿 = () => {
  const box = { 出た: null };
  box.res = {
    setHeader() {}, status(c) { this._c = c; return this; },
    json(b) { box.出た = { status: this._c, body: b }; return this; },
    end() { box.出た = { status: this._c, body: null }; return this; },
  };
  return box;
};
const 呼ぶ = async (err) => {
  handler.__setClient({ messages: { create: async () => { throw err; } } });
  const box = 受け皿();
  await handler({ method: 'POST', body: { message: 'これ何？', history: [] }, headers: {} }, box.res);
  return box.出た;
};

console.log('');
console.log('[api-claude] ★失敗した時に 200 で「答えのふり」をしないか★');

const 表 = [
  ['残高が尽きた（Anthropicは 400 で返す）', 失敗({ status: 400, message: 'Your credit balance is too low to access the Anthropic API.' }), 402, 'zandaka'],
  ['鍵がだめ（401）', 失敗({ status: 401, message: 'invalid x-api-key' }), 401, 'kagi'],
  ['鍵がだめ（403）', 失敗({ status: 403, message: 'forbidden' }), 401, 'kagi'],
  ['鍵が置かれていない', 失敗({ message: 'The ANTHROPIC_API_KEY environment variable is missing or empty' }), 401, 'kagi'],
  ['混み合っている（429）', 失敗({ status: 429, message: 'rate limit' }), 429, 'komiai'],
  ['時間切れ（ETIMEDOUT）', 失敗({ code: 'ETIMEDOUT', message: 'timeout' }), 504, 'jikangire'],
  ['向こうが落ちた（500）', 失敗({ status: 500, message: 'overloaded' }), 502, 'ai_shippai'],
  ['何だか分からない', 失敗({ message: 'なにか' }), 502, 'ai_shippai'],
];
for (const [名, err, 待つ番号, 待つ合言葉] of 表) {
  const k = 分ける(err);
  T('★' + 名 + ' → ' + 待つ番号 + ' / ' + 待つ合言葉, () => {
    if (k.status !== 待つ番号) throw new Error('番号が ' + k.status + '（★200 を返すと 画面は つながった事にする★）');
    if (k.合言葉 !== 待つ合言葉) throw new Error('合言葉が ' + k.合言葉);
  });
}

const AT = async (n, fn) => { try { await fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };

/* ── ★誰でも叩ける口だったのを 絞る（2026-08-22 指示役）★ ──
 *  前は Access-Control-Allow-Origin: '*' ＝ よその画面からでも使えた。
 *  ★これだけでは 道具で直接叩く相手は止まらない★（それは Vercel の入口＝指示役の担当）。 */
const 押す入口 = async (origin, opts) => {
  const o = opts || {};
  handler.__setClient({ messages: { create: async () => ({ content: [{ type: 'text', text: 'はい' }], usage: { input_tokens: 12, output_tokens: 3 } }) } });
  const 出た = { headers: {}, status: null, body: null };
  const res = {
    setHeader(k, v) { 出た.headers[k] = v; },
    status(c) { 出た.status = c; return this; },
    json(b) { 出た.body = b; return this; },
    end() { return this; },
  };
  await handler({ method: o.method || 'POST', headers: origin ? { origin } : {}, body: { message: 'これ何？', history: [] } }, res);
  return 出た;
};

await AT('★うちの画面(本番)からは 今までどおり通る★', async () => {
  const r = await 押す入口('https://exally.vercel.app');
  if (r.headers['Access-Control-Allow-Origin'] !== 'https://exally.vercel.app') throw new Error('うちの画面を断っている：' + r.headers['Access-Control-Allow-Origin']);
  if (r.status !== 200) throw new Error('答えが返っていない：' + r.status);
});
await AT('★よその画面には 許しを出さない（前は * で 誰にでも出していた）★', async () => {
  const r = await 押す入口('https://example.com');
  if (r.headers['Access-Control-Allow-Origin']) throw new Error('★よその画面に 許しを出している★：' + r.headers['Access-Control-Allow-Origin']);
  if (r.headers['Access-Control-Allow-Origin'] === '*') throw new Error('★* のまま★');
});
await AT('★名乗りが無くても うちの画面は動く（same-origin は Origin を送らない）★', async () => {
  const r = await 押す入口('');
  if (r.status !== 200) throw new Error('名乗りなしを断っている：' + r.status);
});
await AT('★どの入口でも * は 二度と出さない★', async () => {
  for (const o of ['https://exally.vercel.app', 'https://example.com', '']) {
    const r = await 押す入口(o);
    if (r.headers['Access-Control-Allow-Origin'] === '*') throw new Error('★* を出している★：' + o);
  }
});
await AT('★許す入口に うちの画面が2つとも入っている（本番とテスト線）★', async () => {
  const a = handler.__許す入口 || [];
  for (const 要る of ['https://exally.vercel.app', 'https://exally-zeroact.github.io']) {
    if (a.indexOf(要る) < 0) throw new Error('★' + 要る + ' が 一覧に無い＝客の画面が止まる★');
  }
});

/* ── ★使った量を1行 残す（上限ではない）★ ── */
const 記録を拾う = async (fn) => {
  const 元 = console.log; const 行 = [];
  console.log = (...a) => { 行.push(a.join(' ')); };
  try { await fn(); } finally { console.log = 元; }
  return 行.filter((l) => l.indexOf('[ai] ') === 0).map((l) => JSON.parse(l.slice(5)));
};
await AT('★答えられた時に 使った量が1行 残る（入力/出力トークン）★', async () => {
  const 行 = await 記録を拾う(() => 押す入口('https://exally.vercel.app'));
  if (行.length !== 1) throw new Error('★記録が ' + 行.length + '行（1行でない）★');
  const r = 行[0];
  if (r.結果 !== 'ok') throw new Error('結果が ' + r.結果);
  if (r.入力トークン !== 12 || r.出力トークン !== 3) throw new Error('★使った量が 残っていない★：' + JSON.stringify(r));
  if (typeof r.かかった秒 !== 'number') throw new Error('かかった秒が無い');
});
await AT('★失敗した時も 何で失敗したかが1行 残る★', async () => {
  const 行 = await 記録を拾う(() => 呼ぶ(失敗({ status: 400, message: 'Your credit balance is too low' })));
  if (行.length !== 1) throw new Error('記録が ' + 行.length + '行');
  if (行[0].結果 !== 'zandaka') throw new Error('★何で失敗したか 残っていない★：' + JSON.stringify(行[0]));
});
await AT('★客が書いた文そのものは 残さない（長さだけ）★', async () => {
  const 行 = await 記録を拾う(() => 押す入口('https://exally.vercel.app'));
  const 字 = JSON.stringify(行[0]);
  if (字.indexOf('これ何？') >= 0) throw new Error('★客の中身を そのまま記録している★：' + 字);
  if (行[0].送った字数 !== 4) throw new Error('長さが違う：' + 行[0].送った字数);
});
await AT('★上限も 既定オフも 足していない（勝手に決めない）★', async () => {
  const fs2 = await import('node:fs');
  const src = fs2.readFileSync(path.join(ROOT, 'api/claude.js'), 'utf8').replace(/\/\*[\s\S]*?\*\//g, '');
  for (const 語 of ['RATE_LIMIT', 'MAX_PER_DAY', '1日', '回まで', 'MAX_MESSAGE_LEN']) {
    if (src.indexOf(語) >= 0) throw new Error('★上限らしき物を足している（数字は司さんと決める）★：' + 語);
  }
});

/* ── ★置いたまま使い回す（prompt caching）★ 2026-08-22 ──
 *  ★毎回まるごと送り直していた★のを直した。ここでは ★本当に送っている形★ を捕まえて数える。
 *  一次情報の決まり … 印は最大4か所／読み直し0.1倍・置く時1.25倍／Sonnet 4.6 は1,024トークン未満だと置かれない */
const 送った物を捕まえる = async (opt) => {
  const o = opt || {};
  let 送った = null;
  handler.__setClient({ messages: { create: async (p) => { 送った = p; return { content: [{ type: 'text', text: 'はい' }], usage: { input_tokens: 5, output_tokens: 3, cache_creation_input_tokens: o.置いた || 0, cache_read_input_tokens: o.読み直した || 0 } }; } } });
  const box = 受け皿();
  await handler({ method: 'POST', headers: {}, body: { message: o.文 || 'これ何？', history: o.履歴 || [], excelVersion: o.版 } }, box.res);
  return 送った;
};

await AT('★前置きは2つに分かれ、どちらにも「置いたまま」の印が付く★', async () => {
  const p = await 送った物を捕まえる({});
  if (!Array.isArray(p.system)) throw new Error('★前置きが 1本のまま（分けていない）★');
  if (p.system.length !== 2) throw new Error('前置きの数が ' + p.system.length);
  for (const b of p.system) {
    if (!b.cache_control || b.cache_control.type !== 'ephemeral') throw new Error('★印が付いていない★：' + JSON.stringify(b).slice(0, 80));
  }
});
await AT('★共通の所だけ 1時間もつ置き方／版ごとは 5分のまま★', async () => {
  const p = await 送った物を捕まえる({});
  const a = p.system[0].cache_control, b = p.system[1].cache_control;
  if (!a || a.ttl !== '1h') throw new Error('★共通が 1時間になっていない（単発で押す人が ずっと+25%）★：' + JSON.stringify(a));
  if (!b || b.ttl) throw new Error('★版ごとまで 1時間にしている（人ごとに変わる物に 2倍の置き賃を払う）★：' + JSON.stringify(b));
});
await AT('★長くもつ物を 先に置く（混ぜる時の決まり）★', async () => {
  const p = await 送った物を捕まえる({ 履歴: [{ role: 'user', content: 'a' }, { role: 'assistant', content: 'b' }] });
  const 印 = [];
  for (const b of p.system) if (b.cache_control) 印.push(b.cache_control.ttl || '5m');
  for (const m of p.messages) if (Array.isArray(m.content)) for (const b of m.content) if (b.cache_control) 印.push(b.cache_control.ttl || '5m');
  const 順 = 印.map((t) => (t === '1h' ? 1 : 0));
  for (let i = 1; i < 順.length; i++) {
    if (順[i] > 順[i - 1]) throw new Error('★短い物の後ろに 長い物が在る（APIに断られる）★：' + 印.join(','));
  }
  if (印[0] !== '1h') throw new Error('先頭が 1時間ではない：' + 印.join(','));
});
await AT('★1時間の分と 5分の分を 分けて記録する（置き賃が違う）★', async () => {
  const 元 = console.log; const 行 = [];
  console.log = (...a) => { 行.push(a.join(' ')); };
  try {
    handler.__setClient({ messages: { create: async () => ({ content: [{ type: 'text', text: 'はい' }], usage: { input_tokens: 21, output_tokens: 3, cache_creation_input_tokens: 1996, cache_read_input_tokens: 0, cache_creation: { ephemeral_1h_input_tokens: 1100, ephemeral_5m_input_tokens: 896 } } }) } });
    const box = 受け皿();
    await handler({ method: 'POST', headers: {}, body: { message: 'これ何？', history: [] } }, box.res);
  } finally { console.log = 元; }
  const r = JSON.parse(行.filter((l) => l.indexOf('[ai] ') === 0)[0].slice(5));
  if (r.置いた1時間 !== 1100 || r.置いた5分 !== 896) throw new Error('★分けて残していない（どちらの置き賃か 分からない）★：' + JSON.stringify(r));
});

await AT('★変わる物（版ごとの説明）は 共通の後ろに置く（前に置くと毎回 置き直し）★', async () => {
  const p365 = await 送った物を捕まえる({ 版: 'excel_365' });
  const p2016 = await 送った物を捕まえる({ 版: 'excel_2016' });
  if (p365.system[0].text !== p2016.system[0].text) throw new Error('★共通の所が 版で変わっている＝置いたまま使えない★');
  if (p365.system[1].text === p2016.system[1].text) throw new Error('版ごとの所が 変わっていない（分け方が違う）');
  if (p365.system[0].text.indexOf('Excel 365') >= 0) throw new Error('★共通の中に 版の名前が混ざっている★');
});
await AT('★前の会話の終わりに印を付ける／今 打った文には付けない★', async () => {
  const 履歴 = [
    { role: 'user', content: 'A1は？' }, { role: 'assistant', content: '合計だよ' },
    { role: 'user', content: 'B1は？' }, { role: 'assistant', content: '単価だよ' },
  ];
  const p = await 送った物を捕まえる({ 履歴, 文: 'C1は？' });
  const 最後の履歴 = p.messages[p.messages.length - 2];
  const 今の文 = p.messages[p.messages.length - 1];
  const 印 = (m) => Array.isArray(m.content) && m.content.some((b) => b.cache_control);
  if (!印(最後の履歴)) throw new Error('★前の会話に 印が付いていない（使い回せない）★');
  if (印(今の文)) throw new Error('★毎回変わる文に 印を付けている＝毎回 置き直しになる★');
  if (typeof 今の文.content !== 'string' || 今の文.content !== 'C1は？') throw new Error('打った文を 変えている');
});
await AT('★会話が無くても 落ちない（初めての1回目）★', async () => {
  const p = await 送った物を捕まえる({ 履歴: [] });
  if (p.messages.length !== 1) throw new Error('メッセージの数が ' + p.messages.length);
  if (Array.isArray(p.messages[0].content)) throw new Error('1回目に 印を付けている');
});
await AT('★印は 4か所を超えない（一次情報の上限）★', async () => {
  const 履歴 = [];
  for (let i = 0; i < 20; i++) 履歴.push({ role: i % 2 ? 'assistant' : 'user', content: 'x' + i });
  const p = await 送った物を捕まえる({ 履歴 });
  const n = p.system.filter((b) => b.cache_control).length
    + p.messages.filter((m) => Array.isArray(m.content) && m.content.some((b) => b.cache_control)).length;
  if (n > 4) throw new Error('★印が ' + n + 'か所＝APIに断られる★');
});
await AT('★置いた量・読み直した量を 記録に残す（後で値段を数えられる）★', async () => {
  const 元 = console.log; const 行 = [];
  console.log = (...a) => { 行.push(a.join(' ')); };
  try { await 送った物を捕まえる({ 置いた: 1200, 読み直した: 3400 }); } finally { console.log = 元; }
  const r = JSON.parse(行.filter((l) => l.indexOf('[ai] ') === 0)[0].slice(5));
  if (r.置いたトークン !== 1200 || r.読み直したトークン !== 3400) throw new Error('★使い回した量が 残っていない★：' + JSON.stringify(r));
});
await AT('★記録に 値段(円)を書き込んでいない（値段が変わったら嘘になる）★', async () => {
  const fs2 = await import('node:fs');
  const src = fs2.readFileSync(path.join(ROOT, 'api/claude.js'), 'utf8').replace(/\/\*[\s\S]*?\*\//g, '');
  for (const 語 of ['円', 'JPY', '0.1倍', '1.25倍']) {
    if (src.indexOf(語) >= 0) throw new Error('★値段を コードに書いている★：' + 語);
  }
});

await AT('★ハンドラを 本当に呼ぶ：失敗したら 200 を返さない★', async () => {
  const r = await 呼ぶ(失敗({ status: 401, message: 'invalid x-api-key' }));
  if (!r) throw new Error('返事が無い');
  if (r.status === 200) throw new Error('★失敗したのに 200 を返している（前の作り）★');
  if (r.status !== 401) throw new Error('番号が ' + r.status);
  if (r.body.error !== 'kagi') throw new Error('合言葉が ' + r.body.error);
});
await AT('★ハンドラを 本当に呼ぶ：中の言葉を 客に返さない★', async () => {
  const r = await 呼ぶ(失敗({ message: 'The ANTHROPIC_API_KEY environment variable is missing' }));
  const 字 = JSON.stringify(r.body);
  for (const 中の字 of ['Vercel', 'Environment', 'ANTHROPIC', 'エラーが発生しました', 'お試しください']) {
    if (字.indexOf(中の字) >= 0) throw new Error('★中の言葉「' + 中の字 + '」を 客に返している★：' + 字);
  }
  if (r.body.text !== '') throw new Error('★text に 何か入れている（画面が それを AIの答えとして出す）★：' + 字);
});
await AT('★ちゃんと答えられた時は 今までどおり 200 で text を返す（殺していない）★', async () => {
  handler.__setClient({ messages: { create: async () => ({ content: [{ type: 'text', text: 'これは合計だよ。' }] }) } });
  const box = 受け皿();
  await handler({ method: 'POST', body: { message: 'これ何？', history: [] }, headers: {} }, box.res);
  if (!box.出た || box.出た.status !== 200) throw new Error('200 を返していない：' + JSON.stringify(box.出た));
  if (box.出た.body.text !== 'これは合計だよ。') throw new Error('答えを 書き換えている：' + JSON.stringify(box.出た.body));
});
await AT('★中身なしの POST は 今までどおり 400（画面が「聞きたい事を書いてね」と言う）★', async () => {
  const box = 受け皿();
  await handler({ method: 'POST', body: {}, headers: {} }, box.res);
  if (!box.出た || box.出た.status !== 400) throw new Error('400 でない：' + JSON.stringify(box.出た));
});


/* ── ★自己確認：わざと壊して 赤になるか（★repo は読むだけ★）★ ──
 *  api/claude.js の中身を ★読み込んだ字の上で★ 壊し、vm の中で組み立て直して押す。
 *  ファイルには 1バイトも書かない。 */
if (process.argv.includes('--self-test')) {
  const vm = await import('node:vm');
  const fs = await import('node:fs');
  const 元 = fs.readFileSync(path.join(ROOT, 'api/claude.js'), 'utf8');
  const 偽AI = class { constructor() { this.messages = { create: async () => { throw Object.assign(new Error('invalid x-api-key'), { status: 401 }); } }; } };
  const 積む = (src) => {
    const m = { exports: {} };
    const req = (id) => (id === '@anthropic-ai/sdk' ? 偽AI : require_(path.resolve(ROOT, 'api', id)));
    vm.runInNewContext(src, { module: m, exports: m.exports, require: req, console: { error() {}, log: (...a) => console.log(...a) }, process, Date, JSON, Math, RegExp, Object, Array, String, Number, Boolean, Error });
    return m.exports;
  };
  const 押す = async (h, err) => {
    if (h.__setClient) h.__setClient({ messages: { create: async () => { throw err; } } });
    let 出た = null;
    const res = { setHeader() {}, status(c) { this._c = c; return this; }, json(b) { 出た = { status: this._c, body: b }; return this; }, end() { return this; } };
    await h({ method: 'POST', body: { message: 'これ何？', history: [] }, headers: {} }, res);
    return 出た;
  };
  const 残高err = Object.assign(new Error('Your credit balance is too low to access the Anthropic API.'), { status: 400 });
  const 鍵err = Object.assign(new Error('invalid x-api-key'), { status: 401 });
  const BREAKS = [
    ['★前の作りに戻す（何でも 200 ＋ 言い訳）★',
      (t) => t.replace('    const k = 失敗を分ける(err);', '    return res.status(200).json({ text: "申し訳ありません。エラーが発生しました。しばらくしてから再度お試しください。", tsv: "" });'),
      async (h) => (await 押す(h, 鍵err)).status !== 200],
    ['★残高切れを 400 のまま返す（画面が「送る中身が足りません」と嘘を言う）★',
      (t) => t.replace("  if (/credit balance/i.test(msg) || /insufficient[_ ]?quota/i.test(msg)) return { status: 402, 合言葉: 'zandaka' };", ''),
      async (h) => (await 押す(h, 残高err)).status === 402],
    ['★鍵の見分けを 外す★',
      (t) => t.replace("  if (st === 401 || st === 403 || /api[ _-]?key/i.test(msg)) return { status: 401, 合言葉: 'kagi' };", ''),
      async (h) => (await 押す(h, 鍵err)).body.error === 'kagi'],
    ['★誰の画面からでも叩ける（* に戻す）★',
      (t) => t.replace("  if (許す入口.indexOf(入口) >= 0) {", "  if (true) { 入口 = '*';").replace("const 入口 = (req.headers", "let 入口 = (req.headers"),
      async (h) => { const r = await 押す入口2(h, 'https://example.com'); return !r.headers['Access-Control-Allow-Origin']; }],
    ['★うちの画面(テスト線)を 一覧から落とす★',
      (t) => t.replace("  'https://exally-zeroact.github.io',", ''),
      async (h) => (h.__許す入口 || []).indexOf('https://exally-zeroact.github.io') >= 0],
    ['★使った量を 記録しない★',
      (t) => t.replace(/    記録\(\{[\s\S]{0,30}結果: 'ok',/, "    if (false) 記録({ 結果: 'ok',"),
      async (h) => { const 行 = await 記録2(h); return 行.length === 1 && 行[0].入力トークン === 12; }],
    ['★客が書いた文を そのまま記録に残す★',
      (t) => t.replace('      送った字数: o.字数 || 0,', '      送った字数: o.字数 || 0, 中身: o.中身,')
              .replace('      字数: message.length,', '      字数: message.length, 中身: message,'),
      async (h) => { const 行 = await 記録2(h); return JSON.stringify(行[0]).indexOf('これ何？') < 0; }],
    ['★前置きを1本のまま送る（置いたまま使い回せない）★',
      (t) => t.replace('      system: システム,', '      system: 部品.共通 + 部品.版ごと,'),
      async (h) => { const p = await 捕まえる2(h); return Array.isArray(p.system); }],
    ['★変わる物を 前に置く（毎回 置き直しになる）★',
      (t) => t.replace('text: 部品.共通,', 'text: __入替__,').replace('text: 部品.版ごと,', 'text: 部品.共通,').replace('text: __入替__,', 'text: 部品.版ごと,'),
      async (h) => { const p = await 捕まえる2(h); return p.system[0].text.indexOf('Excel') < 0; }],
    ['★今 打った文にも 印を付ける（毎回 置き直し）★',
      (t) => t.replace("      messages: [...会話, { role: 'user', content: message }],",
        "      messages: [...会話, { role: 'user', content: [{ type: 'text', text: message, cache_control: { type: 'ephemeral' } }] }],"),
      async (h) => { const p = await 捕まえる2(h, [{ role: 'user', content: 'a' }]); const m = p.messages[p.messages.length - 1]; return !(Array.isArray(m.content) && m.content.some((b) => b.cache_control)); }],
    ['★前の会話に 印を付けない（使い回さない）★',
      (t) => t.replace('    if (会話.length) {', '    if (false) {'),
      async (h) => { const p = await 捕まえる2(h, [{ role: 'user', content: 'a' }]); const m = p.messages[0]; return Array.isArray(m.content) && m.content.some((b) => b.cache_control); }],
    ['★共通まで 5分に戻す（単発の人が ずっと+25%）★',
      (t) => t.replace("{ type: 'ephemeral', ttl: '1h' }", "{ type: 'ephemeral' }"),
      async (h) => { const p = await 捕まえる2(h); return p.system[0].cache_control.ttl === '1h'; }],
    ['★版ごとまで 1時間にする（人ごとに変わる物に 2倍の置き賃）★',
      (t) => t.replace("      { type: 'text', text: 部品.版ごと, cache_control: { type: 'ephemeral' } },", "      { type: 'text', text: 部品.版ごと, cache_control: { type: 'ephemeral', ttl: '1h' } },"),
      async (h) => { const p = await 捕まえる2(h); return !p.system[1].cache_control.ttl; }],
    ['★使い回した量を 記録しない★',
      (t) => t.replace('      置いた: (response.usage && response.usage.cache_creation_input_tokens) || 0,', '      置いた: 0,'),
      async (h) => { const 行 = await 記録3(h); return 行.length === 1 && 行[0].置いたトークン === 1200; }],
    ['★中の言葉を text に入れて返す★',
      (t) => t.replace("    return res.status(k.status).json({ error: k.合言葉, text: '', tsv: '' });",
        "    return res.status(k.status).json({ error: k.合言葉, text: 'VercelのEnvironment VariablesにANTHROPIC_API_KEYを設定してください。', tsv: '' });"),
      async (h) => (await 押す(h, 鍵err)).body.text === ''],
  ];
  /* ★壊した版の記録は 検査の画面に出さない★（毎回そこで受け取って捨てる） */
  let 拾った = [];
  const 押す入口2 = async (h, origin) => {
    if (h.__setClient) h.__setClient({ messages: { create: async () => ({ content: [{ type: 'text', text: 'はい' }], usage: { input_tokens: 12, output_tokens: 3 } }) } });
    const 出た = { headers: {}, status: null, body: null };
    const res = { setHeader(k, v) { 出た.headers[k] = v; }, status(c) { 出た.status = c; return this; }, json(b) { 出た.body = b; return this; }, end() { return this; } };
    const 元 = console.log; 拾った = [];
    console.log = (...a) => { 拾った.push(a.join(' ')); };
    try {
      await h({ method: 'POST', headers: origin ? { origin } : {}, body: { message: 'これ何？', history: [] } }, res);
    } finally { console.log = 元; }
    return 出た;
  };
  const 捕まえる2 = async (h, 履歴) => {
    let 送った = null;
    if (h.__setClient) h.__setClient({ messages: { create: async (p) => { 送った = p; return { content: [{ type: 'text', text: 'はい' }], usage: { input_tokens: 5, output_tokens: 3, cache_creation_input_tokens: 1200, cache_read_input_tokens: 0 } }; } } });
    const res = { setHeader() {}, status() { return this; }, json() { return this; }, end() { return this; } };
    const 元 = console.log; console.log = () => {};
    try { await h({ method: 'POST', headers: {}, body: { message: 'これ何？', history: 履歴 || [] } }, res); } finally { console.log = 元; }
    return 送った;
  };
  const 記録3 = async (h) => {
    const 元 = console.log; const 行 = [];
    console.log = (...a) => { 行.push(a.join(' ')); };
    try { await 捕まえる2(h); } finally { console.log = 元; }
    return 行.filter((l) => l.indexOf('[ai] ') === 0).map((l) => JSON.parse(l.slice(5)));
  };
  const 記録2 = async (h) => {
    await 押す入口2(h, 'https://exally.vercel.app');
    return 拾った.filter((l) => l.indexOf('[ai] ') === 0).map((l) => JSON.parse(l.slice(5)));
  };

  console.log('');
  console.log('[self-test] わざと壊して 赤くなるかを数える（★ファイルには書かない★）');
  let red = 0;
  for (const [名, 壊す, 通れば緑] of BREAKS) {
    const bad = 壊す(元);
    if (bad === 元) { console.log('  ★置換できず★  ' + 名); continue; }
    let 赤 = false;
    try { 赤 = !(await 通れば緑(積む(bad))); } catch (e) { 赤 = true; }
    if (赤) { red++; console.log('  赤くなった  ' + 名); } else console.log('  ★素通り★  ' + 名);
  }
  console.log('');
  console.log('  ' + red + '/' + BREAKS.length + ' 通りで赤くなった');
  process.exit(red === BREAKS.length ? 0 : 1);
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
