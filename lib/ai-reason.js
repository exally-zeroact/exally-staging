/* ai-reason.js — ★AIに繋がらない時の 理由と 次の一手★（純関数）
 *
 *  ★なぜ在るか（2026-08-21 指示役が実配信を押して見つけた）★
 *    「🤖 AIに解説させる」を押したら ★「AIに接続できなかったよ。もう一度試してみてね。」★ だけが出た。
 *    ⇒ ★もう一度 試しても 何度でも失敗する★（テスト線には AIの窓口が そもそも無いため）。
 *    ★Rakually で直したのと同じ型★＝「出来ました」だけ言って 次に進めない。
 *
 *  ★お金を使わずに 測った事（2026-08-21）★
 *    テスト線（GitHub Pages・静的配信）
 *      GET  /api/claude                  → ★404★（HTML）
 *      GET  /exally-staging/api/claude   → ★404★（HTML）
 *      POST /api/claude 中身なし          → ★405★（HTML）
 *      ＝★鍵の問題ではない。窓口そのものが 存在しない★（静的配信は Node を動かせない）
 *    本番（Vercel）
 *      GET  /api/claude                  → ★405 {"error":"Method not allowed"}★
 *      POST /api/claude 中身なし          → ★400 {"error":"message is required"}★
 *      ＝★うちの api/claude.js が 本番で 動いている★
 *    ★どちらも Anthropic に届く前に返るので お金は1円も使っていない★。
 *
 *  ★2026-08-22：客に見せる字から ★ を外した★
 *    実配信で押して撮ったら 窓に「この画面には ★AIの窓口が…★」と ★が そのまま出ていた。
 *    ★★は うちの覚え書きの印であって 客の字ではない★（注記の中には 残してよい）。
 *
 *  ★本番で 1回だけ 実際に押した（2026-08-22 00:0x）★
 *    POST /api/claude {"message":"1+1は？数字だけ1行で答えて。"}
 *      → ★200 {"text":"2だよ。","tsv":""}★（2.84秒・叩いた回数 1回）
 *    ＝★鍵も 残高も 生きている。本番では 本当に つながる★
 *
 *  ★同時に見つけた穴（2026-08-22 に直した）★
 *    api/claude.js は ★失敗しても 200 を返し、言い訳を text に入れていた★。
 *    画面は 200 を「つながった」と読むので ★言い訳が AIの答えとして 吹き出しに出る★。
 *    鍵が無い時の字は「VercelのEnvironment Variablesに…」＝★中の言葉を客に見せていた★。
 *    ⇒ サーバが ★本当の番号＋合言葉★ を返し、客に見せる言葉は ★ここ1か所★ で作る。
 */
(function (root) {
  'use strict';

  /**
   * 返事から ★理由★ と ★次にどうすればよいか★ を作る
   * @param {Object} o { status:number|null, ネット切れ:boolean, 中身:string, 本番か:boolean }
   * @returns {{ok:boolean, 理由:string, 次:string, 言葉:string}}
   */
  function 読む(o) {
    o = o || {};
    if (o.ネット切れ) {
      return 作る('ネットに つながっていないようです。',
        'Wi-Fi や 電波を 確かめてから もう一度 押してね。');
    }
    /* ★合言葉が来ていたら 番号より そちらを先に読む★
       （サーバが「なぜ失敗したか」を そのまま返してくる。番号だけでは
         残高切れと「送る中身が足りない」が どちらも 400 で 区別できない） */
    var 合言葉 = o.理由 || '';
    if (合言葉 === 'zandaka') return 残高();
    if (合言葉 === 'kagi') return 鍵();
    if (合言葉 === 'komiai') return 混雑();
    if (合言葉 === 'jikangire') return 時間切れ();
    if (合言葉 === 'ai_shippai') return AI側();
    var s = o.status;
    if (s === 402) return 残高();
    if (s === 504) return 時間切れ();
    if (s === 404 || s === 405) {
      return 作る('この画面には AIの窓口が 置かれていません。',
        o.本番か ? '管理者に「/api/claude が配信されていない」と伝えてね。'
                 : 'テスト環境には AIが在りません（練習用の倉庫だけ）。AIは 本番の画面で使ってね。');
    }
    if (s === 401 || s === 403) return 鍵();
    if (s === 429) return 混雑();
    if (s === 400) {
      return 作る('送る中身が 足りませんでした。',
        '聞きたい事を 書いてから もう一度 押してね。');
    }
    if (s && s >= 500) return AI側();
    if (s && s >= 200 && s < 300) {
      if (!o.中身) {
        return 作る('AIから 空の返事が来ました。',
          'もう一度 押してみてね。何度も続くなら 管理者に伝えてね。');
      }
      return { ok: true, 理由: '', 次: '', 言葉: o.中身 };
    }
    return 作る('AIに つながりませんでした（' + (s === null || s === undefined ? '返事なし' : s) + '）。',
      'もう一度 押してみてね。何度も続くなら 管理者に伝えてね。');
  }

  /* ★言い方は ここ1か所★（番号でも 合言葉でも 同じ言葉になる） */
  function 鍵() {
    return 作る('AIの鍵が 受け付けられませんでした。',
      '押し直しても 直りません。管理者に「AIの鍵を見てほしい」と伝えてね。');
  }
  function 残高() {
    /* ★2026-08-18 アマかせで実際に起きた★＝残高が尽きると 本番のAIが止まる。
       前は「しばらくしてから もう一度」と言っていたので ★何度でも空振りした★。 */
    return 作る('AIの 使える分が 今 足りません。',
      '押し直しても 直りません。管理者に「AIの残高を足してほしい」と伝えてね。');
  }
  function 混雑() {
    return 作る('今 AIが 混み合っています。', '1分ほど 待ってから もう一度 押してね。');
  }
  function 時間切れ() {
    return 作る('AIの返事が 時間切れになりました。',
      'もう一度 押してみてね。何度も続くなら 管理者に伝えてね。');
  }
  function AI側() {
    return 作る('AI側で 失敗しました。',
      'もう一度 押してみてね。何度も続くなら 管理者に伝えてね。');
  }

  function 作る(理由, 次) {
    return { ok: false, 理由: 理由, 次: 次, 言葉: 理由 + '\n' + 次 };
  }

  /**
   * ★中身が無い物に AIを呼ばない（お金を使わない）★
   * @param {Array<string>} 字たち 選んだ所に入っている字
   * @returns {{呼ぶ:boolean, 言葉:string}}
   */
  function 呼んでよいか(字たち) {
    var 何か在る = (字たち || []).some(function (s) {
      return s !== undefined && s !== null && String(s).trim() !== '';
    });
    if (何か在る) return { 呼ぶ: true, 言葉: '' };
    return {
      呼ぶ: false,
      言葉: 'ここには まだ 何も入っていません。\n中身が在る所を 選んでから 押してね。',
    };
  }

  var api = { 読む: 読む, 呼んでよいか: 呼んでよいか };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.AiReason = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
