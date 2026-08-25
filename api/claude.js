const Anthropic = require('@anthropic-ai/sdk');
// ★法定の数値は kyuyo/lib/ の本体を読む（写しを作らない）。
//   2026-08-02: リポジトリ直下に写しを置いていて、掃除でそれを消した時にここが
//   MODULE_NOT_FOUND になり /api/claude が毎回500（＝チャットが全部落ちた）。
//   本体を直接読めば、消しても場所が変わっても同じ所を指す。
//   参照が生きているかは tests/refs-resolve.test.mjs がCIで見張っている。
const SHAKAIHOKEN_HYO = require('../kyuyo/lib/shakaihoken-hyo.js');
const KOYO_HOKEN      = require('../kyuyo/lib/koyo-hoken.js');
const SHOUHIZEI_RITSU = require('../kyuyo/lib/shouhizei-ritsu.js');

/* ★let にしてある理由★＝下の __setClient（テスト用の窓）から 偽のAIに差し替えて、
   ★失敗した時に本当に何を返すか★を機械で押すため。本番では 1ミリも挙動が変わらない。 */
let client = new Anthropic({
  apiKey: process.env.ANTHROPIC_API_KEY,
});

// ===== Excelバージョン情報マップ =====
const VERSION_MAP = {
  'excel_365':    { name: 'Excel 365',        group: 'latest' },
  'excel_2024':   { name: 'Excel 2024',       group: 'latest' },
  'excel_2021':   { name: 'Excel 2021',       group: 'newer'  },
  'excel_2019':   { name: 'Excel 2019',       group: 'older'  },
  'excel_2016':   { name: 'Excel 2016',       group: 'older'  },
  'excel_mac':    { name: 'Excel for Mac',    group: 'older'  },
  'excel_online': { name: 'Excel Online',     group: 'online' },
  'excel_none':   { name: 'Excel持ってない',   group: 'exally_only' }
};

// ===== Exally未対応関数リスト =====
const EXALLY_UNSUPPORTED = {
  // 完全未対応（構造上Exally内で動作不可）
  full: [
    'RTD',
    'CUBEMEMBER', 'CUBEVALUE', 'CUBESET', 'CUBESETCOUNT',
    'CUBEMEMBERPROPERTY', 'CUBERANKEDMEMBER', 'CUBEKPIMEMBER',
    'WEBSERVICE', 'FILTERXML', 'GETPIVOTDATA'
  ],
  // まだ未対応（将来実装予定）
  pending: [
    'GROUPBY', 'PIVOTBY',
    'VSTACK', 'HSTACK', 'TOCOL', 'TOROW',
    'CHOOSEROWS', 'CHOOSECOLS',
    'TEXTSPLIT', 'TEXTBEFORE', 'TEXTAFTER'
  ]
};

// ===== バージョン情報取得（デフォルトexcel_365） =====
function getVersionInfo(versionKey) {
  return VERSION_MAP[versionKey] || VERSION_MAP['excel_365'];
}

// ===== 共通ベースプロンプト =====
const SYSTEM_PROMPT_BASE = `あなたはExally（エクサリー）というExcel完全代替SaaSのAI助手です。
Exallyは日本の中小企業・個人事業主向けに作られた「Excelそのものになる」サービスで、
Excel関数500件超に対応した自作グリッド（ブック）・チャット型AI・テンプレート自動生成を備えています。
ユーザーはExally内だけで書類を完成させられます（Excelに戻らなくていい）。
ExcelがPCに無いユーザーでも全機能が使えます。
Excelバージョン別（365/2024/2021/2019/2016/Mac/Online/なし）に合わせた最適な回答をします。

【回答ルール・絶対厳守】
- 常に日本語で回答する
- 回答は簡潔に・冗長な説明は禁止
- 前置き・長い解説・余計な注意書きは禁止
- ##・###などのMarkdown見出しは絶対に使わない
- |（縦棒）を使ったMarkdownテーブルは絶対に使わない
- 箇条書きの多用禁止
- 関数・数式は必ずExcelで動く形式（=で始まる）で提示する
- 数式を本文中に書く場合は必ずバッククォート（\`）で囲む
- セル結合の提案は禁止
- 語尾は「〜だよ」「〜してみて」調で短く
- 結論を最初に書いてから必要なら一言補足する
- 表やテンプレートを作る場合は必ずTSV形式でも出力する

【数式の返答フォーマット・絶対厳守】
数式を説明するときは必ず以下の順番で返答する：

①1行目：1文で結論（何ができるかだけ）

②各引数を絵文字で説明（必ず以下の順番・色を守る）
🔵 A2 → （第1引数の説明）
🟠 B1:D20 → （第2引数の説明）
🟣 2 → （第3引数の説明）
🟢 FALSE → （第4引数の説明）

③最後に数式をコードブロック（バッククォート3つ）で提示
例：
\`\`\`
=VLOOKUP(A2, Sheet2!$A:$C, 2, FALSE)
\`\`\`

- 日本語プレースホルダー（検索値・範囲・列番号など）は絶対に使わない
- 必ず実際のセル参照（A2・B1:D20・$A:$C等）を使う
- SUMIFは🔵範囲 🟠条件 🟣合計範囲 の3色
- IFは🔵条件 🟠真の値 🟣偽の値 の3色
- IFERRORは🔵数式 🟠エラー時の値 の2色

【TSV出力ルール・絶対厳守】
- 表・テンプレートを作る場合は、本文の説明の後に必ず以下の形式で出力する：

--- TSV_START ---
（ここにタブ区切りのデータ）
--- TSV_END ---

- TSVのセル区切りは必ずタブ文字（\\t）を使う・縦棒（|）は絶対に使わない
- TSVの数式セルはExcelで動く形式（=SUM(B2:B10) など）で記載する
- セル幅・書式は貼り付け後に手動調整が必要な旨を末尾に添える`;

// ===== 税務・給与の基準数値（kyuyo/lib の本体から・年度は対象月で自己選択） =====
//  ・健保/介護は社保年度（3月起算）、雇用保険は労働保険年度（4月起算）で切り替わる。
//  ・呼び出しの【たびに】組み立てる＝年度をまたいでも古い値を返さない
//    （関数が温まったまま年度が変わる、を避ける）。
function buildStatutoryPrompt(ymArg) {
  const ym = ymArg || new Date().toISOString().slice(0, 7);  // 'YYYY-MM'
  const kenko = SHAKAIHOKEN_HYO.getKenko('tokyo', ym);       // {jugyoin, nendo}
  const koyoYear = KOYO_HOKEN.employYearOfYm(ym);
  const koyoRate = KOYO_HOKEN.employRate('ippan', koyoYear);
  const koyoNendo = '令和' + (koyoYear - 2018) + '年度';
  return `

【税務・給与計算の基準数値】
- 健康保険料率（東京）: ${(kenko.jugyoin * 100).toFixed(3)}%（従業員負担・労使折半・${kenko.nendo}）
- 厚生年金保険料率: ${(SHAKAIHOKEN_HYO.KOSEI_NENKIN_RITSU_JUGYOIN * 100).toFixed(2)}%（従業員負担・労使折半・全国一律）
- 雇用保険料率: ${(koyoRate * 100).toFixed(2)}%（従業員負担・一般の事業・${koyoNendo}）
- 消費税: ${(SHOUHIZEI_RITSU.hyojun * 100).toFixed(0)}%（標準）/ ${(SHOUHIZEI_RITSU.keigen * 100).toFixed(0)}%（軽減）`;
}

// ===== 動的プロンプト生成 =====
function buildDynamicPrompt(versionInfo) {
  const group = versionInfo.group;
  const name = versionInfo.name;

  // グループ別の回答ルール
  let groupRule = '';

  if (group === 'latest') {
    groupRule = `
【ユーザーのExcel環境】
使用中のExcel：${name}（最新バージョン）

【回答ルール】
- XLOOKUP・スピル関数（FILTER/SORT/UNIQUE等）・LET・LAMBDA等の最新関数を積極的に提案する
- 「Exally内でも同じ数式が動くよ」程度の短い補足でOK（詳細対比は不要）
- 基本関数（SUM/AVERAGE等）では補足不要`;
  }
  else if (group === 'newer') {
    groupRule = `
【ユーザーのExcel環境】
使用中のExcel：${name}

【回答ルール】
- XLOOKUP等の新関数は使えるので積極提案
- 動的配列関数（FILTER/SORT等）は一部対応・使用時は注記を添える
- 「Exally内ならもっと便利な方法もあるよ」程度の補足を時々添える`;
  }
  else if (group === 'older') {
    groupRule = `
【ユーザーのExcel環境】
使用中のExcel：${name}（旧バージョン）

【回答ルール】
- メイン回答は古いExcelで動く関数（VLOOKUP・ネストIF・配列数式・CONCATENATE等）を使う
- XLOOKUP・FILTER・SORT・UNIQUE・LET・LAMBDA等は使わない
- 必ず「💡 Exally内なら〜」を併記する（以下の対比がある場合）：
  - VLOOKUP → XLOOKUP
  - ネストIF → IFS/SWITCH
  - 配列数式（Ctrl+Shift+Enter） → FILTER/SORT/UNIQUE
  - CONCATENATE → TEXTJOIN/CONCAT
- 併記フォーマット：
  💡 Exally内なら 〇〇 がもっと便利
  🔵🟠🟣🟢 で引数を色分けして説明
  \`\`\`代替数式\`\`\`
  ※違いを3点以内で示す
- 基本関数（SUM/AVERAGE/COUNT等）は併記不要`;
  }
  else if (group === 'online') {
    groupRule = `
【ユーザーのExcel環境】
使用中のExcel：${name}（機能制限あり）

【回答ルール】
- Excel Onlineは一部機能に制限があるので基本関数を中心に提案
- 複雑な機能は「Excelデスクトップでお試し」と補足
- 「💡 Exally内ならもっと便利に使えるよ」を時々併記`;
  }
  else if (group === 'exally_only') {
    groupRule = `
【ユーザーのExcel環境】
使用中のExcel：${name}（Exally内完結）

【回答ルール】
- XLOOKUP・スピル関数・LET・LAMBDA等の最新関数を積極的に使用
- 「AIに話しかけてセルに書き込み」機能を時々案内
- Excelへの配慮は不要・Exallyの全機能を活かした回答をする`;
  }

  // 全グループ共通ルール
  const commonRule = `

【全グループ共通ルール】
1. ユーザーが「動かない」「エラーが出る」「#NAME?」「#VALUE!」「#REF!」「古いExcel」「使えない」「対応してない」と反応してきたら、古いExcelで動く数式に切り替えて回答する
2. 切り替え時は一言添える：「もしかしてExcel 2019以前？古い版でも動く書き方を提案するね」
3. 適切なタイミングで「設定画面でExcelバージョンを変更できるよ」を案内する

【Exally未対応関数の扱い】

完全未対応（Exally内で構造上動かない）：
${EXALLY_UNSUPPORTED.full.join(', ')}

まだ未対応（将来実装予定）：
${EXALLY_UNSUPPORTED.pending.join(', ')}

これらの関数について質問されたら：
1. Excelでの使い方を通常通り🔵🟠🟣🟢の色分けで説明する
2. ⚠️ を付けて「Exally内ではこの関数は動かないよ」と明記する
3. 💡 で代替手段を色分けで提案する（SUMIFS/FILTER/UNIQUE/INDEX等の対応関数で同じ結果を出す方法）
4. 将来対応予定なら「※Exally内で対応予定」を添える（具体的な時期は書かない）
`;

  return SYSTEM_PROMPT_BASE + buildStatutoryPrompt() + groupRule + commonRule;
}

/* ★前置きを「置いたまま使い回す」ために 2つに分ける（2026-08-22）★
   ・前半＝★どの版の人にも同じ★（作りの説明＋法定の基準数値）＝ここを置いたままにする
   ・後半＝★版ごとに変わる★（Excel 365 / 2016 / 持っていない …）＝置き場所の後ろに回す
   ★変わる物を前に置くと 毎回 置き直しになって 逆に高くなる★（一次情報の決まり） */
function buildPromptParts(versionInfo) {
  const 共通 = SYSTEM_PROMPT_BASE + buildStatutoryPrompt();
  const 全部 = buildDynamicPrompt(versionInfo);
  const 版ごと = 全部.slice(共通.length);
  return { 共通, 版ごと };
}

/* ★うちの画面から来た物だけ受ける（2026-08-22 指示役）★
   前は Access-Control-Allow-Origin: '*' ＝★誰の画面からでも叩けた★。
   ★正直に：これだけでは止まりません★＝道具(curl等)で直接叩く相手には効かない（名乗りは詐称できる）。
   連打を止めるのは Vercel の入口（指示役の担当）。ここは「よその画面から使われる」のを断るだけ。
   ★うちの画面は 同じ入れ物(same-origin)なので、名乗りが無くても 今までどおり動く★ */
/* ★1回に送れる大きさ（司さん承認 2026-08-22）★
   ★2026-08-25：数字は 下の 事故止め に集めた★（20,000 が2か所に在ると 片方が古くなる） */

/* ══ ★4 事故止め（2026-08-25 司さんの数字・指示役の指示）★ ══════════════════
   ★なぜ在るか（私が 偽のAIで 0円で押して 実測した穴）★
     message 20,000字 ＋ history 40件×50,000字 → ★202万字が そのまま AIへ渡り 200 が返った★
     ＝ ★1回で 会社が傾く額を 出せる作り★だった（1回の上限は 字数しか見ていなかった）。
   ★数字は ここ1か所だけ★（散らすと 直す時に 必ず 片方が残る）
   ★止めた時も 必ず 記録を残す★（止まった事が 見えないと 誰も気づけない） */
const 事故止め = {
  分の回数: 10,          // ★1分に10回（人ごと）★
  分の窓ミリ秒: 60 * 1000,
  日の回数: 100,         // ★1日に100回（人ごと）★
  日の窓ミリ秒: 24 * 60 * 60 * 1000,
  会話の合計字数: 40000, // ★history 合計40,000字＝古い方から捨てる★
  渡せるトークン: 20000, // ★1回にAIへ渡すのは2万トークンまで（2026-08-09 司さんの決定）★
  /* ★1回に送れる字数（司さん承認 2026-08-22）★＝たまたま同じ数だが 別の物なので 別々に持つ */
  一度に送れる字数: 20000,
};
const 一度に送れる字数 = 事故止め.一度に送れる字数;

/* ★数え場＝この機械の中だけ★
   ★正直に書く★＝Vercel は 機械が増える。増えた分は 別勘定になるので ★すり抜けが在る★。
   ★共有の数え場（Supabaseに表を1つ／Vercel KV）は 倉庫を触るので 指示待ち★。
   それでも ★1台に集中する連打（実際の事故の形）は ここで止まる★。 */
const 数え場 = new Map();
const __数え場を空にする = () => 数え場.clear();

/** ★誰の分か★＝ログインの人ID（あれば）と 入口のIP（必ず）の2本で数える。
 *  ★人IDだけだと 名乗りを書き換えれば すり抜ける★ので IP も必ず数える。 */
function 誰か(req) {
  const h = (req && req.headers) || {};
  const 生 = String(h.authorization || h.Authorization || '');
  let 人 = '';
  const m = 生.match(/^Bearer\s+([\w-]+)\.([\w-]+)\./);
  if (m) {
    try {
      const 中 = JSON.parse(Buffer.from(m[2].replace(/-/g, '+').replace(/_/g, '/'), 'base64').toString('utf8'));
      if (中 && 中.sub) 人 = 'u:' + String(中.sub);
    } catch (e) { /* 読めない名乗りは 無い物として扱う（IPで数える） */ }
  }
  const ip = String(h['x-forwarded-for'] || h['x-real-ip'] || (req.socket && req.socket.remoteAddress) || '')
    .split(',')[0].trim();
  const 鍵 = [];
  if (人) 鍵.push(人);
  鍵.push('ip:' + (ip || 'unknown'));
  return 鍵;
}

/** ★押した回数を見る（数えるのは 通した分だけ）★
 *  止めた分まで数えると ★窓がいつまでも空かない★＝待っても直らなくなる。
 *  @returns {{止める:boolean, あと秒:number, どれ:string}}
 */
function 押した回数を見る(鍵たち, いま) {
  for (const k of 鍵たち) {
    const 山 = 数え場.get(k) || [];
    const 分内 = 山.filter((t) => いま - t < 事故止め.分の窓ミリ秒);
    if (分内.length >= 事故止め.分の回数) {
      const あと = 事故止め.分の窓ミリ秒 - (いま - 分内[0]);
      return { 止める: true, あと秒: Math.max(1, Math.ceil(あと / 1000)), どれ: '分' };
    }
    const 日内 = 山.filter((t) => いま - t < 事故止め.日の窓ミリ秒);
    if (日内.length >= 事故止め.日の回数) {
      const あと = 事故止め.日の窓ミリ秒 - (いま - 日内[0]);
      return { 止める: true, あと秒: Math.max(1, Math.ceil(あと / 1000)), どれ: '日' };
    }
  }
  return { 止める: false, あと秒: 0, どれ: '' };
}

/** ★通した分を 1回 数える（★捨てるのは 1日より古い物だけ★＝溜め続けない）★ */
function 数えておく(鍵たち, いま) {
  for (const k of 鍵たち) {
    const 山 = (数え場.get(k) || []).filter((t) => いま - t < 事故止め.日の窓ミリ秒);
    山.push(いま);
    数え場.set(k, 山);
  }
}

/** ★トークンの見積もり（多めに見る＝安全側）★
 *  ・日本語などは 1文字＝1トークンとして数える
 *  ・英数字・記号は 4文字＝1トークン（一次情報の目安）
 *  ★これは 見積もり★＝本物の数はAIが返してから分かる。だから ★多めに見る★。 */
function 見積もりトークン(s) {
  s = String(s == null ? '' : s);
  let 和 = 0, 英 = 0;
  for (const ch of s) {
    if (ch.charCodeAt(0) < 128) 英++; else 和++;
  }
  return 和 + Math.ceil(英 / 4);
}

/** ★会話を 合計40,000字までに削る（古い方から捨てる）★
 *  ★最後の1往復（最後の2件）は 必ず残す★＝直前の話が消えると 会話にならない。 */
function 会話を字数で削る(会話, 上限) {
  const out = 会話.slice();
  const 合計 = () => out.reduce((a, m) => a + String(m.content || '').length, 0);
  while (out.length > 2 && 合計() > 上限) out.shift();
  return out;
}

/** ★渡す物 全部（前置き＋会話＋今の文）を 2万トークン以内に削る★
 *  ★決まりの強さ★＝お金の上限（2万トークン）が 先。
 *    最後の1往復も 残せない時だけ ★会話を全部 捨てる★（そして 記録に残す）。 */
function 会話をトークンで削る(会話, 前置き字, 今の文, 上限) {
  const 土台 = 見積もりトークン(前置き字) + 見積もりトークン(今の文);
  const out = 会話.slice();
  const 合計 = () => 土台 + out.reduce((a, m) => a + 見積もりトークン(m.content || ''), 0);
  while (out.length > 2 && 合計() > 上限) out.shift();
  if (合計() > 上限) return { 会話: [], 全部捨てた: true };
  return { 会話: out, 全部捨てた: false };
}

const 許す入口 = [
  'https://exally.vercel.app',
  'https://exally-zeroact.github.io',
  'http://localhost:8080',
  'http://127.0.0.1:8080',
];

/* ★使った量を1行 残す（★上限ではない。「見えない」を「見える」にするだけ★）★
   これが無いと ★何が起きても 原因も 止め方も 分からない★（2026-08 に残高が尽きて 本番が止まった）。
   ★人が書いた文そのものは残さない★（客の中身なので 長さだけ）。 */
function 記録(o) {
  try {
    console.log('[ai] ' + JSON.stringify({
      結果: o.結果,
      入力トークン: o.入力 || 0,
      出力トークン: o.出力 || 0,
      /* ★置いたまま使い回した量★（置いた＝1.25倍／読み直した＝0.1倍・一次情報）
         ★値段はここに書かない★＝値段が変わったら嘘になる。数だけ残して 外で計算する。 */
      置いたトークン: o.置いた || 0,
      置いた1時間: o.置いた1時間 || 0,
      置いた5分: o.置いた5分 || 0,
      /* ★うちの段位と残量（Anthropic が返事に付けてくる物・お金は増えない）★ */
      段位: o.段位,
      読み直したトークン: o.読み直した || 0,
      送った字数: o.字数 || 0,
      会話の数: o.会話 || 0,
      /* ★4 事故止め：何を捨てたか／なぜ止めたか（黙って小さくしない・黙って止めない）★ */
      会話を削った: o.会話を削った || 0,
      会話の字数: o.会話の字数 || 0,
      会話の元の字数: o.会話の元の字数 || 0,
      会話を全部捨てた: o.会話を全部捨てた || false,
      渡した見積もりトークン: o.渡した見積もりトークン || 0,
      待ち秒: o.待ち秒 || 0,
      どれ: o.どれ || '',
      かかった秒: o.秒,
      入口: o.入口 || '(名乗りなし)',
    }));
  } catch (e) { /* 記録で本体を落とさない */ }
}

module.exports = async (req, res) => {
  // CORS … ★うちの画面から来た物だけ★
  const 入口 = (req.headers && (req.headers.origin || req.headers.Origin)) || '';
  if (許す入口.indexOf(入口) >= 0) {
    res.setHeader('Access-Control-Allow-Origin', 入口);
    res.setHeader('Vary', 'Origin');
  }
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const { message, history, excelVersion } = req.body || {};

    if (!message || typeof message !== 'string') {
      return res.status(400).json({ error: 'message is required', text: '', tsv: '' });
    }
    /* ★1回に送れる大きさ＝20,000文字（司さん承認 2026-08-22）★
       ★これは 司さんが決めた数字★＝私が思い付きで置いた物ではない。
       ★回数の上限は まだ入れない★（1分◯回・1時間◯回は 数字が決まっていない）。 */
    if (message.length > 一度に送れる字数) {
      記録({ 結果: 'ookisugi', 字数: message.length, 入口: 入口 });
      return res.status(413).json({ error: 'ookisugi', text: '', tsv: '' });
    }

    /* ★4 事故止め＝回数（2026-08-25 司さんの数字）★
       ★数えるのは 通した分だけ★／★止めた時も 必ず 記録に残す★
       ★「混み合っています」とは言わない★＝待てば直ると分かる言い方にする（合言葉 tsukaisugi） */
    const 誰 = 誰か(req);
    const いま = Date.now();
    const 見張り = 押した回数を見る(誰, いま);
    if (見張り.止める) {
      記録({ 結果: 'tsukaisugi', 字数: message.length, 待ち秒: 見張り.あと秒, どれ: 見張り.どれ, 入口: 入口 });
      res.setHeader('Retry-After', String(見張り.あと秒));
      return res.status(429).json({
        error: 'tsukaisugi', どれ: 見張り.どれ, 待ち秒: 見張り.あと秒, text: '', tsv: '',
      });
    }
    数えておく(誰, いま);

    // 会話履歴のサニタイズ（不正エントリ除去・最大40メッセージ=20ターンに制限）
    const 生の会話 = (Array.isArray(history) ? history : [])
      .filter(m =>
        m &&
        (m.role === 'user' || m.role === 'assistant') &&
        typeof m.content === 'string' &&
        m.content.trim().length > 0
      )
      .slice(-40);
    /* ★4 事故止め＝大きさ（2026-08-25）★
       ★件数だけ見ていたのが 穴だった★＝40件でも 1件が50,000字なら 200万字が そのまま渡る。
       ①★合計40,000字まで（古い方から捨てる・最後の1往復は必ず残す）★
       ②★前置きも足して 2万トークンまで★（お金の上限が 先） */
    const 字で削った = 会話を字数で削る(生の会話, 事故止め.会話の合計字数);
    const 削る前の字数 = 生の会話.reduce((a, m) => a + m.content.length, 0);

    // バージョンに応じた動的プロンプトを構築
    const versionInfo = getVersionInfo(excelVersion);
    const dynamicPrompt = buildDynamicPrompt(versionInfo);

    /* ★置いたまま使い回す（prompt caching）★ 2026-08-22
       なぜ … 会話40件(20往復)を ★毎回まるごと送り直していた★＝1回 平均 約4円。
              個人1,280円/月500回 なら 原価2,000円＝★1人目から赤字★（指示役の実測）。
       やり方 … ①前置きの共通部分 ②版ごとの部分 ③★前の会話まで★ の3か所に印を付ける。
              ★印は最大4か所★／★読み直し 0.1倍・置く時 1.25倍（5分もつ）★＝一次情報。
              ★Sonnet 4.6 は 1,024トークン未満だと 黙って置かれない★ので 前置きは1つに束ねる。
       ★客の画面は 1文字も変えていない（我慢も 上限も していない）★ */
    const 部品 = buildPromptParts(versionInfo);
    /* ★①の共通の所だけ 1時間もつ置き方にする（2026-08-22 指示役の 1-b）★
       なぜ … 5分の置き方は ★2回目から★ 安くなる仕掛け。
              ★5分 空けて 1回だけ押して終わる人★は 毎回が「1回目」＝★ずっと +25%★。
              今の Exally は 押す間隔が5分より長いので ★このままでは 損★。
       ①は ★誰が押しても 同じ字★なので、1時間 置いておけば 別の人の1回目でも 読み直せる。
       ★置く時 2倍／読み直し 0.1倍＝元が取れるのは 3回 読み直してから★（一次情報）
       ★長くもつ物を 先に置く★（混ぜる時の決まり）＝①が先・②と会話は 5分のまま */
    const システム = [
      { type: 'text', text: 部品.共通, cache_control: { type: 'ephemeral', ttl: '1h' } },
      /* ★②も1時間（2026-08-22 指示役の裁定）★
         ②は ★版ごと＝最大8通りしか無い＝みんなで共有する字★（人ごとには変わらない）。
         5分だと ★間隔が空く人には 置き賃(1.25倍)が 毎回かかる★＝実測で N=3 でも まだ+4%。
         ★1時間にすると N=10 で −68%★（★③前の会話だけ 5分のまま＝人ごと・会話ごとに変わるから★） */
      { type: 'text', text: 部品.版ごと, cache_control: { type: 'ephemeral', ttl: '1h' } },
    ];
    /* ★前の会話は「変わらない所の終わり」に印を付ける★
       （今回 打った文は 毎回変わるので ★印を付けない★＝付けると毎回 置き直しになる） */
    const 前置き字 = 部品.共通 + 部品.版ごと;
    const トークンで削った = 会話をトークンで削る(字で削った, 前置き字, message, 事故止め.渡せるトークン);
    const sanitizedHistory = トークンで削った.会話;
    const 会話 = sanitizedHistory.map((m) => ({ role: m.role, content: m.content }));
    if (会話.length) {
      const 最後 = 会話[会話.length - 1];
      会話[会話.length - 1] = {
        role: 最後.role,
        content: [{ type: 'text', text: 最後.content, cache_control: { type: 'ephemeral' } }],
      };
    }

    const 始めた = Date.now();
    /* ★返事の見出し(ヘッダ)から うちの段位と残量を読む（2026-08-25 指示役）★
       Anthropic は 返事のたびに anthropic-ratelimit-*-limit / -remaining / -reset を返している。
       ★今まで 捨てていた★＝★次に1回 AIを呼ぶだけで 段位と残量が分かる★（人に聞かなくて済む）。
       ★お金は1円も余分にかからない（同じ1回の返事に付いてくる物）★
       ★withResponse が無い版でも 落ちない★（その時は 段位は「未測定」になるだけ） */
    let 見出し = null;
    const 送り = client.messages.create({
      model: 'claude-sonnet-4-6',
      max_tokens: 2000,
      system: システム,
      messages: [...会話, { role: 'user', content: message }],
    });
    let response;
    if (送り && typeof 送り.withResponse === 'function') {
      const 包み = await 送り.withResponse();
      response = 包み.data;
      見出し = 包み.response && 包み.response.headers;
    } else {
      response = await 送り;
    }

    記録({
      結果: 'ok',
      入力: (response.usage && response.usage.input_tokens) || 0,
      出力: (response.usage && response.usage.output_tokens) || 0,
      置いた: (response.usage && response.usage.cache_creation_input_tokens) || 0,
      読み直した: (response.usage && response.usage.cache_read_input_tokens) || 0,
      /* ★1時間の分と 5分の分は 置き賃が違う（2倍 と 1.25倍）ので 分けて残す★ */
      段位: 見出しを読む(見出し),
      置いた1時間: (response.usage && response.usage.cache_creation
        && response.usage.cache_creation.ephemeral_1h_input_tokens) || 0,
      置いた5分: (response.usage && response.usage.cache_creation
        && response.usage.cache_creation.ephemeral_5m_input_tokens) || 0,
      字数: message.length,
      会話: sanitizedHistory.length,
      /* ★何を捨てたかを 必ず残す（黙って小さくしない）★ */
      会話を削った: 生の会話.length - sanitizedHistory.length,
      会話の字数: sanitizedHistory.reduce((a, m) => a + m.content.length, 0),
      会話の元の字数: 削る前の字数,
      会話を全部捨てた: トークンで削った.全部捨てた || false,
      渡した見積もりトークン: 見積もりトークン(前置き字) + 見積もりトークン(message)
        + sanitizedHistory.reduce((a, m) => a + 見積もりトークン(m.content), 0),
      秒: Math.round((Date.now() - 始めた) / 100) / 10,
      入口: 入口,
    });

    const fullText = response.content
      .filter(block => block.type === 'text')
      .map(block => block.text)
      .join('');

    const tsvMatch = fullText.match(/---\s*TSV_START\s*---\n([\s\S]+?)\n---\s*TSV_END\s*---/);
    const tsv = tsvMatch ? tsvMatch[1].trim() : '';
    const text = tsvMatch ? fullText.replace(tsvMatch[0], '').trim() : fullText;

    return res.status(200).json({ text, tsv });

  } catch (err) {
    console.error('Claude API error:', err);
    /* ★2026-08-21 直した：失敗しても 200 で「答えのふり」をしていた★
       ・鍵が無い時 …「APIキーが設定されていません。VercelのEnvironment Variables…」
         ＝★中の言葉を そのまま客に見せていた★（客は Vercel を知らない）
       ・残高が尽きた時・混み合い …「しばらくしてから再度お試しください」だけ
         ＝★押し直しても直らない物に「もう一度」と言う＝何度でも空振りする★
         （アマかせ 2026-08-18 の事故と同じ型）
       ・画面(lib/ai-reason.js)は 200 を「つながった」と読むので、
         ★この言い訳が AIの答えとして 吹き出しに出ていた★。
       ⇒★理由は 番号と合言葉で返す。客に見せる言葉は 画面が1か所で作る★ */
    const k = 失敗を分ける(err);
    記録({
      結果: k.合言葉,
      入力: 0,
      出力: 0,
      字数: (req.body && typeof req.body.message === 'string') ? req.body.message.length : 0,
      会話: 0,
      入口: 入口,
    });
    return res.status(k.status).json({ error: k.合言葉, text: '', tsv: '' });
  }
};

/* ★失敗を 客に伝わる形に分ける（純関数・テストが直接 叩く）★
   Anthropic の返し方:
     残高切れ … status 400 ＋ message に "credit balance is too low"
                （2026-08-18 アマかせで実際に出た字）
     鍵 …       status 401 / 403
     混み合い … status 429
   ★400 を そのまま返すと 画面は「送る中身が 足りません」と言う＝嘘になる★ので
   ★残高切れは 402（お金が要る）に 分けてから返す★。 */
/** ★返事の見出しから 段位と残量を読む（無ければ null＝未測定。0にしない）★ */
function 見出しを読む(h) {
  if (!h) return null;
  const 取る = (k) => {
    try { return typeof h.get === 'function' ? h.get(k) : (h[k] !== undefined ? h[k] : null); }
    catch (e) { return null; }
  };
  const 出 = {};
  for (const [名, 鍵] of [
    ['1分の上限', 'anthropic-ratelimit-requests-limit'],
    ['1分の残り', 'anthropic-ratelimit-requests-remaining'],
    ['戻る時刻', 'anthropic-ratelimit-requests-reset'],
    ['入力の上限', 'anthropic-ratelimit-input-tokens-limit'],
    ['入力の残り', 'anthropic-ratelimit-input-tokens-remaining'],
    ['出力の上限', 'anthropic-ratelimit-output-tokens-limit'],
    ['出力の残り', 'anthropic-ratelimit-output-tokens-remaining'],
  ]) {
    const v = 取る(鍵);
    if (v !== null && v !== undefined && v !== '') 出[名] = v;
  }
  return Object.keys(出).length ? 出 : null;
}

function 失敗を分ける(err) {
  const msg = (err && err.message) || '';
  const st = err && err.status;
  if (/credit balance/i.test(msg) || /insufficient[_ ]?quota/i.test(msg)) return { status: 402, 合言葉: 'zandaka' };
  if (st === 401 || st === 403 || /api[ _-]?key/i.test(msg)) return { status: 401, 合言葉: 'kagi' };
  if (st === 429) return { status: 429, 合言葉: 'komiai' };
  if (err && (err.code === 'ECONNRESET' || err.code === 'ETIMEDOUT' || err.code === 'ETIME')) {
    return { status: 504, 合言葉: 'jikangire' };
  }
  return { status: 502, 合言葉: 'ai_shippai' };
}

// ★テスト用の窓（tests/api-claude.test.mjs が使う）。
//   Vercel は module.exports を「関数として呼ぶ」だけなので、
//   関数に付け足したこのプロパティは本番の挙動を1ミリも変えない。
//   なぜ要るか: 基準数値が黙って NaN / undefined になっても、画面は普通に出てしまう。
//   機械が数値そのものを見るための口。
module.exports.__buildStatutoryPrompt = buildStatutoryPrompt;
/* ★失敗した時に 本当に何を返すかを 機械が押すための窓★
   （本番は module.exports を 関数として呼ぶだけなので 挙動は変わらない） */
module.exports.__失敗を分ける = 失敗を分ける;
module.exports.__許す入口 = 許す入口;
module.exports.__一度に送れる字数 = 一度に送れる字数;
module.exports.__見出しを読む = 見出しを読む;
module.exports.__buildPromptParts = buildPromptParts;
module.exports.__setClient = (c) => { client = c; };
/* ★4 事故止め の窓（試験が 実物を直接 押すため。本番の挙動は 1ミリも変わらない）★ */
module.exports.__事故止め = 事故止め;
module.exports.__誰か = 誰か;
module.exports.__押した回数を見る = 押した回数を見る;
module.exports.__数えておく = 数えておく;
module.exports.__数え場を空にする = __数え場を空にする;
module.exports.__見積もりトークン = 見積もりトークン;
module.exports.__会話を字数で削る = 会話を字数で削る;
module.exports.__会話をトークンで削る = 会話をトークンで削る;
