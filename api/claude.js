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

module.exports = async (req, res) => {
  // CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
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

    // 会話履歴のサニタイズ（不正エントリ除去・最大40メッセージ=20ターンに制限）
    const sanitizedHistory = (Array.isArray(history) ? history : [])
      .filter(m =>
        m &&
        (m.role === 'user' || m.role === 'assistant') &&
        typeof m.content === 'string' &&
        m.content.trim().length > 0
      )
      .slice(-40);

    // バージョンに応じた動的プロンプトを構築
    const versionInfo = getVersionInfo(excelVersion);
    const dynamicPrompt = buildDynamicPrompt(versionInfo);

    const response = await client.messages.create({
      model: 'claude-sonnet-4-6',
      max_tokens: 2000,
      system: dynamicPrompt,
      messages: [...sanitizedHistory, { role: 'user', content: message }],
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
module.exports.__setClient = (c) => { client = c; };
