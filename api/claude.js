const Anthropic = require('@anthropic-ai/sdk');
const https = require('node:https');
/* ★法定の数値は「倉庫」から拾う。★Exally の中に 写しを 置かない★（2026-09-05 司さん）★
 *   司さん「金関係のこと聞かれたりAIが入力する時だけSupabaseの共有から拾うやないんか」
 *
 *  ★倉庫★ = Supabase `public.statutory`（★全アプリ 共通・anon で 読める★）
 *            給与(Rakunally)も 同じ 行を 読む＝★法が 変わった時に 直す所が 1つ★
 *
 *  ★ここまでの 経緯（同じ穴を 3回 掘らない為に 残す）★
 *    2026-08-02 … repo直下に 写しを 置き、掃除で 消して MODULE_NOT_FOUND＝/api/claude が毎回500
 *    2026-09-05 … 給与アプリ(kyuyo/)ごと Exally から 外した。
 *                 ★その時 私は 3本の lib を Exally の lib/ へ 移そうとした＝写しを 残す向き★
 *                 ⇒ 司さんに 止められた。★倉庫から 拾うのが 元からの 設計★。
 */

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

/* ★使えない関数の 一覧は prompt/kansuu.md へ 出しました（2026-09-05）★
   ★手で 書かない★＝`node scripts/make-prompt.mjs` が lib/formula-extra.js の
   台帳から 作る。★手書きの 22個は 17個 間違っていた★（実測）。 */

// ===== バージョン情報取得（デフォルトexcel_365） =====
function getVersionInfo(versionKey) {
  return VERSION_MAP[versionKey] || VERSION_MAP['excel_365'];
}

/* ★SYSTEM_PROMPT_BASE は prompt/base.md へ 出しました（2026-09-05 司さん）★
   ここに べた書きすると 台帳と 二重に なり、★片方だけ 古くなる★。
   ★実際に 起きた★＝手書きの『使えない関数』22個のうち ★17個が 間違い★。 */

/* ══ ★AIの 頭は prompt/ の ファイルから 読む★（2026-09-05 司さん）══════════
 *
 *  ★司さん★「★AIの構造は ファイルにして 更新できるような 設計に しとけよ★」
 *
 *  ★なぜ（実測で 起きていた 事）★
 *    前は この ファイルの 中に ★べた書き★だった。
 *    ⇒ 台帳（`lib/formula-extra.js` の 数える()）と ★二重管理★
 *    ⇒★手書きの「使えない関数」22個のうち 17個が 間違い★
 *        ・足してあるのに「使えない」…★4個★（TOCOL/TOROW/CHOOSEROWS/CHOOSECOLS）
 *        ・動かないのに 1つも 教えていない…★13個★（LAMBDA/LET/MAP/REDUCE/…）
 *    ⇒ しかも `latest` の 決まりは ★「LET・LAMBDA等を 積極的に 提案」★
 *       ＝★AIが 動かない 式を 勧めていた★
 *
 *  ★読めなかったら 500 で 落とす★
 *    ★黙って 前置き 無しで 答えるのが 一番 危ない★＝法定の 時と 同じ 型。
 *    ★合言葉は `shitaku_tarinai`★＝画面は「★うちの 支度が 足りません★」と 出す。
 *    （★`ai_shippai` に しては いけない★＝「AI側の 問題です」は ★嘘★。
 *      客も 管理者も AI を 疑い、★誰も うちを 見ない★＝2026-09-05 指示役）
 *
 *  ★Vercel に 載せる★
 *    `vercel.json` の `functions.includeFiles` で `prompt/**` を 入れる。
 *    ★効いたかは 配信して 初めて 分かる★＝Preview で 1回 叩いて 確かめる。
 */
const 前置きの部屋 = require('path').join(__dirname, '..', 'prompt');
let _前置き = null;

/** ★prompt/ を 読む（1回だけ・読めなければ 投げる）★ */
function 前置きを読む() {
  if (_前置き) return _前置き;
  const fs = require('fs'), path = require('path');
  /* ★2026-09-06 に 2本 足した★＝kiite-ageru.md（聞く順）と nayami.md（客の 悩み）
     ★prompt/README.md に「入れる」と 書いてあったのに ★存在しなかった★★
     ＝会議で 決めた「聞いてあげる」は ★AIに 1文字も 渡っていなかった★ */
  const 要る = ['base.md', 'kansuu.md', 'kyoutsuu.md', 'version.md',
    'kiite-ageru.md', 'nayami.md'];
  const 出 = {};
  const 欠け = [];
  for (const f of 要る) {
    try {
      const t = fs.readFileSync(path.join(前置きの部屋, f), 'utf8');
      if (!t || !t.trim()) { 欠け.push(f + '（空）'); continue; }
      出[f] = t;
    } catch (e) { 欠け.push(f + '（' + (e.code || 'よめない') + '）'); }
  }
  if (欠け.length) {
    /* ★1本でも 欠けたら 止める★＝★前置きが 半分の AI は 嘘を 言う★ */
    const err = new Error('prompt が 読めません: ' + 欠け.join(' / '));
    err.支度足りない = true;
    throw err;
  }
  _前置き = 出;
  return 出;
}

/** ```で 囲った 中身だけ 取り出す（囲いが 無ければ そのまま） */
function 中身だけ(md) {
  const m = /```\r?\n([\s\S]*?)\r?\n```/.exec(md);
  return (m ? m[1] : md).trim();
}

/** 版ごとの 決まり（version.md の `## <group>` を 取る） */
function 版の決まり(md, group) {
  const re = new RegExp('##\\s*' + group + '\\s*\\r?\\n+```\\r?\\n([\\s\\S]*?)\\r?\\n```');
  const m = re.exec(md);
  return m ? m[1].trim() : '';
}

/* ══ ★AI の 答えを「本文」と「表」に 分ける★ ══════════════════════════════
 *
 *  ★2026-09-05 まで こう だった★
 *      const text = fullText.replace(tsvMatch[0], '').trim();   // ★表を 本文から 消す★
 *      return { text, tsv };
 *    ⇒ 画面（book.html:16986）は `d.text` しか 読まない＝★`d.tsv` を 読む 所は
 *      repo 全体で ★0か所★★（履歴でも 一度も 無い＝★消したのでは なく 最初から 片側だけ★）
 *    ⇒★AI が 表を 作る たび、サーバが 抜いて、誰も 受け取らず、捨てていた★
 *    ⇒★答えが 丸ごと 表だった 時は text が 空★ ⇒ 画面は「AIから 空の返事が来ました」
 *      ＝★出来ていたのに「空」と 言う★
 *    ★api/claude.js が 生まれた 日（2026-03-28 5481820）から 約5か月★
 *
 *  ★直し（2026-09-05 指示役の 裁定「丙」）★
 *      ★印の 行（--- TSV_START --- / --- TSV_END ---）だけ 外し、★中身は 本文に 残す★★
 *      ★`tsv` の 欄は そのまま 残す★（消さない＝画面に 表として 貼る「甲」で 使う）
 *      ★前置きは 触らない★（AI には 今までどおり 表を 出させる）
 *
 *  ★やらなかった 事と その 訳★
 *      ・前置きから TSV を やめる … ★却下★（元の 狙いを 消す／記録が 無い＝消してよい 証拠も 無い）
 *      ・表として 画面に 貼る …… ★機能なので 司さん待ち★
 *
 *  ★知っておく 事（正直に）★
 *      本文は `\n` を `<br>` に して 出すだけ（book.html formatAIText）＝
 *      ★タブは HTML で 潰れる＝列が そろって 見えない★。
 *      ただし ★字としては タブが 残る★ので ★選んで Excel に 貼れば 列は 分かれる★。
 *      ★見た目まで 直すのは「甲」★。
 */
function 答えを分ける(fullText) {
  const 全 = String(fullText == null ? '' : fullText);
  /* ★tsv は 今までどおり 最初の 塊を 取る★（欄を 消さない＝甲で 使う） */
  const 塊 = 全.match(/---[ \t]*TSV_START[ \t]*---\r?\n([\s\S]+?)\r?\n---[ \t]*TSV_END[ \t]*---/);
  const tsv = 塊 ? 塊[1].trim() : '';
  /* ★印の 行だけ 消す★（★中身は 1字も 消さない★）。塊が 2つ以上でも 全部 消す。 */
  const 印 = /^[ \t]*---[ \t]*TSV_(?:START|END)[ \t]*---[ \t]*\r?\n?/gm;
  const text = 全.replace(印, '').replace(/\n{3,}/g, '\n\n').trim();
  return { text, tsv };
}

/* ══ ★法定の基準数値＝倉庫(Supabase public.statutory)から 拾う★ ═══════════════
 *  ★AI が 使うのは 4つだけ★／★どの 行から 来るか★（乙・2026-09-05 指示役）
 *    健康保険料率（東京）… kind='shakaihoken'  year=★社保年度(3月起算)★  data.kenko_total.tokyo
 *                          ★倉庫の値は 労使合計★ ⇒ 従業員負担 = ÷2
 *    厚生年金保険料率     … kind='shakaihoken'  year=同上                data.kosei_total
 *                          ★同じく 労使合計★     ⇒ 従業員負担 = ÷2
 *    雇用保険料率（一般） … kind='koyo'         year=★労働保険年度(4月起算)★ data.ippan
 *                          ★これは 最初から 従業員(労働者)負担★＝÷2 しない
 *    消費税（標準/軽減）  … kind='shouhizei'    year=★その kind の 一番 新しい 行★  data.hyojun / data.keigen
 *                          ★2019 と 決め打ちしない★＝倉庫の year は「施行年」で、
 *                          税率が 変われば ★新しい year の 行が 増える★。決め打つと
 *                          ★増えても 2019 を 読み続け、数字は 出るので 誰も 気づかない★。
 *                          ★2019 は「消費税が 今の形に なった 年」＝年度では ない★
 *  ★台帳の year は 種類ごとに 意味が ちがう★（社保年度／労働保険年度／施行年）＝上の通り。
 *
 *  ★年度は 呼ばれる たびに 選び直す★＝関数が 温まったまま 年を またいでも 古い率を 返さない。
 *  ★行そのものは 10分だけ 手元に 置く★（法は 10分で 変わらない／毎回 倉庫を 叩くと AI が 遅くなる）
 *    ＝★repo に 置く 写しとは 別物★（配ってもいないし、次に 冷えたら 消える）
 *
 *  ★拾えなかった 時は 数字を 出さない★（甲・2026-09-05 指示役）
 *    ★黙って 抜くのが 一番 危ない★＝AI は 自分の 記憶の 率を 書いてしまう。
 *    ⇒ 数字の 代わりに ★「今 出せない・推測で 書くな」と AI に 言う★。
 */
const 倉庫の行の寿命ミリ秒 = 10 * 60 * 1000;
let _法定の行 = null, _法定を取った時 = 0;

/** 共有の倉庫の口（★anon で読める 全アプリ共通の棚★＝サーバの鍵は要らない） */
/* ★js/supa-config.js を 読む 所は 1つだけ★（2026-09-05）
   ＝前は ★同じ 注記外しが この ファイルに 2つ★在った。片方だけ 直すと 静かに ずれる。
   ★注記を外してから 読む★＝注記の中の URL を 本物と 読まない。
   （tests/ の 見張りは scripts/lib/chuki.mjs を 使うが、ここは Vercel が 読む
     CommonJS なので ESM を 取り込めない＝この ファイルの 中で 1つに 寄せる） */
function supaConfigの中身() {
  const fs = require('fs'), path = require('path');
  return fs.readFileSync(path.join(__dirname, '..', 'js', 'supa-config.js'), 'utf8')
    .replace(/\/\*[\s\S]*?\*\//g, ' ');
}

let _共有の口 = undefined;
function 共有の口() {
  if (_共有の口 !== undefined) return _共有の口;
  _共有の口 = null;
  try {
    const src = supaConfigの中身();
    const u = /https:\/\/[a-z0-9]+\.supabase\.co/.exec(src);
    const k = /['"](ey[A-Za-z0-9_.-]{40,}|sb_publishable_[A-Za-z0-9_-]{20,})['"]/.exec(src);
    if (u && k) _共有の口 = { url: u[0], key: k[1] };
  } catch (e) { _共有の口 = null; }
  return _共有の口;
}

/** 倉庫から 3種類の行を 拾う。★取れなければ null★（前の値で ごまかさない）
 *  ★fetch を 使わない★＝undici が 線を 繋いだまま 残し、process.exit と ぶつかって
 *    ★node ごと 落ちる（Assertion failed: UV_HANDLE_CLOSING / src\win\async.c:76）★
 *    ＝2026-09-05 実測（tests/api-claude は 49本 全部 緑なのに 終了の印 127）。
 *    ⇒ node:https で 取り、★agent:false で 線を 毎回 閉じる★。Vercel でも 同じに 動く。 */
function 倉庫を取る(u, key, ミリ秒) {
  return new Promise((resolve) => {
    let 済み = false;
    const 終わり = (v) => { if (!済み) { 済み = true; resolve(v); } };
    let req;
    try {
      req = https.get(u, { agent: false, headers: { apikey: key, Authorization: 'Bearer ' + key } }, (res) => {
        if (res.statusCode !== 200) { res.resume(); return 終わり(null); }
        let buf = '';
        res.setEncoding('utf8');
        res.on('data', (c) => { buf += c; });
        res.on('end', () => { try { 終わり(JSON.parse(buf)); } catch (e) { 終わり(null); } });
        res.on('error', () => 終わり(null));
      });
    } catch (e) { return 終わり(null); }
    req.setTimeout(ミリ秒, () => { req.destroy(); 終わり(null); });   /* ★AI を 待たせ続けない★ */
    req.on('error', () => 終わり(null));
  });
}

async function 法定を倉庫から拾う(いま) {
  const now = いま || Date.now();
  if (_法定の行 && (now - _法定を取った時) < 倉庫の行の寿命ミリ秒) return _法定の行;
  const 口 = 共有の口();
  if (!口) return null;
  const u = 口.url + '/rest/v1/statutory?select=kind,year,data'
    /* ★2026-09-06 に 2つ 足した★＝saitei_chingin（最低賃金）と rousai_ritsu（労災率）
       ★会議(09-05)の 穴 D★＝設計書は「倉庫から」と 書いてあったのに
       ★読んでいたのは 4つだけ★ ⇒ AIは この 2つを ★記憶から 書いていた★
       ★表は 前置きに 貼らない★（47都道府県／54業種＝毎回 置き賃）
       ⇒★「在る／年度」だけ 書いて ★数字は 言わせない★★ */
    + '&or=(kind.eq.shakaihoken,kind.eq.koyo,kind.eq.shouhizei,'
    + 'kind.eq.saitei_chingin,kind.eq.rousai_ritsu)';
  const 行 = await 倉庫を取る(u, 口.key, 4000);
  if (!Array.isArray(行) || !行.length) return null;
  _法定の行 = 行; _法定を取った時 = now;
  return 行;
}

/* 年度の 決まり（★数値では なく 暦の 決まり★＝法定の数字では ない）
 *   社保年度 … 3月起算（2026-02 は 令和7年度・2026-03 から 令和8年度）
 *   労働保険年度 … 4月起算（2026-03 は 令和7年度・2026-04 から 令和8年度） */
function 社保年度(ym) { const y = +String(ym).slice(0, 4), m = +String(ym).slice(5, 7); return m >= 3 ? y : y - 1; }
function 労働保険年度(ym) { const y = +String(ym).slice(0, 4), m = +String(ym).slice(5, 7); return m >= 4 ? y : y - 1; }
function 令和(y) { return '令和' + (y - 2018) + '年度'; }

/** その年度の 行を 選ぶ。★無ければ 手前の 一番 新しい 年★（先の年は 使わない） */
function その年度の行(行たち, kind, year) {
  const 候補 = (行たち || []).filter((r) => r.kind === kind && typeof r.year === 'number');
  if (!候補.length) return null;
  const 以下 = 候補.filter((r) => r.year <= year).sort((a, b) => b.year - a.year);
  if (以下.length) return 以下[0];
  return 候補.sort((a, b) => a.year - b.year)[0];   /* 全部 先の年＝一番 手前を 使う */
}

/* ★拾えなかった 時に AI へ 言う 事★（甲）＝★数字を 書かせない★ */
const 法定が取れない時 = `

【税務・給与計算の基準数値】
- ★今この場で 法定の数値（保険料率・税率）を取り出せませんでした★
- 保険料率・税率・税額をたずねられたら、★数字を答えず★「今この数字を出せないので、少し時間をおいてもう一度きいてください」と伝えること
- ★おぼえている率・推測した率を 答えに書かないこと★（古い率で計算すると 給与も税額もまちがう）
- 率を使わない説明（数式の作り方・表の組み立て方）は これまでどおり答えてよい`;

/**
 * 法定の基準数値の 前置きを 組み立てる。
 *   行たち … 法定を倉庫から拾う() の 返り（★null なら 数字を 出さない★）
 *   ymArg  … 対象月 'YYYY-MM'（省略＝今日）
 */
/** その kind の ★一番 新しい 年の 行★（施行年で 持っている 種類＝消費税 など） */
function そのkindの一番新しい行(行たち, kind) {
  const 候補 = (行たち || []).filter((r) => r.kind === kind && typeof r.year === 'number');
  if (!候補.length) return null;
  return 候補.sort((a, b) => b.year - a.year)[0];
}

function buildStatutoryPrompt(ymArg, 行たち) {
  const ym = ymArg || new Date().toISOString().slice(0, 7);
  const 社保 = その年度の行(行たち, 'shakaihoken', 社保年度(ym));
  const 雇用 = その年度の行(行たち, 'koyo', 労働保険年度(ym));
  /* ★消費税は「年度」では なく「施行年」★＝対象月から 選ぶ 物では ない。
     ⇒★その kind の 一番 新しい 行★を 取る（★2019 と 書かない★）。 */
  const 消費 = そのkindの一番新しい行(行たち, 'shouhizei');
  const 健保合計 = 社保 && 社保.data && 社保.data.kenko_total && 社保.data.kenko_total.tokyo;
  const 厚年合計 = 社保 && 社保.data && 社保.data.kosei_total;
  const 雇用率 = 雇用 && 雇用.data && 雇用.data.ippan;
  const 標準 = 消費 && 消費.data && 消費.data.hyojun;
  const 軽減 = 消費 && 消費.data && 消費.data.keigen;
  /* ★1つでも 欠けたら 全部 出さない★＝半分だけの 表は 一番 危ない */
  /* ★最低賃金・労災率は ★表が 大きい★（47都道府県／54業種）＝前置きに 貼らない★
     ⇒★在るか どうかと 年度だけ★ 前置きに 出し、★数字は 言わせない★
     ⇒★1行でも 欠けたら 全部 出さない★＝半分だけの 表は 一番 危ない（下と 同じ 決まり） */
  const 最賃 = そのkindの一番新しい行(行たち, 'saitei_chingin');
  const 労災 = そのkindの一番新しい行(行たち, 'rousai_ritsu');
  /* ★ここは 数字では なく「在るか／年度」だけ★＝欠けても 消費税まで 消すのは 行き過ぎ
     ⇒★取れなければ「取れませんでした」と 書いて ★同じく 数字を 言わせない★★
     （★半分の 表を 出さない★の 決まりは ★数字を 出す 5つ★に かかる） */
  const 最賃の年度 = 最賃 && 最賃.year ? 令和(最賃.year) : '★倉庫から 取れませんでした★';
  const 労災の年度 = 労災 && 労災.year ? 令和(労災.year) : '★倉庫から 取れませんでした★';
  const 揃った = [健保合計, 厚年合計, 雇用率, 標準, 軽減].every((v) => typeof v === 'number' && isFinite(v));
  if (!揃った) return 法定が取れない時;
  return `

【税務・給与計算の基準数値】
- 健康保険料率（東京）: ${(健保合計 / 2 * 100).toFixed(3)}%（従業員負担・労使折半・${令和(社保.year)}）
- 厚生年金保険料率: ${(厚年合計 / 2 * 100).toFixed(2)}%（従業員負担・労使折半・全国一律）
- 雇用保険料率: ${(雇用率 * 100).toFixed(2)}%（従業員負担・一般の事業・${令和(雇用.year)}）
- 消費税: ${(標準 * 100).toFixed(0)}%（標準）/ ${(軽減 * 100).toFixed(0)}%（軽減）

【★数字を 言っては いけない 物★】
- ★最低賃金★ … 倉庫（${最賃の年度}）に 在るが ★都道府県ごとに 違う★。
  ⇒ ★金額を 言わない★。まず ★どの 都道府県か★を 聞く。
  ⇒ 聞いた上でも ★この 前置きには 表が 無い★ので、
    「その 県の 額は こちらで 出します」と 言い、★自分の 記憶から 書かない★。
- ★労災保険率★ … 倉庫（${労災の年度}）に 在るが ★事業の 種類ごとに 違う★。
  ⇒ ★率を 言わない★。まず ★何の 事業か★を 聞く。
- ★源泉徴収の 率・税額★ … ★倉庫に 入っていない★。
  ⇒ ★率も 税額も 言わない★（10.21% などを 記憶から 書かない）。
  ⇒「国税庁の 表に 当てる 必要が あるので、うちでは 数字を 出しません」と 言う。
- 上の どれも ★率を 使わない 説明（表の 組み立て方・数式）は これまでどおり 答えてよい★。`;
}

// ===== 動的プロンプト生成 =====
function buildDynamicPrompt(versionInfo, 法定の行) {
  const 部品 = buildPromptParts(versionInfo, 法定の行);
  return 部品.共通 + 部品.版ごと;
}

/* ★前置きを「置いたまま使い回す」ために 2つに分ける（2026-08-22）★
   ・前半＝★どの版の人にも同じ★（作りの説明＋法定の基準数値）＝ここを置いたままにする
   ・後半＝★版ごとに変わる★（Excel 365 / 2016 / 持っていない …）＝置き場所の後ろに回す
   ★変わる物を前に置くと 毎回 置き直しになって 逆に高くなる★（一次情報の決まり） */
function buildPromptParts(versionInfo, 法定の行) {
  const P = 前置きを読む();                       /* ★読めなければ 投げる★ */
  /* ★``` の 中だけ 渡す★＝人向けの 注記（作り方・日付）を AIへ 送らない。
     ★①は 置いたまま 使い回す 所★なので、日付が 混ざると ★毎回 置き直し＝毎回 置き賃★。
     ⇒ 実際に 踏んだ（2026-09-05）＝`tests/api-claude.test.mjs` の
       「①に 毎回変わる物が 混ざっていない」が 赤に なった。 */
  /* ★聞く順と 悩みは ★版で 変わらない★＝①（置いたまま 使い回す 所）に 入れる★
     ⇒ 変わる物を 前に 置くと 毎回 置き直し＝毎回 置き賃（一次情報の 決まり） */
  const 共通 = 中身だけ(P['base.md'])
    + '\n\n' + 中身だけ(P['kansuu.md'])
    + '\n\n' + 中身だけ(P['nayami.md'])
    + '\n\n' + 中身だけ(P['kiite-ageru.md'])
    + buildStatutoryPrompt(null, 法定の行);
  const 版ごと = '\n' + 版の決まり(P['version.md'], versionInfo.group)
    + '\n' + 中身だけ(P['kyoutsuu.md']);
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
/* ★回数と 字数は lib/ai-genkai.js の 1か所から 読む★（2026-09-05）
   ＝ここに 書いてある「数字は ここ1か所だけ」を ★掘る側も 含めて★ 本当にした。
   （前は 掘る数だけ lib/horu.js に 在り、15回 叩いて 10回で 止まっていた） */
const AI限界 = require('../lib/ai-genkai.js').限界;

const 事故止め = {
  分の回数: AI限界.分の回数,          // ★1分に10回（人ごと）★
  分の窓ミリ秒: 60 * 1000,
  日の回数: 100,         // ★1日に100回（人ごと）★
  日の窓ミリ秒: 24 * 60 * 60 * 1000,
  会話の合計字数: 40000, // ★history 合計40,000字＝古い方から捨てる★
  渡せるトークン: 20000, // ★1回にAIへ渡すのは2万トークンまで（2026-08-09 司さんの決定）★
  /* ★1回に送れる字数（司さん承認 2026-08-22）★＝たまたま同じ数だが 別の物なので 別々に持つ */
  一度に送れる字数: 20000,
};
const 一度に送れる字数 = 事故止め.一度に送れる字数;

/* ══ ★数え場（2026-08-26 指示役が 鍵と 本番の表を 入れた）★ ══════════════
   ★倉庫（exally.ai_tsukatta）に 1回ぶん 1行 書く★＝機械が増えても すり抜けない。
   ★決まり★
     ①★書けなくても AIは止めない★（数え場が落ちて 客が使えなくなる形にしない）
     ②★向き先は この repo の js/supa-config.js から読む★
        ＝本番の repo なら 本番の倉庫、テスト線なら DB-test。
        ★環境変数の URL が repo の向き先と違ったら 1行も書かない★（混ざりを止める）
     ③★書けない時は 今までどおり「この機械の中だけ」で数える★（0件・異常なしにしない）
     ④★90日の掃除は 書き込む所と同じ所で回す★＝その日の最初の書き込みで 1回だけ */
const 数え場 = new Map();
const __数え場を空にする = () => { 数え場.clear(); 掃除した日 = ''; };

/** ★倉庫の向き先（この repo の js/supa-config.js が 唯一の正）★ */
let _repoの倉庫 = undefined;
function repoの倉庫() {
  if (_repoの倉庫 !== undefined) return _repoの倉庫;
  try {
    const src = supaConfigの中身();
    const m = /https:\/\/([a-z0-9]+)\.supabase\.co/.exec(src);
    _repoの倉庫 = m ? m[1] : null;
  } catch (e) { _repoの倉庫 = null; }
  return _repoの倉庫;
}

/** ★倉庫が使えるか★（鍵が無い・向き先が違う なら 使わない） */
function 倉庫の口() {
  const url = process.env.SUPABASE_URL || '';
  const key = process.env.SUPABASE_SERVICE_ROLE_KEY || '';
  if (!url || !key) return null;
  const m = /https:\/\/([a-z0-9]+)\.supabase\.co/.exec(url);
  const 先 = m ? m[1] : null;
  const 正 = repoの倉庫();
  /* ★repo の向き先と違う倉庫には 1行も書かない★（環境の混ざりを止める） */
  if (!先 || (正 && 先 !== 正)) return null;
  return { url: url.replace(/\/+$/, ''), key: key };
}

const 倉庫の頭 = (口, 足す) => Object.assign({
  apikey: 口.key,
  Authorization: 'Bearer ' + 口.key,
  'Accept-Profile': 'exally',
  'Content-Profile': 'exally',
  'Content-Type': 'application/json',
}, 足す || {});

/** ★倉庫で 直近◯ミリ秒の「通した回数」を数える★（数えられなければ null＝未測定） */
async function 倉庫で数える(鍵, 窓ミリ秒, いま) {
  const 口 = 倉庫の口();
  if (!口) return null;
  const から = new Date(いま - 窓ミリ秒).toISOString();
  const u = 口.url + '/rest/v1/ai_tsukatta?select=id&hito=eq.' + encodeURIComponent(鍵)
    + '&kekka=eq.ok&oshita=gte.' + encodeURIComponent(から);
  const res = await fetch(u, { headers: 倉庫の頭(口, { Prefer: 'count=exact', Range: '0-0' }) });
  if (!res.ok) return null;
  const cr = res.headers.get('content-range') || '';
  const m = /\/(\d+)$/.exec(cr);
  return m ? Number(m[1]) : null;
}

/** ★倉庫へ 1行 書く★（書けなくても AIは止めない） */
async function 倉庫へ書く(行) {
  const 口 = 倉庫の口();
  if (!口) return false;
  try {
    const res = await fetch(口.url + '/rest/v1/ai_tsukatta', {
      method: 'POST',
      headers: 倉庫の頭(口, { Prefer: 'return=minimal' }),
      body: JSON.stringify(行),
    });
    return res.ok;
  } catch (e) { return false; }
}

/* ★90日の掃除★＝★その日の最初の書き込みで 1回だけ★（外の見張りは この repo に鍵が無く 回らない） */
let 掃除した日 = '';
async function 掃除する(いま) {
  const 口 = 倉庫の口();
  if (!口) return null;
  const 今日 = new Date(いま).toISOString().slice(0, 10);
  if (掃除した日 === 今日) return null;
  掃除した日 = 今日;
  try {
    const より前 = new Date(いま - 90 * 24 * 60 * 60 * 1000).toISOString();
    const res = await fetch(口.url + '/rest/v1/ai_tsukatta?oshita=lt.' + encodeURIComponent(より前), {
      method: 'DELETE',
      headers: 倉庫の頭(口, { Prefer: 'return=minimal' }),
    });
    return res.ok;
  } catch (e) { return false; }
}

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

/** ★倉庫でも数える（在れば そちらが 正）★
 *  ★倉庫が答えられない時は 機械の中の数で決める★（0件・異常なしにしない）
 *  @returns {{止める:boolean, あと秒:number, どれ:string, どこで:string}}
 */
async function 倉庫も見て数える(鍵たち, いま) {
  const 中 = 押した回数を見る(鍵たち, いま);
  if (中.止める) return Object.assign({ どこで: '機械の中' }, 中);
  if (!倉庫の口()) return Object.assign({ どこで: '機械の中' }, 中);
  try {
    for (const k of 鍵たち) {
      const 分 = await 倉庫で数える(k, 事故止め.分の窓ミリ秒, いま);
      if (分 !== null && 分 >= 事故止め.分の回数) {
        return { 止める: true, あと秒: 60, どれ: '分', どこで: '倉庫' };
      }
      const 日 = await 倉庫で数える(k, 事故止め.日の窓ミリ秒, いま);
      if (日 !== null && 日 >= 事故止め.日の回数) {
        return { 止める: true, あと秒: 3600, どれ: '日', どこで: '倉庫' };
      }
    }
  } catch (e) {
    /* ★倉庫が落ちても AIは止めない★（数えるのは 機械の中の分で続ける） */
    return Object.assign({ どこで: '機械の中（倉庫は未測定）' }, 中);
  }
  return Object.assign({ どこで: '倉庫' }, 中);
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
      /* ★どこで数えたか（倉庫／機械の中）★＝すり抜けているのか 効いているのかが 見える */
      どこで: o.どこで || '',
      かかった秒: o.秒,
      入口: o.入口 || '(名乗りなし)',
    }));
  } catch (e) { /* 記録で本体を落とさない */ }
  /* ★倉庫にも 1行 残す★（機械が増えても すり抜けない・9のクレジットも 同じ表に乗る）
     ★書けなくても AIは止めない★＝待たないし、失敗しても 何も起きない。
     ★人が書いた文そのものは 1文字も入れない★（長さと数だけ） */
  try {
    if (o.誰 && o.誰.length) {
      for (const k of o.誰) {
        倉庫へ書く({
          hito: k,
          hito_kind: k.slice(0, 2) === 'u:' ? 'user' : 'ip',
          kekka: o.結果,
          nyuryoku_token: o.入力 || 0,
          shutsuryoku_token: o.出力 || 0,
          oita_token: o.置いた || 0,
          yominaoshita_token: o.読み直した || 0,
          watashita_mitsumori: o.渡した見積もりトークン || 0,
          kaiwa_kezutta: o.会話を削った || 0,
          iriguchi: o.入口 || null,
          credit: 0,
        }).catch(() => {});
      }
    }
  } catch (e) { /* 倉庫で本体を落とさない */ }
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
    const 見張り = await 倉庫も見て数える(誰, いま);
    if (見張り.止める) {
      記録({ 結果: 'tsukaisugi', 字数: message.length, 待ち秒: 見張り.あと秒, どれ: 見張り.どれ,
              どこで: 見張り.どこで, 誰: 誰, 入口: 入口 });
      res.setHeader('Retry-After', String(見張り.あと秒));
      return res.status(429).json({
        error: 'tsukaisugi', どれ: 見張り.どれ, 待ち秒: 見張り.あと秒, text: '', tsv: '',
      });
    }
    数えておく(誰, いま);
    /* ★90日の掃除は 書き込む所と同じ所で回す★（その日の最初の1回だけ） */
    掃除する(いま).catch(() => {});

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
    /* ★法定の数字は 押された その時に 倉庫から 拾う★（司さん 2026-09-05）
       ★取れなければ null★＝下の 前置きが「数字を出すな」に 切り替わる */
    const 法定の行 = await 法定を倉庫から拾う();
    const dynamicPrompt = buildDynamicPrompt(versionInfo, 法定の行);

    /* ★置いたまま使い回す（prompt caching）★ 2026-08-22
       なぜ … 会話40件(20往復)を ★毎回まるごと送り直していた★＝1回 平均 約4円。
              個人1,280円/月500回 なら 原価2,000円＝★1人目から赤字★（指示役の実測）。
       やり方 … ①前置きの共通部分 ②版ごとの部分 ③★前の会話まで★ の3か所に印を付ける。
              ★印は最大4か所★／★読み直し 0.1倍・置く時 1.25倍（5分もつ）★＝一次情報。
              ★Sonnet 4.6 は 1,024トークン未満だと 黙って置かれない★ので 前置きは1つに束ねる。
       ★客の画面は 1文字も変えていない（我慢も 上限も していない）★ */
    const 部品 = buildPromptParts(versionInfo, 法定の行);
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
      誰: 誰,
      入口: 入口,
    });

    const fullText = response.content
      .filter(block => block.type === 'text')
      .map(block => block.text)
      .join('');

    const 分け = 答えを分ける(fullText);
    return res.status(200).json({ text: 分け.text, tsv: 分け.tsv });

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
      誰: (typeof 誰 !== 'undefined') ? 誰 : null,
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
  /* ★うちの 支度が 足りない（prompt/ が 読めない）★（2026-09-05 指示役）
     ★一番 先に 見る★＝後ろに 置くと ai_shippai に 落ちて ★「AI側の 問題です」と 嘘を 言う★。
     ⇒ 客も 管理者も AI を 疑い ★誰も うちを 見ない★。 */
  if (err && err.支度足りない) return { status: 500, 合言葉: 'shitaku_tarinai' };
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
module.exports.__答えを分ける = 答えを分ける;
module.exports.__法定を倉庫から拾う = 法定を倉庫から拾う;
module.exports.__法定が取れない時 = 法定が取れない時;
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
