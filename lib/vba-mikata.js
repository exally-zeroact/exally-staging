/* vba-mikata.js — ★マクロの見立て★＝取り出したVBAを ★機械が★ 分ける（AIは0回）
 *
 *  ★なぜ在るか（司さん 2026-08-28）★
 *    「マクロを組んだ人が 辞めて 誰も触れない」＝ここが 一番 困っている所。
 *    だから ★出来ない事は 無い★状態にする。まず ★何をしているのか 日本語で 出す★。
 *
 *  ★決まり（指示役 2026-08-28）★
 *    ・分けるのは ★機械★（0円）。AIには ★要約だけ★ 渡す（中身の全文は 渡さない）。
 *    ・1本ずつ ★可否★を出す … 今できる ／ 置き換える ／ 別の道
 *    ・★VBAは 動かさない・書き換えない★。ここは ★読んで 言うだけ★。
 *    ・★分からない物は「分からない」と言う★＝当てずっぽうで 分類しない。
 *
 *  ★言い方★
 *    ・客に見せる字に ★ を書かない（★は うちの覚え書きの印）
 *    ・「壊れています」と言わない。「何をしているか」と「うちでは こうします」を言う。
 */
(function (root) {
  'use strict';

  /* ══ ①手続き（Sub / Function）に 切る ══════════════════════ */
  /** VBAの中身 → [{名, 種類, 行, 中身, 行数}]（宣言だけの頭は 一覧に出さない） */
  function 手続きに切る(文) {
    var 行たち = String(文 || '').split(/\r\n|\r|\n/);
    var 出 = [], 今 = null;
    var 始 = /^\s*(?:(Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+([^\s(]+)/i;
    var 終 = /^\s*End\s+(Sub|Function|Property)\s*$/i;
    /** 今 開いている物を 閉じる（End が 無かった時も ここを通す） */
    function 閉じる(閉じていない) {
      if (!今) return;
      今.中身 = 今.中身.join('\n');
      今.行数 = 今.中身.split('\n').length;
      今.閉じていない = !!閉じていない;
      出.push(今);
      今 = null;
    }
    for (var i = 0; i < 行たち.length; i++) {
      var l = 行たち[i];
      var m = 始.exec(l);
      /* ★End が 無いまま 次の Sub が 始まる事が 実物に ある★（司さんの実物で 5件）。
         ★黙って 1本に まとめると 次の手続きの中身が 前の手続きの物として 数えられる★。
         ⇒ ここで 前の物を 閉じる（★閉じていない と 覚えておく★）。 */
      if (m && 今) 閉じる(true);
      if (m && !今) {
        今 = {
          名: m[3],
          種類: /^Sub$/i.test(m[2]) ? '手続き' : (/^Function$/i.test(m[2]) ? '関数' : '性質'),
          見せるか: !/^Private$/i.test(m[1] || ''),
          行: i + 1,
          中身: [l],
        };
        continue;
      }
      if (今) {
        今.中身.push(l);
        if (終.test(l)) 閉じる(false);
      }
    }
    閉じる(true);
    return 出;
  }

  /* ══ ②分け方の表（★ここが 本体★） ═════════════════════════
     可否 … 'できる'   ＝ うちに 同じ物が 在る（そのまま置き換わる）
            'かえる'   ＝ 同じ事を うちの別のやり方で やる（形は変わる）
            'べつの道' ＝ ブラウザの中では 出来ない（うちの外に 道を作る話）
     ★手掛かりは「VBAの言葉」で書く。日本語の変数名では 判定しない★ */
  var 分け方 = [
    /* ── 表をいじる（うちに 在る） ───────────────────────── */
    { key: 'narabekae', 印: '.Sort', 名: '並べ替え', 可否: 'できる',
      手掛かり: [/\.Sort\b/i, /\bSortFields\b/i, /\bxlAscending\b|\bxlDescending\b/i],
      何を: '表を 並べ替えています。', うち: '並べ替え（列を選んで 昇順・降順）' },
    { key: 'shiborikomi', 印: 'AutoFilter', 名: '絞り込み', 可否: 'できる',
      手掛かり: [/\bAutoFilter\b/i, /\bAdvancedFilter\b/i, /\bShowAllData\b/i],
      何を: '条件で 行を 絞り込んでいます。', うち: '絞り込み（列ごとの条件）' },
    { key: 'shukei', 印: 'WorksheetFunction.Sum / SUM', 名: '集計', 可否: 'できる',
      手掛かり: [/WorksheetFunction\.(Sum|Count|Average|Max|Min|SumIf|CountIf)/i, /\bSubtotal\b/i, /"=SUM\(/i],
      何を: '合計や 件数を 出しています。', うち: '集計（合計・件数・平均）' },
    { key: 'pivot', 印: 'PivotTable', 名: 'クロス集計', 可否: 'できる',
      手掛かり: [/\bPivotTable/i, /\bPivotCache/i, /\bPivotFields\b/i],
      何を: '縦横に 集計しています（ピボット）。', うち: 'クロス集計' },
    { key: 'juufuku', 印: 'RemoveDuplicates', 名: '重複を消す', 可否: 'できる',
      手掛かり: [/\bRemoveDuplicates\b/i],
      何を: '同じ行を 消しています。', うち: '重複を消す' },
    { key: 'sagasu', 印: '.Find / .Replace', 名: '探す・置き換える', 可否: 'できる',
      手掛かり: [/\.Find\s*\(/i, /\.Replace\s*\(/i, /\bFindNext\b/i],
      何を: '字を 探して 置き換えています。', うち: '探す／置き換える' },
    { key: 'shoshiki', 印: '.Interior.Color / .NumberFormat', 名: '色や書式', 可否: 'できる',
      手掛かり: [/\.Interior\.(Color|ColorIndex|Pattern)/i, /\.Font\.(Bold|Color|Size)/i, /\.NumberFormat\b/i, /\bFormatConditions\b/i],
      何を: '色や 表示の形を 変えています。', うち: '条件つきの色・表示の形' },
    { key: 'insatsu', 印: 'PrintOut', 名: '印刷', 可否: 'できる',
      手掛かり: [/\bPrintOut\b/i, /\bPageSetup\b/i, /\bPrintPreview\b/i],
      何を: '印刷しています。', うち: '印刷（紙だけの窓で 刷る）' },
    { key: 'kakidashi', 印: 'SaveAs / ExportAsFixedFormat', 名: '書き出し', 可否: 'できる',
      手掛かり: [/\bExportAsFixedFormat\b/i, /\bSaveAs\b/i, /\bSaveCopyAs\b/i],
      何を: 'ファイルに 書き出しています。', うち: '書き出し（Excel／PDF／CSV）' },
    { key: 'torikomi', 印: 'GetOpenFilename / Workbooks.Open', 名: '取り込み', 可否: 'できる',
      手掛かり: [/\bGetOpenFilename\b/i, /Workbooks\.Open\b/i, /\bQueryTables\b/i, /\bTextToColumns\b/i],
      何を: '別のファイルを 読み込んで 貼っています。', うち: '取り込み（Excel／CSV）' },
    { key: 'nyuuryoku_kimari', 印: '.Validation', 名: '入力の決まり', 可否: 'できる',
      手掛かり: [/\.Validation\b/i, /\bxlValidateList\b/i],
      何を: '入力できる値を 決めています。', うち: '入力の決まり（一覧から選ぶ）' },
    { key: 'shiito', 印: 'Sheets.Add', 名: 'シートを作る・名前を変える', 可否: 'できる',
      手掛かり: [/Sheets\.Add\b/i, /Worksheets\.Add\b/i, /\.Name\s*=\s*"/i],
      何を: 'シートを 作ったり 名前を 変えたりしています。', うち: 'シートを作る・名前を変える' },
    { key: 'gyou_retsu', 印: 'Rows/Columns.Insert', 名: '行・列をいじる', 可否: 'できる',
      手掛かり: [/Rows\([^)]*\)\.(Insert|Delete|Hidden)/i, /Columns\([^)]*\)\.(Insert|Delete|Hidden|AutoFit)/i, /\bEntireRow\b/i, /\bEntireColumn\b/i],
      何を: '行や列を 足したり 消したり 隠したりしています。', うち: '行・列の 追加／削除／隠す' },
    { key: 'utsusu', 印: '.Copy / .Value = .Value', 名: '写して貼る', 可否: 'できる',
      手掛かり: [/\.Copy\b/i, /\bPasteSpecial\b/i, /\.Value\s*=\s*[^=]*\.Value\b/i],
      何を: '別の所へ 写しています。', うち: 'レシピ（写す手順を 覚えさせる）' },
    { key: 'kesu', 印: 'ClearContents', 名: '消す', 可否: 'できる',
      手掛かり: [/\bClearContents\b/i, /\bClearFormats\b/i, /\.Clear\b/i],
      何を: '中身を 消しています。', うち: '消す（レシピの1手順）' },

    /* ── やり方を かえる ─────────────────────────────────── */
    { key: 'ivento', 印: 'Worksheet_Change / Workbook_Open', 名: '自動で動く（きっかけ）', 可否: 'かえる', 先に見る: true,
      手掛かり: [/\bWorksheet_(Change|SelectionChange|BeforeDoubleClick|Activate)\b/i,
        /\bWorkbook_(Open|BeforeSave|BeforeClose|SheetChange)\b/i, /\bAuto_(Open|Close)\b/i],
      何を: '人が 何かした時に 自動で 動いています。', うち: '式（打った時に そのまま 出る）' },
    { key: 'hitotsuzutsu', 印: 'For Each / Do While', 名: '1つずつ回している', 可否: 'かえる',
      手掛かり: [/\bFor\s+Each\b/i, /\bFor\s+\w+\s*=\s*[^\n]*\bTo\b/i, /\bDo\s+While\b/i, /\bDo\s+Until\b/i],
      何を: 'セルを 1つずつ 見て回っています。', うち: '式か レシピ（1つずつ回さずに 一度に）' },
    { key: 'erabu',後回し: true, 印: '.Select / Selection', 名: '選んでから作業', 可否: 'かえる',
      手掛かり: [/\.Select\b/i, /\.Activate\b/i, /\bSelection\b/i],
      何を: 'セルを 選んでから 作業しています。', うち: '選ばずに 場所を指して 直す（要らない手順）' },
    { key: 'chiratsuki',後回し: true, 印: 'ScreenUpdating', 名: 'ちらつき止め', 可否: 'かえる',
      手掛かり: [/ScreenUpdating\b/i, /\bCalculation\s*=\s*xl/i, /EnableEvents\b/i],
      何を: '画面のちらつきや 再計算を 止めています。', うち: 'うちでは 要りません（元から 止まりません）' },
    { key: 'shirase', 印: 'MsgBox / InputBox', 名: '知らせ・聞く', 可否: 'かえる',
      手掛かり: [/\bMsgBox\b/i, /\bInputBox\b/i],
      何を: '画面に 知らせを出したり 人に 聞いたりしています。', うち: '知らせの出口（1つ）／聞く形' },
    { key: 'namae', 印: 'Names.Add', 名: '名前をつける', 可否: 'かえる',
      手掛かり: [/Names\.Add\b/i, /\bThisWorkbook\.Names\b/i],
      何を: '範囲に 名前を つけています。', うち: '名前つきの範囲' },
    { key: 'keisan',後回し: true, 印: '.Calculate', 名: '計算のやり直し', 可否: 'かえる',
      手掛かり: [/\.Calculate\b/i, /\bCalculateFull\b/i],
      何を: '計算を やり直させています。', うち: 'うちでは 要りません（いつも 出来ています）' },

    /* ── 別の道（ブラウザの中では 出来ない） ──────────────── */
    { key: 'mail', 印: 'Outlook.Application / SendMail', 名: 'メールを出す', 可否: 'べつの道',
      手掛かり: [/Outlook\.Application/i, /\bCDO\./i, /\.SendMail\b/i, /MailEnvelope/i],
      何を: 'メールを 出しています。', うち: 'うちの外（送る仕組みを 別に作る話）' },
    { key: 'hoka_no_app', 印: 'CreateObject / Shell', 名: '他のアプリを動かす', 可否: 'べつの道',
      手掛かり: [/CreateObject\s*\(/i, /GetObject\s*\(/i, /\bShell\s*\(/i, /Word\.Application/i, /Access\.Application/i],
      何を: '別のアプリを 動かしています。', うち: 'うちの外（何をさせているかを 見て 別の道を作る）' },
    { key: 'file', 印: 'FileSystemObject / Open For', 名: 'ファイルを直に触る', 可否: 'べつの道',
      手掛かり: [/FileSystemObject/i, /\bOpen\s+[^\n]*\bFor\s+(Input|Output|Append)\b/i, /\bKill\b/i, /\bMkDir\b/i, /\bDir\s*\(/i],
      何を: 'パソコンの中の ファイルを 直に 触っています。', うち: 'ブラウザからは 触れません（取り込み／書き出しで 置き換える）' },
    { key: 'db', 印: 'ADODB / Recordset', 名: 'データベースにつなぐ', 可否: 'べつの道',
      手掛かり: [/ADODB\./i, /\bRecordset\b/i, /\bConnectionString\b/i, /\bODBC\b/i],
      何を: 'データベースに つないでいます。', うち: 'うちの外（つなぎ先を 見て 決める話）' },
    { key: 'web', 印: 'XMLHTTP', 名: 'ネットから取る', 可否: 'べつの道',
      手掛かり: [/XMLHTTP/i, /InternetExplorer/i, /WinHttp/i, /\bMSXML2\b/i],
      何を: 'ネットから 取ってきています。', うち: 'うちの外（取り先を 見て 決める話）' },
    { key: 'form', 印: 'UserForm', 名: '自作の画面', 可否: 'べつの道',
      手掛かり: [/UserForm/i, /\.Show\s+vbModeless/i, /\bControls\s*\(/i],
      何を: '自分で作った 入力画面を 出しています。', うち: '同じ画面を うちで 作り直す話' },
    { key: 'api', 印: 'Declare Lib', 名: 'Windowsの機能を直に呼ぶ', 可否: 'べつの道',
      手掛かり: [/\bDeclare\s+(PtrSafe\s+)?(Function|Sub)\b/i, /\bLib\s+"/i],
      何を: 'Windowsの機能を 直に 呼んでいます。', うち: 'ブラウザからは 呼べません（何のためかを 見て 別の道）' },
    { key: 'jikan', 印: 'Application.OnTime', 名: '時間で動かす', 可否: 'べつの道',
      手掛かり: [/Application\.OnTime\b/i, /\bTimer\b/i],
      何を: '決まった時刻に 動かしています。', うち: 'うちの外（決まった時刻に動かす仕組みを 別に作る）' },
    { key: 'hogo', 印: 'Protect / Password', 名: '保護・合い言葉', 可否: 'べつの道',
      手掛かり: [/\bUnprotect\b/i, /\bProtect\b/i, /\bPassword\s*:?=/i],
      何を: 'シートの保護や 合い言葉を 触っています。', うち: '入れる人を 決める形（合い言葉ではなく 入口で分ける）' },
  ];

  /* ══ ③1本を 見立てる ═══════════════════════════════════ */
  function 見立てる1本(手続き) {
    var 中 = String(手続き.中身 || '');
    var 当たり = [];
    for (var i = 0; i < 分け方.length; i++) {
      var d = 分け方[i], n = 0;
      for (var j = 0; j < d.手掛かり.length; j++) {
        var re = new RegExp(d.手掛かり[j].source, d.手掛かり[j].flags.replace(/g/g, '') + 'g');
        var m = 中.match(re);
        if (m) n += m.length;
      }
      /* ★手掛かりに 客のコードを 入れない★＝出すのは うちの表に書いた 印（決め打ちの字）だけ。
         2026-08-28 実測＝合わせた字をそのまま出したら AIに渡す物に
         客のシート名（Sheets("…")）が 入っていた。 */
      if (n) 当たり.push({ key: d.key, 名: d.名, 可否: d.可否, 何を: d.何を, うち: d.うち, 数: n, 手掛かり: d.印 || d.key, 先に見る: !!d.先に見る, 後回し: !!d.後回し });
    }
    /* ★並べる順は 数だけで 決めない★
       ①きっかけ（いつ動くか）が 先 ②仕事の中身 ③段取りの話（ちらつき止め・選んでから・再計算）
       ＝数だけで 並べると ★ScreenUpdating が 2回 出てくるだけで 先頭に来る★（実測で そうなった）。
       客が 知りたいのは ★何をしている マクロか★であって 段取りではない。 */
    当たり.sort(function (a, b) {
      var 順 = function (x) { return x.先に見る ? 0 : (x.後回し ? 2 : 1); };
      if (順(a) !== 順(b)) return 順(a) - 順(b);
      return b.数 - a.数;
    });
    /* ★分からない物は「分からない」と言う★＝当てずっぽうで 分類しない */
    if (!当たり.length) {
      return {
        名: 手続き.名, 種類: 手続き.種類, 行: 手続き.行, 行数: 手続き.行数,
        閉じていない: !!手続き.閉じていない,
        /* ★中身は 持っておく★＝手順を 取り出す側（lib/vba-tejun.js）が 使う。
           ★AIに渡す形には 入れない★（AIに渡す形() が 拾う物を 決めている・試験で 見張る）。 */
        中身: 手続き.中身,
        分類: [], 可否: 'わからない',
        何をしているか: '何をしているか 読み取れませんでした。中身を 見て 決めてください。',
        うちのやり方: '', 手掛かり: [],
      };
    }
    /* ★可否は 一番 重い物に合わせる★（別の道 が1つでも在れば その手続きは 別の道） */
    var 重さ = { 'できる': 0, 'かえる': 1, 'べつの道': 2 };
    var 可否 = 'できる';
    for (var k = 0; k < 当たり.length; k++) if (重さ[当たり[k].可否] > 重さ[可否]) 可否 = 当たり[k].可否;
    var 文 = [], やり方 = [];
    for (var q = 0; q < 当たり.length && q < 4; q++) {
      if (文.indexOf(当たり[q].何を) < 0) 文.push(当たり[q].何を);
      if (当たり[q].うち && やり方.indexOf(当たり[q].うち) < 0) やり方.push(当たり[q].うち);
    }
    return {
      名: 手続き.名, 種類: 手続き.種類, 行: 手続き.行, 行数: 手続き.行数,
      閉じていない: !!手続き.閉じていない,
      中身: 手続き.中身,        /* ★手順を 取り出す側が 使う（AIには 渡さない）★ */
      分類: 当たり, 可否: 可否,
      何をしているか: 文.join(''),
      うちのやり方: やり方.join('／'),
      手掛かり: 当たり.map(function (a) { return a.手掛かり; }).filter(Boolean).slice(0, 6),
    };
  }

  /* ══ ④まとめて 見立てる ══════════════════════════════════
     @param モジュール … vba.js の 読む() が返す モジュールの配列
     ★確か:false（正しく読めたと 言えない物）は 見立てない＝未測定にする★ */
  function 見立てる(モジュール) {
    var 出 = [], 未測定 = [];
    var 数 = { できる: 0, かえる: 0, 'べつの道': 0, わからない: 0 };
    for (var i = 0; i < (モジュール || []).length; i++) {
      var m = モジュール[i];
      if (!m || !m.確か) { 未測定.push({ 名: (m && m.名) || '', なぜ: (m && m.なぜ) || '正しく読めたと 言えません' }); continue; }
      var たち = 手続きに切る(m.中身);
      for (var j = 0; j < たち.length; j++) {
        var v = 見立てる1本(たち[j]);
        v.モジュール = m.名;
        出.push(v);
        数[v.可否] = (数[v.可否] || 0) + 1;
      }
    }
    return {
      手続き: 出,
      数: 数,
      本数: 出.length,
      未測定: 未測定,
    };
  }

  /* ══ ⑤AIに渡す物＝★要約だけ★（中身は 渡さない） ═══════════
     ★決まり（指示役 2026-08-28）＝AIには 要約だけ★
       ・コードの本文は 1文字も 入れない（客の中身・会社名が 入っているため）
       ・入れるのは 名前・分類・可否・行数・手掛かりの言葉 だけ */
  function AIに渡す形(見立て) {
    var 出 = [];
    var たち = (見立て && 見立て.手続き) || [];
    for (var i = 0; i < たち.length; i++) {
      var v = たち[i];
      出.push({
        名: v.名, 種類: v.種類, 行数: v.行数, 可否: v.可否,
        分類: v.分類.map(function (a) { return a.名; }),
        手掛かり: v.手掛かり,
      });
    }
    return { 本数: (見立て && 見立て.本数) || 0, 数: (見立て && 見立て.数) || {}, 手続き: 出 };
  }

  /* ══ ⑥客に見せる 1行 ══════════════════════════════════ */
  function 知らせの字(見立て) {
    if (!見立て || !見立て.本数) {
      if (見立て && 見立て.未測定 && 見立て.未測定.length) {
        return 'マクロは ありますが、正しく読めたと 言えない物が ' + 見立て.未測定.length + '本 あります。';
      }
      return 'マクロは 見つかりませんでした。';
    }
    var d = 見立て.数;
    var s = 'マクロが ' + 見立て.本数 + '本 あります。';
    var 内 = [];
    if (d.できる) 内.push('そのまま できる ' + d.できる + '本');
    if (d.かえる) 内.push('やり方を かえる ' + d.かえる + '本');
    if (d['べつの道']) 内.push('別の道が いる ' + d['べつの道'] + '本');
    if (d.わからない) 内.push('読み取れない ' + d.わからない + '本');
    if (内.length) s += '（' + 内.join('／') + '）';
    if (見立て.未測定.length) s += ' 正しく読めなかった物が ' + 見立て.未測定.length + '本 あります。';
    return s;
  }

  var api = {
    手続きに切る: 手続きに切る,
    見立てる1本: 見立てる1本,
    見立てる: 見立てる,
    AIに渡す形: AIに渡す形,
    知らせの字: 知らせの字,
    分け方: 分け方,
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.VbaMikata = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
