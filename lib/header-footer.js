/* header-footer.js — ★印刷の ヘッダーと フッター（実Excel と 同じ 印）★ 2026-08-31
 *
 *  ★司さんの 指示（2026-08-31）★
 *    「ヘッダーフッターは ★ごちゃごちゃに ならないよう ドロップダウンでも
 *      いいから 綺麗に★ する」
 *
 *  ★★印の 意味は 実Excel に 刷らせて 測った★★（2026-08-31）
 *    新規の 空ブックに 120行 入れて、ヘッダー/フッターに 印を 入れ、
 *    ★Excel 自身に PDF を 書き出させて★ 中の 字を 読んだ。
 *    ＝★推測は 1つも 入っていない★。
 *
 *      印    入れた物            ★刷って 出た 字（実測）★
 *      &F    L=&F                L=hf-probe        … ファイル名（拡張子なし）
 *      &P    C=&P / &N           C=1 / 4           … ページ番号
 *      &N    〃                  〃                 … ページ数
 *      &A    R=&A                R=シート名テスト   … シート名
 *      &D    LF=&D               LF=2026/8/31      … 日付
 *      &T    CF=&T               CF=0:44           … 時刻
 *      &Z    RF=&Z               RF=C:\Users\...   … 置き場所（フォルダ）
 *
 *    ★入れ物の 形も 実測★（Excel が 保存した xlsx の中）
 *      <headerFooter><oddHeader>&L左&C中央&R右</oddHeader>
 *                    <oddFooter>&L左&C中央&R右</oddFooter></headerFooter>
 *      <pageMargins ... header="0.3" footer="0.3"/>   ＝21.6pt（COMでも 同じ値）
 *
 *    ★&& は 「&」そのもの★（Excel が そのまま 返した）
 *
 *  ★うちで 勝手に 決めた 物★（隠さない）
 *    ・★「よく使う形」の 一覧★は ★うちの 物★。
 *      実Excel の ヘッダーの 一覧は ★UIAから 起こせなかった★（4通り 試した）ので
 *      ★見た目は 写していない★。中身（印）だけ 同じに してある。
 *
 *  見張り: tests/header-footer.test.mjs
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.HeaderFooter = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* ★印の 表★（実Excel に 刷らせて 測った 物だけ） */
  var 印 = [
    { 印: '&P', 名: 'ページ番号',   例: '1' },
    { 印: '&N', 名: 'ページ数',     例: '4' },
    { 印: '&D', 名: '日付',         例: '2026/8/31' },
    { 印: '&T', 名: '時刻',         例: '0:44' },
    { 印: '&F', 名: 'ファイル名',   例: 'hf-probe' },
    { 印: '&A', 名: 'シート名',     例: 'シート名テスト' },
    { 印: '&Z', 名: '置き場所',     例: 'C:\\Users\\...' },
    { 印: '&&', 名: '「&」そのもの', 例: '&' }
  ];

  /* ★よく使う形★＝★うちの 一覧★（実Excel の 一覧は 起こせなかった＝写していない） */
  var よく使う形 = [
    { 名: '（なし）',                    左: '', 中: '', 右: '' },
    { 名: 'ページ 1',                    左: '', 中: 'ページ &P', 右: '' },
    { 名: 'ページ 1 / 4',                左: '', 中: 'ページ &P / &N', 右: '' },
    { 名: 'シート名',                    左: '', 中: '&A', 右: '' },
    { 名: 'シート名 ／ ページ 1',        左: '&A', 中: '', 右: 'ページ &P' },
    { 名: 'ファイル名 ／ ページ 1 / 4',  左: '&F', 中: '', 右: 'ページ &P / &N' },
    { 名: '日付 ／ ページ 1',            左: '&D', 中: '', 右: 'ページ &P' },
    { 名: '日付 ／ シート名 ／ ページ 1', 左: '&D', 中: '&A', 右: 'ページ &P' },
    { 名: 'ファイル名（置き場所つき）',  左: '&Z&F', 中: '', 右: '' }
  ];

  function 空か(v) { return !v || !String(v).trim(); }

  /** 左/中央/右 → Excel の 1本の 字（&L…&C…&R…）
   *  ★空の 所は 書かない★（Excel も 空の &L を 書かない） */
  function 組み立てる(o) {
    o = o || {};
    var s = '';
    if (!空か(o.左)) s += '&L' + o.左;
    if (!空か(o.中)) s += '&C' + o.中;
    if (!空か(o.右)) s += '&R' + o.右;
    return s;
  }

  /** Excel の 1本の 字 → 左/中央/右
   *  ★&& は 「&」そのもの★なので ★先に よけてから★ &L/&C/&R を 探す
   *  （でないと 「A&&L B」 の &&L を 区切りと 読み違える） */
  function 解く(s) {
    var 出 = { 左: '', 中: '', 右: '' };
    if (!s) return 出;
    var 今 = '左';
    var i = 0;
    var t = String(s);
    while (i < t.length) {
      if (t.charAt(i) === '&' && i + 1 < t.length) {
        var 次 = t.charAt(i + 1);
        if (次 === '&') { 出[今] += '&&'; i += 2; continue; }   /* ★よける★ */
        if (次 === 'L') { 今 = '左'; i += 2; continue; }
        if (次 === 'C') { 今 = '中'; i += 2; continue; }
        if (次 === 'R') { 今 = '右'; i += 2; continue; }
      }
      出[今] += t.charAt(i);
      i++;
    }
    return 出;
  }

  /** 印を 本物の 字に 差し替える
   *  @param 文  '&A ページ &P / &N'
   *  @param 今  { ページ, ページ数, ファイル名, シート名, 置き場所, 日付, 時刻 }
   *  ★&& を 先に よける★＝「&&P」は 「&P」（ページ番号では ない）
   */
  function 差し込む(文, 今) {
    if (!文) return '';
    今 = 今 || {};
    var 表 = {
      P: String(今.ページ === undefined ? 1 : 今.ページ),
      N: String(今.ページ数 === undefined ? 1 : 今.ページ数),
      D: String(今.日付 === undefined ? '' : 今.日付),
      T: String(今.時刻 === undefined ? '' : 今.時刻),
      F: String(今.ファイル名 === undefined ? '' : 今.ファイル名),
      A: String(今.シート名 === undefined ? '' : 今.シート名),
      Z: String(今.置き場所 === undefined ? '' : 今.置き場所)
    };
    var 出 = '';
    var i = 0, t = String(文);
    while (i < t.length) {
      if (t.charAt(i) === '&' && i + 1 < t.length) {
        var c = t.charAt(i + 1);
        if (c === '&') { 出 += '&'; i += 2; continue; }
        if (表[c] !== undefined) { 出 += 表[c]; i += 2; continue; }
      }
      出 += t.charAt(i);
      i++;
    }
    return 出;
  }

  /* ★ページ番号だけ 後回しに する時の 目印★
     紙は ★1枚ずつ 番号が 違う★ので、他の 印（日付・シート名…）を 先に 入れて
     ★&P と &N だけ 残す★。残した 所を 刷る 窓が 数字に 差し替える。
     ★2026-08-31 実ブラウザで 実測★＝
       ・position:fixed は ★毎ページ 出る★（全ページ 上端 y=11）
       ・★CSS の counter(page) は 増えない★（10ページ目も「1」のまま）
       ⇒ だから ★自分で ページを 割って 番号を 書く★しか ない。 */
  var 頁印 = 'P';
  var 総印 = 'N';

  /** ページ番号だけ 残して 差し込む（他の 印は 本物の 字に する） */
  function 頁を後回しで差し込む(文, 今) {
    if (!文) return '';
    var 出 = '';
    var i = 0, t = String(文);
    while (i < t.length) {
      if (t.charAt(i) === '&' && i + 1 < t.length) {
        var c = t.charAt(i + 1);
        if (c === '&') { 出 += '&&'; i += 2; continue; }   /* ★まだ 1つに しない★ */
        if (c === 'P') { 出 += 頁印; i += 2; continue; }
        if (c === 'N') { 出 += 総印; i += 2; continue; }
      }
      出 += t.charAt(i);
      i++;
    }
    return 差し込む(出, 今);
  }

  /** 残した 目印を 本物の 数字に する（刷る 窓が 使う） */
  function 頁を入れる(文, ページ, ページ数) {
    return String(文 === undefined || 文 === null ? '' : 文)
      .split(頁印).join(String(ページ))
      .split(総印).join(String(ページ数));
  }

  /** 使っているか（1つも 書いていなければ 刷る時に 場所を 取らない） */
  function 使っているか(o) {
    if (!o) return false;
    return !空か(o.左) || !空か(o.中) || !空か(o.右);
  }

  /** 数える（見張り用） */
  function 数える() {
    return { 印: 印.length, よく使う形: よく使う形.length };
  }

  return {
    印: 印, よく使う形: よく使う形, 頁印: 頁印, 総印: 総印,
    組み立てる: 組み立てる, 解く: 解く, 差し込む: 差し込む,
    頁を後回しで差し込む: 頁を後回しで差し込む, 頁を入れる: 頁を入れる,
    使っているか: 使っているか, 数える: 数える
  };
}));
