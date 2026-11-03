/* formula-nokori.js — ★実Excel に 在って まだ 動かない 関数★（2026-09-06）
 *
 *  ★司さん★「きっちり ファイルに するのと ★細胞レベルで 対応できるように しろ★」
 *
 *  ★★どこから 来た 一覧か★★
 *    `docs/measured/exally-missing-2026-09-06.txt`（★83個★）
 *    ＝一次情報の 関数一覧 521個 → 実Excel に 1つずつ 打たせて 519個 →
 *      ★本番の 道（JS層 → convertFormula → engine）で 押して 動かなかった 83個★
 *    ⇒ ★手で 選んでいません★（測って 出た 物）
 *
 *  ★★答えは 全部 実Excel に 打たせて 取った★★
 *    Excel 16.0 build 20326（日本語UI 1041）／答えは
 *    `docs/measured/golden-2026-09-06.tsv` に 1行ずつ 置いた。
 *    ⇒ ★私が 考えた 値は 1つも 入っていない★
 *
 *  ★ここは 純粋な 計算だけ★（エンジンに 繋ぐのは lib/formula-nokori-plug.js）
 *    ＝node で そのまま 試験できる（[[feedback_tests_must_run_not_just_read_source]]）
 *
 *  見張り: tests/formula-nokori.test.mjs
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.FormulaNokori = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  function 表にする(v) {
    if (!Array.isArray(v)) return [[v]];
    if (!v.length) return [[]];
    if (!Array.isArray(v[0])) return [v.slice()];
    return v.map(function (r) { return r.slice(); });
  }
  function 平らに(v) {
    var 出 = [], A = 表にする(v);
    for (var r = 0; r < A.length; r++) for (var c = 0; c < A[r].length; c++) 出.push(A[r][c]);
    return 出;
  }
  function 数だけ(v) {
    return 平らに(v).map(Number).filter(function (x) { return isFinite(x); });
  }

  /* ══ ①字の 幅で 数える 3つ（B が 付く 物）══════════════════
     ★実測（実Excel 日本語UI）★
       FINDB("い","あいう") ＝ 3   … ★全角は 2バイト★
       SEARCHB("い","あいう") ＝ 3
       REPLACEB("あいう",3,2,"X") ＝ あXう
     ⇒★1文字＝1か2★。半角は 1・全角は 2。 */
  function バイト幅(ch) {
    var c = ch.charCodeAt(0);
    /* ★半角＝ASCII と 半角カナ★（実Excel の LENB と 同じ数え方） */
    if (c <= 0x7F) return 1;
    if (c >= 0xFF61 && c <= 0xFF9F) return 1;
    return 2;
  }
  function バイト数(s) {
    var n = 0, t = String(s);
    for (var i = 0; i < t.length; i++) n += バイト幅(t.charAt(i));
    return n;
  }
  /** ★バイトの 位置 → 文字の 位置★（1から 数える） */
  function バイト位置を文字位置に(s, b) {
    var t = String(s), n = 0;
    for (var i = 0; i < t.length; i++) {
      n += バイト幅(t.charAt(i));
      if (n >= b) return i + 1;
    }
    return t.length + 1;
  }
  /* ══ ★何文字目から／何文字ぶん の 見張り★（2026-09-08 に 足した）══════
     ★実Excel に 180本 打って 決めた★
       紙 … `docs/measured/kansuu46/golden-hikisuu-mihari-2026-09-08.tsv`
       取り方 … `docs/measured/kansuu46/toru-hikisuu-mihari.ps1`
     ★どこから 断るかは 私が 決めていません＝実Excel の 答えの 通り★

       =FINDB("b","abc",n)          n=-2 -1 -0.5 0 0.5 … ★#VALUE!★ ／ n=1 1.5 2 2.5 … 2
       =REPLACEB("abcde",n,1,"x")   n=-2 … 0.5 … ★#VALUE!★ ／ n=32767 … abcdex ／ ★n=32768 … #VALUE!★
       =REPLACEB("abcde",1,n,"x")   n=-2 … -0.5 … ★#VALUE!★ ／ n=0 … xabcde ／ ★n=32768 … #VALUE!★

     ★★32,767 は 2の15乗−1＝実Excel の 1マスの 字数の 上限★★
       ⇒ 2分探索を 2回 走らせて 2回とも 同じ ⇒★丸い 数＝仕様＝写して よい★
       （MUNIT の 境目 7,326 は 丸くなかった＝その 機械の 記憶容量＝写さない、と 分けた）

     ★★『小数は 切り捨て』＝ここが 一番 危ない★★
       うちは `Math.max(1, Number(開始) || 1)` で 渡して いたので
       `バイト位置を文字位置に` の `n >= 2.5` が ★2.5 を 3文字目★に していた。
       実Excel は ★切り捨てて 2文字目★。
       ⇒ 誤りでは なく ★静かに 違う 答え★が 出ていた
         =REPLACEB("abcde",1.5,1,"x") … 実Excel ★xbcde★ ／ うち ★axcde★
       ⇒★★『断らない』より『違う 答えを 返す』方が 見つかりません★★

     ★`Math.max(1, …)` が やっていた 事★ … 0 も −1 も ★黙って 1 に 直して いた★
       ⇒ 実Excel は そこで #VALUE! と 言う ⇒★黙って 直すのを やめた★ */
  var 字数の上限 = 32767;   /* ★実Excel の 1マスの 字数（2の15乗−1）★ */

  /** ★何文字目から／何文字ぶん★ を 見張って 整数に する（駄目なら null）
   *
   *  ★★下限は ★切り捨てる 前★に 見る★★（2026-09-08 実測）
   *    `=REPLACEB("abcde",1,-0.5,"x")` … 実Excel ★#VALUE!★
   *    ⇒ `Math.trunc(-0.5)` は ★マイナスの 0★＝`-0 < 0` が false で ★すり抜ける★
   *    ⇒ 先に 切り捨てると −0.5 が 0 に なって ★xabcde を 返して しまった★
   *  ★TOCOL の 2つ目は 逆で、−0.5 を ★実Excel が 受ける★（切り捨てて 0）★
   *    ⇒★★同じ「小数」でも 関数ごとに 違う＝1か所で まとめて 判定 できない★★
   *      （書式の 付いた 数でも 同じ 事が 数で 出た） */
  function 字の位置を見張る(v, 下限) {
    if (v === null || v === undefined || v === '') return 下限;   /* ★省いた 時＝既定★ */
    var n = Number(v);
    if (!isFinite(n)) return null;
    if (n < 下限) return null;                                    /* ★切り捨てる 前に 見る★ */
    n = Math.trunc(n);                                            /* ★切り捨て★ */
    if (n > 字数の上限) return null;
    return n;
  }

  /** FINDB … ★大文字小文字を 区別する★（FIND と 同じ） */
  function 探すB(探す字, 元, 開始) {
    var s = String(元), f = String(探す字);
    var 始 = 字の位置を見張る(開始, 1);
    if (始 === null) return { 誤り: 'VALUE' };
    var 頭 = バイト位置を文字位置に(s, 始) - 1;
    var i = s.indexOf(f, 頭);
    if (i < 0) return { 誤り: 'VALUE' };
    return バイト数(s.slice(0, i)) + 1;
  }
  /** SEARCHB … ★区別しない★（SEARCH と 同じ）／`*` `?` は ここでは 見ない */
  function 探すB大小なし(探す字, 元, 開始) {
    var s = String(元), f = String(探す字);
    var 始 = 字の位置を見張る(開始, 1);
    if (始 === null) return { 誤り: 'VALUE' };
    var 頭 = バイト位置を文字位置に(s, 始) - 1;
    var i = s.toLowerCase().indexOf(f.toLowerCase(), 頭);
    if (i < 0) return { 誤り: 'VALUE' };
    return バイト数(s.slice(0, i)) + 1;
  }
  /** REPLACEB … ★バイトで 数えて 入れ替える★ */
  function 入れ替えB(元, 開始B, 長さB, 新しい) {
    var s = String(元);
    var 始 = 字の位置を見張る(開始B, 1);      /* ★1から★ */
    var 長 = 字の位置を見張る(長さB, 0);      /* ★0でも よい＝入れるだけ★ */
    if (始 === null || 長 === null) return { 誤り: 'VALUE' };
    var 頭 = バイト位置を文字位置に(s, 始) - 1;
    var 尻 = バイト位置を文字位置に(s, 始 + 長) - 1;
    if (尻 < 頭) 尻 = 頭;
    return s.slice(0, 頭) + String(新しい) + s.slice(尻);
  }

  /* ══ ②区切って 分ける（TEXTSPLIT）══════════════════════════
     ★実測★ TEXTSPLIT("a,b,c",",") ＝ a / b / c（横に 3つ） */
  function 区切って分ける(元, 横の区切り, 縦の区切り) {
    var s = String(元);
    var 縦 = (縦の区切り === undefined || 縦の区切り === null || 縦の区切り === '')
      ? [s] : s.split(String(縦の区切り));
    var 出 = 縦.map(function (行) {
      return (横の区切り === undefined || 横の区切り === null || 横の区切り === '')
        ? [行] : 行.split(String(横の区切り));
    });
    /* ★形を 保つ★＝短い 行は 空で 埋める（2026-08-29 の FILTER の 失敗を 繰り返さない） */
    var 幅 = 出.reduce(function (a, r) { return Math.max(a, r.length); }, 0);
    return 出.map(function (r) {
      var x = r.slice();
      while (x.length < 幅) x.push('');
      return x;
    });
  }

  /* ══ ③正規表現の 3つ（Excel 365 の 新しい 関数）════════════
     ★実測★ REGEXTEST("abc123","[0-9]+") ＝ TRUE
             REGEXEXTRACT("abc123","[0-9]+") ＝ 123
             REGEXREPLACE("abc123","[0-9]+","#") ＝ abc#
     ★大小の 区別★＝既定は ★する★（Excel の 3つ目の 引数 0＝する） */
  function 正規表現を作る(型, 大小を区別しないか, 全部か) {
    var f = (全部か ? 'g' : '') + (大小を区別しないか ? 'i' : '');
    try { return new RegExp(String(型), f); } catch (e) { return null; }
  }
  function 正規で調べる(元, 型, 大小なし) {
    var re = 正規表現を作る(型, 大小なし, false);
    if (!re) return { 誤り: 'VALUE' };
    return re.test(String(元));
  }
  function 正規で取り出す(元, 型, 大小なし) {
    var re = 正規表現を作る(型, 大小なし, false);
    if (!re) return { 誤り: 'VALUE' };
    var m = re.exec(String(元));
    if (!m) return { 誤り: 'NA' };
    return m[0];
  }
  function 正規で入れ替える(元, 型, 新しい, 大小なし) {
    var re = 正規表現を作る(型, 大小なし, true);
    if (!re) return { 誤り: 'VALUE' };
    /* ★$1 などの 置き換えは Excel も 使える★ */
    return String(元).replace(re, String(新しい));
  }

  /* ══ ④並べ替え（SORTBY）════════════════════════════════════
     ★実測★ A=1,4,3,2 ／ B=9,7,8,6
             SORTBY(A1:A4,B1:B4,1) の 1つ目 ＝ 2（Bが 小さい順）
             SORTBY(A1:A4,B1:B4,-1) の 1つ目 ＝ 1（Bが 大きい順） */
  function 別の列で並べる(表, 鍵, 向き) {
    var A = 表にする(表);
    var K = 平らに(鍵);
    var 上 = (Number(向き) === -1) ? -1 : 1;
    var 印 = A.map(function (行, i) { return { 行: 行, 鍵: K[i], i: i }; });
    印.sort(function (a, b) {
      var x = a.鍵, y = b.鍵;
      if (x === y) return a.i - b.i;                 /* ★同じなら 元の 順★ */
      var xn = Number(x), yn = Number(y);
      if (isFinite(xn) && isFinite(yn)) return (xn - yn) * 上;
      return String(x).localeCompare(String(y)) * 上;
    });
    return 印.map(function (p) { return p.行.slice(); });
  }

  /* ══ ⑤行列（MUNIT / MINVERSE）══════════════════════════════
     ★実測★ MUNIT(3) の (1,1)=1 (1,2)=0
             D1:E2 = [[4,7],[2,6]] ⇒ MINVERSE の (1,1)=0.6 (1,2)=-0.7 */
  function 単位行列(n) {
    var N = Math.trunc(Number(n));
    if (!(N >= 1)) return { 誤り: 'VALUE' };
    var 出 = [];
    for (var r = 0; r < N; r++) {
      var 行 = [];
      for (var c = 0; c < N; c++) 行.push(r === c ? 1 : 0);
      出.push(行);
    }
    return 出;
  }
  function 逆行列(表) {
    var A = 表にする(表).map(function (r) { return r.map(Number); });
    var n = A.length;
    if (!n || A[0].length !== n) return { 誤り: 'VALUE' };
    for (var i = 0; i < n; i++) for (var j = 0; j < n; j++) if (!isFinite(A[i][j])) return { 誤り: 'VALUE' };
    /* ガウス・ジョルダン（右に 単位行列を 付けて 掃き出す） */
    var M = A.map(function (行, i) {
      return 行.slice().concat(単位行列(n)[i]);
    });
    for (var k = 0; k < n; k++) {
      var p = k;
      for (var r2 = k + 1; r2 < n; r2++) if (Math.abs(M[r2][k]) > Math.abs(M[p][k])) p = r2;
      if (Math.abs(M[p][k]) < 1e-12) return { 誤り: 'VALUE' };   /* ★逆が 無い＝#VALUE!★（実Excel と 同じ） */
      var t = M[k]; M[k] = M[p]; M[p] = t;
      var d = M[k][k];
      for (var c2 = 0; c2 < 2 * n; c2++) M[k][c2] /= d;
      for (var r3 = 0; r3 < n; r3++) {
        if (r3 === k) continue;
        var f2 = M[r3][k];
        if (!f2) continue;
        for (var c3 = 0; c3 < 2 * n; c3++) M[r3][c3] -= f2 * M[k][c3];
      }
    }
    return M.map(function (行) {
      return 行.slice(n).map(function (x) {
        /* ★丸めの ごみを 落とす★＝0.6000000000000001 を 0.6 に */
        var y = Math.round(x * 1e10) / 1e10;
        return Object.is(y, -0) ? 0 : y;
      });
    });
  }

  /* ══ ⑥順位の 割合（PERCENTRANK.INC / .EXC）══════════════════
     ★実測★ A=1,4,3,2 ／ PERCENTRANK.INC(A1:A4,3) ＝ 0.666
             PERCENTRANK.EXC(A1:A4,3) ＝ 0.6
     ★実Excel は 既定で 3桁に 切り捨てる★（0.6666… → 0.666） */
  function 順位の割合(表, x, 桁, 含むか) {
    var v = 数だけ(表).sort(function (a, b) { return a - b; });
    var n = v.length;
    var X = Number(x);
    if (!n || !isFinite(X)) return { 誤り: 'NA' };
    if (X < v[0] || X > v[n - 1]) return { 誤り: 'NA' };
    var 桁数 = (桁 === undefined || 桁 === null) ? 3 : Math.trunc(Number(桁));
    if (!(桁数 >= 1)) return { 誤り: 'NUM' };
    var 下 = 0;
    while (下 < n && v[下] < X) 下++;
    var 率;
    if (v[下] === X) {
      率 = 含むか ? 下 / (n - 1) : (下 + 1) / (n + 1);
    } else {
      /* ★間に 在る時は 直線で 補う★（実Excel と 同じ） */
      var 小 = v[下 - 1], 大 = v[下];
      var 位置 = (下 - 1) + (X - 小) / (大 - 小);
      率 = 含むか ? 位置 / (n - 1) : (位置 + 1) / (n + 1);
    }
    /* ★切り捨て★（四捨五入では ない＝実Excel の 決まり） */
    var k2 = Math.pow(10, 桁数);
    return Math.floor(率 * k2) / k2;
  }

  /* ══ ⑦確率（PROB）══════════════════════════════════════════
     ★実測★ 値=1,4,3,2 ／ 確率=0.1,0.2,0.3,0.4 ／ PROB(値,確率,2,3) ＝ 0.7 */
  function 確率(値, 確率たち, 下, 上) {
    var v = 平らに(値).map(Number);
    var p = 平らに(確率たち).map(Number);
    if (v.length !== p.length || !v.length) return { 誤り: 'NA' };
    var 和 = 0;
    for (var i = 0; i < p.length; i++) {
      if (!isFinite(p[i]) || p[i] < 0 || p[i] > 1) return { 誤り: 'NUM' };
      和 += p[i];
    }
    if (Math.abs(和 - 1) > 1e-9) return { 誤り: 'NUM' };
    var L = Number(下);
    var U = (上 === undefined || 上 === null) ? L : Number(上);
    var 出 = 0;
    for (var j = 0; j < v.length; j++) if (v[j] >= L && v[j] <= U) 出 += p[j];
    return Math.round(出 * 1e10) / 1e10;
  }

  /* ══ ⑧かたまりの 数（AREAS）════════════════════════════════
     ★実測★ AREAS(A1:B3) ＝ 1 ／ AREAS((A1:B3,D1:D2)) ＝ 2
     ★ここでは「渡された かたまりの 数」を そのまま 数える★
       （式の 書き方を 読むのは 繋ぐ 側の 仕事） */
  function かたまりの数(かたまりたち) {
    if (!Array.isArray(かたまりたち)) return 1;
    return Math.max(1, かたまりたち.length);
  }

  /* ══ ⑨誤りの 番号（ERROR.TYPE）══════════════════════════════
     ★実測★ 1/0 ⇒ 2 ／ NA() ⇒ 7 ／ 誤りでない ⇒ #N/A */
  var 誤りの番号 = {
    NULL: 1, DIV_BY_ZERO: 2, VALUE: 3, REF: 4, NAME: 5, NUM: 6, NA: 7,
    'DIV/0': 2, 'DIV/0!': 2
  };
  function 誤りを番号に(印) {
    var k = String(印 || '').toUpperCase().replace(/^#|[!?]$/g, '');
    if (Object.prototype.hasOwnProperty.call(誤りの番号, k)) return 誤りの番号[k];
    return { 誤り: 'NA' };
  }

  /** 数える（見張り用）★手で 並べない＝ここが 正本★ */
  function 数える() {
    return {
      足す: ['FINDB', 'SEARCHB', 'REPLACEB', 'TEXTSPLIT',
        'REGEXTEST', 'REGEXEXTRACT', 'REGEXREPLACE',
        'SORTBY', 'MUNIT', 'MINVERSE',
        'PERCENTRANK.INC', 'PERCENTRANK.EXC', 'PROB', 'AREAS', 'ERROR.TYPE'],
      /* ★まだ 作っていない 物は ここに 書かない★＝
         `docs/measured/exally-missing-2026-09-06.txt` が 正本（★1つの事は 1か所★） */
      出どころ: 'docs/measured/exally-missing-2026-09-06.txt'
    };
  }

  return {
    バイト数: バイト数, バイト位置を文字位置に: バイト位置を文字位置に,
    探すB: 探すB, 探すB大小なし: 探すB大小なし, 入れ替えB: 入れ替えB,
    区切って分ける: 区切って分ける,
    正規で調べる: 正規で調べる, 正規で取り出す: 正規で取り出す, 正規で入れ替える: 正規で入れ替える,
    別の列で並べる: 別の列で並べる,
    単位行列: 単位行列, 逆行列: 逆行列,
    順位の割合: 順位の割合, 確率: 確率,
    かたまりの数: かたまりの数, 誤りを番号に: 誤りを番号に,
    表にする: 表にする, 数える: 数える
  };
}));
