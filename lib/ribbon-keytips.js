/* ribbon-keytips.js — ★Alt を 押した 時の「押す順」の 木★ 2026-08-31
 *
 *  ★なぜ 要るか（司さん 2026-08-30）★
 *    「Excel を 細胞分解レベルまで 網羅して 把握した上で 持ち込み パクる」
 *    ★毎日 Excel を 使う人ほど Alt の 順番で 打つ★（Alt,H,1＝太字／Alt,H,A,L＝左揃え）
 *    実測＝実Excel 462個 vs うち ★0個★ だった。
 *
 *  ★勝手に 決めない★
 *    キーは lib/ribbon-keys.js（実Excel を UI Automation で 測った 物）だけを 使う。
 *    ★実Excel が 持っていない キーは 作らない★（＝32個は 空のまま。捏造しない）
 *
 *  ★ここは 画面を 触らない★＝木を 作って「次に 何が 出来るか」を 答えるだけ。
 *    画面（札を 描く・キーを 受ける）は lib/ribbon.js 側。
 *    ⇒ ★node で そのまま 試験できる★（見張り tests/ribbon-keytips.test.mjs）
 *
 *  ★木の 形★
 *    { 子: { 'H': { タブ:'ホーム', 子:{ '1':{ 動作:'太字' }, 'A':{ 子:{'L':{動作:'左揃え'}} } } } } }
 *    ・タブ … その字まで 打ったら ★タブを 切り替える★（Excelと同じ。打ち続けられる）
 *    ・動作 … その字まで 打ったら ★押す★（RibbonActions の 名前）
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.RibbonKeytips = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var 尻 = function (s) { return String(s).replace(/\.\.\.$/, '').replace(/…$/, '').trim(); };

  function 空の節() { return { 子: {} }; }

  /** 木を 作る
   *  @param 鍵  RibbonKeys（タブ／部品の鍵）
   *  @param 品  RibbonSpec.ITEMS（つないだ物だけ 拾う）
   *  @param 外  RibbonScope（分母の外は 入れない）※無くてもよい
   */
  function 作る(鍵, 品, 外) {
    var 根 = 空の節();
    if (!鍵 || !品) return 根;

    function 掘る(順) {
      var 今 = 根;
      for (var i = 0; i < 順.length; i++) {
        var c = String(順[i]).toUpperCase();
        if (!今.子[c]) 今.子[c] = 空の節();
        今 = 今.子[c];
      }
      return 今;
    }

    /* ★タブが 先★（同じ字を 部品が 使っても タブの 印は 消さない） */
    var タブ = 鍵.タブ || {};
    for (var t in タブ) {
      if (!Object.prototype.hasOwnProperty.call(タブ, t)) continue;
      掘る(タブ[t]).タブ = t;
    }

    /* ★部品★＝★つないだ物（it.a が 在る）だけ★
       ＝押せない物に キーを 付けると「押したのに 何も 起きない」に なる */
    var 候補 = {};                                 /* 順の文字 → このキーを 名乗る 物たち */
    var 飛ばした = 0;
    for (var j = 0; j < 品.length; j++) {
      var it = 品[j];
      if (!it || !it.a) { 飛ばした++; continue; }
      if (外 && 外.分母の外か && 外.分母の外か(it.t, it.g, it.p)) { 飛ばした++; continue; }
      var 順 = 鍵.部品の鍵 ? 鍵.部品の鍵(it.t, 尻(it.p)) : null;
      if (!順 || !順.length) { 飛ばした++; continue; }
      var 鍵字 = 順.join(',').toUpperCase();
      if (!候補[鍵字]) 候補[鍵字] = { 順: 順, 物: [] };
      候補[鍵字].物.push(it);
    }

    /* ★同じ キーを ★違う名前★の 物が 名乗ったら ★付けない★★
       実例（実測 2026-08-31）：描画の ペン 10個は ★実 Excel 自体が★
       全部に 同じ `Alt, J I, G` を 付けている（＝画廊の キー。
       実 Excel では その後 矢印で 選ぶ）。
       ★ここで 1つを 選ぶと「黒のペンのつもりが 消しゴム」に なる★
       ⇒ ★画廊を 作るまで 付けない★（曖昧に 残して 数える） */
    var 入れた = 0;
    根.曖昧 = [];
    for (var 鍵字b in 候補) {
      if (!Object.prototype.hasOwnProperty.call(候補, 鍵字b)) continue;
      var 組 = 候補[鍵字b];
      var 名たち = {}, 数 = 0;
      for (var y = 0; y < 組.物.length; y++) {
        var 名 = 尻(組.物[y].p);
        if (!名たち[名]) { 名たち[名] = 1; 数++; }
      }
      if (数 > 1) {
        根.曖昧.push({ 鍵: 鍵字b, 数: 数,
          物: 組.物.map(function (v) { return v.t + '|' + v.p; }) });
        飛ばした += 組.物.length;
        continue;
      }
      /* 名前が 同じなら ★リボンの 並び順で 先頭★を 採る（同じ 場所の 分かれボタン） */
      var 節 = 掘る(組.順);
      var 頭 = 組.物[0];
      節.動作 = 頭.a.act; 節.名 = 頭.p; 節.タブ元 = 頭.t;
      入れた++;
      飛ばした += (組.物.length - 1);
    }
    根.入れた = 入れた;
    根.飛ばした = 飛ばした;
    return 根;
  }

  /** 1字 進む
   *  @return {節, どうする} … どうする＝'進む'|'タブ'|'押す'|'外れ'
   *    ★タブの 節でも 子が 居れば 打ち続けられる★（Excelと 同じ）
   */
  function 進む(節, 字) {
    if (!節) return { 節: null, どうする: '外れ' };
    var c = String(字 || '').toUpperCase();
    var 次 = 節.子 ? 節.子[c] : null;
    if (!次) return { 節: null, どうする: '外れ' };
    if (次.動作) return { 節: 次, どうする: '押す' };
    if (次.タブ) return { 節: 次, どうする: 'タブ' };
    return { 節: 次, どうする: '進む' };
  }

  /** その 節から 次に 打てる 字（画面に 札を 出す為） */
  function 次の字(節) {
    if (!節 || !節.子) return [];
    return Object.keys(節.子).sort();
  }

  /** 数える（見張り用） */
  function 数える(根) {
    var タブ = 0, 動作 = 0, 深さ = 0;
    (function 潜る(n, d) {
      if (!n) return;
      if (n.タブ) タブ++;
      if (n.動作) 動作++;
      if (d > 深さ) 深さ = d;
      for (var c in n.子) if (Object.prototype.hasOwnProperty.call(n.子, c)) 潜る(n.子[c], d + 1);
    }(根, 0));
    return { タブ: タブ, 動作: 動作, 深さ: 深さ };
  }

  return { 作る: 作る, 進む: 進む, 次の字: 次の字, 数える: 数える };
}));
