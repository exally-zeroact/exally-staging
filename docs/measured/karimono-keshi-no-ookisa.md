# ★借り物（HyperFormula）を 消すのに 何が 要るか★（2026-09-14 実測）

> ★どの 版を 測ったか★ … `hyperformula.full.min.js` ／ ★版 3.4.0★ ／ ビルド 10/08/2026 18:09:27
> ★944,675バイト・GPLv3★（licenseKey は gpl-v3 で 使って います）
> ★版を 必ず 書く★＝版が 上がると 関数の 数が 変わります
>   （記憶の「エンジン 394個」は ★2.6.1 の 時の 数★・実Excel 507個 の 時点。★上書きしないで ください★）

## ★名簿の 出どころ★（手で 選んで いません）
  借り物 … `hyperformula.full.min.js` を そのまま 読み込んで getRegisteredFunctionNames
  実Excel … `docs/measured/excel-functions-2026-09-06.txt`（519個）
  うち … `lib/shiki-kansuu.js` の 表 ＋ 自前の 皮 `shiki-tsunagi.js`

## ★数★
```
  借り物が 素で 出す ………………………………… 423個
    ★うち 実Excel にも 在る★ …………………… 384個
    ★実Excel に 無い（要らない）★ …………… 39個
        = ARRAYFORMULA ARRAY_CONSTRAIN CHIDISTRT CHIINVRT COUNTUNIQUE COVARIANCEP COVARIANCES FDISTRT FINVRT HF.ADD HF.CONCAT HF.DIVIDE HF.EQ HF.GT HF.GTE HF.LT HF.LTE HF.MINUS HF.MULTIPLY HF.NE HF.POW HF.UMINUS HF.UNARY_PERCENT HF.UPLUS INTERVAL ISBINARY LOGNORMINV MAXPOOL MEDIANPOOL POISSONDIST SKEWP SPLIT STDEVS TDIST2T TDISTRT TINV2T VARS VERSION WEIBULLDIST

  ★★借り物を 消す のに 自分で 書く 要が 在る … 377個★★
      ★実Excel の 実測が 紙に 在る … 332個★
      ★実測が まだ 無い ………………… 45個★

  （参考）実Excel に 在って ★借り物も うちも 出せない★ … 74個
```

## ★★↑この 上の 「377」と 「74」は 古い（2026-09-15 に 押し直しました）★★

```
  ★何が 抜けて いたか★ … ★プラグインと JS層が 足す 分★
  ★★段ごとに 数えました（重なりを 除いて／台帳 519 と 突き合わせ）★★
      ★①借り物だけ★                         ★384個★
      ★②＋`registerExallyFunctions`★        ★418個★（★+34★）
          CONCAT LOOKUP XMATCH INDIRECT AGGREGATE RANK系 …
      ★③＋`lib/formula-*-plug.js` 8本★      ★483個★（★+65★）
          extra +13 ／ nokori +14 ／ kane +22 ／ yosoku +10
          soto +4 ／ filterxml +1 ／ cell +1 ／ complex ★+0★
      ★④＋JS層 `_jsSet`★                    ★489個★（★+6★）
          FREQUENCY REDUCE SCAN MAP MAKEARRAY ISOMITTED
      ★消えた 名前★ ★0個★（プラグインが 借り物を 潰しては いません）
  ★★⇒ 名乗る 489個 ／ 名乗らない 30個★★（384+34+65+6 ＝ 489）
     ★★名乗らない 30個（全部）★★
       AREAS BINOM.DIST.RANGE BYCOL BYROW CALL CUBEKPIMEMBER CUBEMEMBER
       CUBEMEMBERPROPERTY CUBERANKEDMEMBER CUBESET CUBESETCOUNT CUBEVALUE
       FORECAST.ETS FORECAST.ETS.CONFINT FORECAST.ETS.SEASONALITY FORECAST.ETS.STAT
       GETPIVOTDATA GROUPBY IMAGE LAMBDA LET ODDFPRICE ODDFYIELD ODDLPRICE
       ODDLYIELD PHONETIC PIVOTBY REGISTER.ID RTD TRIMRANGE
     ★内訳★ … 09-07の 紙「動かない」27個 ＋ ★その 紙に 無い 3個★
               （★AREAS BINOM.DIST.RANGE LET★）
     ⇒★★`docs/measured/exally-missing-2026-09-07.txt` は 3個 足りません★★
       （★私は その 紙を 信じて 母数を 492 と 出しました＝間違い★）
```

## ★★この 数の 読み方（★3つ とも 大事★）★★

```
  ★①「名乗る」は「答える」では ありません★
     ＝これは `getRegisteredFunctionNames` の 数＝★名前の 数★
     ＝★519本 押した 数では ありません★
     ★裏づけの 在る 分★ … プラグイン 86個は
       `docs/measured/golden-86-karimono-2026-09-15.tsv` で ★2,257行 押して あります★
  ★★②「自前の 台が 知る 76個」を ここに 足さないで ください★★
     ＝`book.html` の `<script src=…>` に ★`shiki-*` は 0本★（実測・grep）
     ⇒★★お客さんの 画面では 自前の 台は 0%★★
     ＝★「書いた 数」と「客に 届いて いる 数」を 同じ 行に 足さない★
  ★★③「自分で 書く 数」は もっと 多い★★
     `exally-formula.js` の `function _js…(` ★79本★の うち
     ★★8本が `_hf`（借り物の 表）を 読みます★★
       `_jsDbFunc`（D系 12個ぶん）／`_jsOffset`／`_jsIndirect`
       `_jsReduceCompute`／`_jsScanCompute`／`_jsMapCompute`
       `_jsMakearrayCompute`／`_jsComputeFormula`
     ⇒★名前の 上では「自前」でも ★借り物を 外したら 動きません★★
     ＝★IM系で 見つけたのと 同じ 型★（[[feedback_jimae_ka_wa_meibo_de_naku_dare_ga_keisan_suru_ka]]）
  ★★数え方の 穴（2つとも 私が 踏みました）★★
     ㋐`getRegisteredFunctionNames` は ★同じ 名前を 2回 返す★
        ⇒★重なりを 除かないと 431・489 が 合わなく なります★
     ㋑`registerExallyFunctions` は ★プラグインより 前★に 走る
        ⇒★「プラグインの 前」を `HF0` で 取ると ★既に +34 済み★★
```

## ★取り方（★次に 数える 人は これを そのまま★）★

```
  ① `hyperformula.full.min.js` を 読む                      → ★384★
  ② `exally-formula.js` の `registerExallyFunctions(HFns)`  → ★418★
  ③ `lib/formula-*-plug.js` 8本の ★`つなぐ(H, 中身)` を 呼ぶ★ → ★483★
     ★★読み込むだけでは 登録されません★★（＝`つなぐ` を 呼ぶまで 0個）
  ④ `var _jsSet` の 名前を 足す                             → ★489★
  ⑤ `docs/measured/excel-functions-2026-09-06.txt`（519行）と 突き合わせ
  ⑥ ★重なりを 除く★／★「名乗る」と 書く（「答える」と 書かない）★
  ⑦ ★★直す 順は「製品を 直す → 最後に 焼く」★★
     ＝`exally-formula.js` も `lib/*` も ★材料の 指紋★に 入って います
     ⇒★先に 焼いて から 製品を 直すと ★焼き直しを もう1回★★
     ★2026-09-15 に 私が 実際に 1回 損しました★（焼く → `_jsDbFunc` を 直す → また 焼く）
     ＝★記憶の「総なめは 最後の 直しの 後に」と 同じ 形★

  ★★この 道具が 在ります★★ … `docs/measured/kazoeru-seihin-ga-kotaeru.mjs`
     `node docs/measured/kazoeru-seihin-ga-kotaeru.mjs` で ①〜⑥を そのまま 出します
     ★試験には して いません★＝★CI で 毎回 走らせる 物では ない★
       （★数が 動くのは 製品を 触った 時だけ★／★毎回 走らせると 借り物を 毎回 建てる★）
     ⇒★数を 紙に 書く 時は ★この 道具の 出しを 貼る★★
```

## ★別枠＝OFFSET（1個）★

  ★OFFSET は 計算する 関数では ありません★＝★式を 組み立て直す 仕掛け★です。
  借り物の 中では `ParsingErrorType.StaticOffsetError`＝★式を 読む 段階★で 扱われ、
  名簿（implementedFunctions）には 出て 来ません（★動かすと 名前としては 出る★）。
  ⇒★普通の 関数と 同じ 組に 入れると 必ず 詰まります★＝soto 4個と 同じ「土台の 仕事」。
  （2026-09-14・経営者1 が 束ねた 1本の 中を 読んで 見つけた）

## ★実測が まだ 無い 45個（先に 測る 要が 在る）★
  BINOM.DIST BINOM.INV BINOMDIST COUNTIF CRITBINOM CUMIPMT CUMPRINC DATEDIF
  DATEVALUE DAVERAGE DAY DCOUNT DCOUNTA DGET DMAX DMIN
  DPRODUCT DSTDEV DSTDEVP DSUM DVAR DVARP EOMONTH FILTER
  HLOOKUP IFERROR IFNA IFS MATCH MOD NEGBINOM.DIST NEGBINOMDIST
  NOT PERCENTILE PERCENTILE.EXC QUARTILE QUARTILE.EXC RATE ROUNDDOWN SUBTOTAL
  T.TEST TEXTJOIN TTEST VLOOKUP XLOOKUP

## ★作り直し方★

```
  node karimono-keshi-no-ookisa.mjs   （scratchpad の 道具／名簿は 全部 機械が 取る）
```

> ★この 紙は 機械が 出しました★。手で 名前を 並べて いません。
