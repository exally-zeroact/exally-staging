# ★★SheetJS（`lib/xlsx.full.min.js`）── ★どこから 取ったか・免状★★★（2026-09-18）

> ★★repo に 免状の 紙が 0本でした★★（`git ls-files` で 数えた）
> ★Apache-2.0 は ★配る 時に 免状の 写しを 付ける★ 事を 求めます★
> ⇒★隣に 置きました★ … `lib/xlsx-0.20.3-LICENSE.txt`
> ★★私は 法の 判じを して いません★★＝★実測と 免状の 字を 出すだけです★

---

## ①★★手元の 物（★自分で 測りました★）★★

```
  場所 ………… `lib/xlsx.full.min.js`
  大きさ ……… ★951,904 バイト★
  sha256 ……… ★cc015130aa8521e7f088f88898eba949ccdcbfb38df0bd129b44b7273c3a6f41★
  版 …………… ★0.20.3★（ファイルの 中の `version="0.20.3"`）
  頭の 1行 …… `/*! xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */`
  ★★ファイルの 中に 免状の 名前は 書いて ありません★★
    `Apache` の 字 … ★0件★／`license` の 字 … ★0件★
```

---

## ②★★配り元（★読むだけ・こちらから 何も 出して いません★）★★

```
  ★免状の 紙★
    https://cdn.sheetjs.com/xlsx-0.20.3/package/LICENSE
    ★http 200 ／ 11,355 バイト★
    頭 … ★Apache License Version 2.0, January 2004★
    sha256 … ★4d2a38ac35cda06a555c84074a819d413339cd3691b822cae50f8f322fe01f64★
  ★台帳★
    https://cdn.sheetjs.com/xlsx-0.20.3/package/package.json
    `"version": "0.20.3"` ／ ★`"license": "Apache-2.0"`★
```

---

## ③★★npm には 在りません（★これが「紙が 無かった」訳★）★★

```
  ★npm の 台帳（`registry.npmjs.org/xlsx`）★ … ★一番 新しいのは 0.18.5★
  ⇒★★0.20.3 は npm に 在りません★★
  ⇒★だから `package.json` の 依存にも 書かれて いない／`node_modules` にも 無い★
  ⇒★★手で 置いた 物です★★
  ★これは 経営者1 が 測りました★（★私は 配り元の 方を 自分で 測りました★）
```

---

## ④★★判じ（★借り物外しの 仕事は HyperFormula 1本★）★★

```
  ★SheetJS ＝ Apache-2.0★ … ★GPLv3 と 同じ 話では ありません★
  ⇒★★外さなくて よい★★（★記憶の 決め「仕事は 免状ごとに 切る」の 通り★）
  ★HyperFormula★ … ★GPLv3★／頭に `Copyright (c) HANDSONCODE`
    ⇒★★こちらが 外す 物★★／★免状の 紙は 追いません（外れたら 消えます）★
```

---

## ⑤★★測って いない 事★★

```
  ・★この 免状が 法として どう 効くか★ … ★私は 判じません★
  ・★配り元の 紙が 明日 変わらないか★ … ★取った 日と sha256 を 上に 書いて あります★
  ・★`lib/` の 他の 借り物★ … ★HyperFormula 以外に 在るか まだ 数えて いません★
```
