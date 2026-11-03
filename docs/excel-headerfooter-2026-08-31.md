# ★紙の 上と 下に 入れる 字（ヘッダー／フッター）の 実測★ 2026-08-31

> 司さん：「ヘッダーフッターは ★ごちゃごちゃに ならないよう
> 　ドロップダウンでも いいから 綺麗に★ する」

★推測は 1つも 入っていない★。全部 その日に 測った。

---

## ① ★印の 意味★（実Excel に PDF を 刷らせて 中の 字を 読んだ）

★測り方★
- 新規の 空ブックに 120行 入れる（★司さんの 実物は 1つも 開いていない★）
- ヘッダー／フッターの 6か所に 印を 入れる
- `Worksheet.ExportAsFixedFormat(0, …)` で ★Excel 自身に PDF を 書かせる★
- PDF の 中の 字を 取り出して 読む（★窓は 1つも 開かせていない★）

| 印 | 入れた物 | ★刷って 出た 字★ | 意味 |
|---|---|---|---|
| `&F` | `L=&F` | `L=hf-probe` | ファイル名（★拡張子は 付かない★） |
| `&P` | `C=&P / &N` | `C=1 / 4` | ページ番号 |
| `&N` | 〃 | 〃 | ページ数 |
| `&A` | `R=&A` | `R=シート名テスト` | シート名 |
| `&D` | `LF=&D` | `LF=2026/8/31` | 日付 |
| `&T` | `CF=&T` | `CF=0:44` | 時刻 |
| `&Z` | `RF=&Z` | `RF=C:\Users\…` | 置き場所 |
| `&&` | `&&` | `&&`（Excel が そのまま 返す） | 「&」そのもの |

★`&&` の 罠★ … 「`&&P`」は 「`&P`」であって ★ページ番号では ない★。
ここを 間違えると ★客の 字が 黙って 数字に 化ける★。

---

## ② ★入れ物の 形★（Excel が 保存した xlsx の 中）

```xml
<headerFooter>
  <oddHeader>&amp;LL=&amp;F&amp;CC=&amp;P / &amp;N&amp;RR=&amp;A</oddHeader>
  <oddFooter>&amp;LLF=&amp;D&amp;CCF=&amp;T&amp;RRF=&amp;Z</oddFooter>
</headerFooter>
<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
```

★`header="0.3"` インチ ＝ 21.6pt★（COM の `HeaderMargin` と 一致）

★既定★（新規の 空ブック）
- 6か所 とも ★空★
- `DifferentFirstPageHeaderFooter` = False
- `OddAndEvenPagesHeaderFooter` = False
- `AlignMarginsHeaderFooter` = True

---

## ③ ★紙の 上で 測った事★（実ブラウザ・Chrome）

| 何 | 結果 |
|---|---|
| `position:fixed` は 毎ページ 出るか | ★出る★（全ページ 上端 y=11 で 一定・比べる相手は y が 増えた） |
| CSS の `counter(page)` は 増えるか | ★増えない★（31ページ 刷って ★10ページ目も 1桁＝「1」のまま★） |
| `counter(pages)` は | ★同じく 増えない★ |

⇒ ★だから 自分で ページを 割って 1枚ずつ 番号を 書く★（`lib/grid-print.js` の `割り付ける`）。

★実測★ … 120行を A4縦・枠線ありで 刷ると ★3ページ★
下に 「ページ 1 / 3」「ページ 2 / 3」「ページ 3 / 3」が 出た（実ブラウザで 絵も 撮った）。

---

## ④ ★うちで 決めた 物（隠さない）★

★「よく使う形」の 一覧（9個）は ★うちの 物★★。
実Excel の ヘッダーの 一覧は ★UIA から 起こせなかった★（4通り 試した）ので
★見た目は 写していない★。★中身（印）だけ 同じ★に してある。

★置き場所（&Z）は ブラウザでは 空★＝ブラウザは ファイルの 置き場所を 知らない。
★知らない物を それらしく 埋めない★。

---

★正本★ `lib/header-footer.js` ／ 見張り `tests/header-footer.test.mjs`（47本）
