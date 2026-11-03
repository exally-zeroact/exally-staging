# 版対応 検証ハーネス レポート

- この紙は **総なめのたびに作り直されます**（取った日は git の commit が覚えています）
- 真値: **O365HomePremRetail 16.0.20430.20092** (x64 / Current Channel)
- ロケール: UI=1041 / 国=81 / 小数点='.' 桁区切り=',' / 日付システム=1900
- 計算経路: **book.html の setCellFormula(本番と同じ)**。生の HyperFormula ではない。

> **★これは「今この瞬間の本番グリッドのお金バグ」である。** 下の「不一致(既知)」のうち money_impact の付いた項目は、
> 金額の端数・表示に直接効く。ハーネスを先に作る判断のため **このコミットでは直していない**（区分Aとして期限付きで台帳に登録）。

## 集計

| 判定 | 件数 |
|---|---|
| 一致 | 613 |
| 不一致(既知) | 2 |
| 不一致(新規) | 0 |
| 未検証 | 0 |
| 揮発性 | 2 |
| **合計** | 617 |

※ 「未検証」は緑ではない。その版の真値がまだ無い、という意味。

## 版ごとの状態

| 版 | 状態 |
|---|---|
| Excel 365 (16.0.20430.20092) | **真値**（このリポジトリの基準） |
| LibreOffice | **未検証**（goldenが無い。CIの別ジョブで生成する） |
| Excel 2016 / 2019 / Mac | **未検証**（実機が無い。golden/RECIPE.md の手順でその環境で1回走らせれば埋まる） |

## 揮発性の関数（別扱い）

TODAY / NOW は毎回答えが変わるため **golden突合の対象外**。固定値と比べると必ず腐るので、実行時点との一致だけを見る。

| ケース | 式 | 見方 | 期待 | 実際 | 判定 |
|---|---|---|---|---|---|
| TODAY_serial | `=TODAY()*1` | 実行時点の日付シリアルと一致するか | (今日の 値) | (今日の 値) | OK |
| NOW_int_is_today | `=INT(NOW())-TODAY()` | NOWの整数部がTODAYと一致するか | 0 | 0 | OK |

## 不一致（新規）＝赤

なし。

## 不一致（既知＝台帳にあり・緑だが必ず全件出す）

| 区分 | 関数 | ケース | 式 | Exally | Excel真値 | 中身と期限 |
|---|---|---|---|---|---|---|
| C | TEXT | TEXT_W_19000101_f5 | `=TEXT(DATE(1900,1,1),"ge.m.d")` | M33.1.2 | M33.1.1 | ★書式の話ではない★。DATE(1900,1,1) の通し番号が 借り物(HF)で 2、実Excel で 1（1900/2/29 が在る Excel の仕様）。RECIPE の決め『serial 60 以前は区分C』。和暦の字そのもの（M33）は合っている |
| C | TEXT | TEXT_W_19000101_f6 | `=TEXT(DATE(1900,1,1),"ggge年m月d日")` | 明治33年1月2日 | 明治33年1月1日 | ★書式の話ではない★。TEXT_W_19000101_f5 と同じ（DATE(1900,1,1) の通し番号が1日ずれる）。区分C |

## 入力の型が保たれるか（別枠）

セルに打ち込んだ文字をどう解釈するか。関数の検証とは別の話なので、混ぜずにここで見る。
期待値は想像ではなく、**実Excelの標準書式セルに同じ文字を打ち込んだ実測値**。

| ケース | 打ち込んだ値 | Exally | Excel(標準書式セル) | Excel型 | 判定 | 中身 |
|---|---|---|---|---|---|---|
| INPUT_typed_0007 | `0007` | 7 | 7 | n | 一致 |  |
| INPUT_typed_comma | `1,234` | 1234 | 1234 | n | 一致 |  |
| INPUT_typed_datelike | `2026-07-31` | 46234 | 46234 | n | 一致 |  |
| INPUT_typed_code | `007-1234` | 007-1234 | 007-1234 | s | 一致 |  |

## ★独自層が生HFより悪くしていないか

独自層に関数を足すと、HyperFormula なら合っていた物を壊すことがある（TEXTの日付書式で実際に起きた）。
**独自層に関数を足す時は、必ずここが増えていないことを確認してから足す。** 台帳に載っていない劣化は赤。

劣化なし。

## 経路の固定（将来 生HF に落ちたら気付くための錠）

- 独自層(_jsComputeFormula)が答えたケース: **0件**
- 生HFと本番経路で答えが違うケース: **388件** … この差が消えたら「素通りに落ちた」ということ
- 独自層の入口: {"jsSetCount":1,"entryPoints":1,"pluginRegistered":true,"pluginCount":43}（1つだけであること）

## 全ケース

| 関数 | ケース | 式 | Exally(本番経路) | Excel365(真値) | LibreOffice | 生HF | 判定 | 区分 |
|---|---|---|---|---|---|---|---|---|
| ROUND | ROUND_pos_2 | `=ROUND(A5,2)` | 2.68 | 2.68 | 未検証 | 2.68 | 一致 |  |
| ROUND | ROUND_half_up | `=ROUND(A6,0)` | 5 | 5 | 未検証 | 5 | 一致 |  |
| ROUND | ROUND_half_neg | `=ROUND(-2.5,0)` | -3 | -3 | 未検証 | -3 | 一致 |  |
| ROUND | ROUND_neg_digits | `=ROUND(A2,-2)` | 2500 | 2500 | 未検証 | 2500 | 一致 |  |
| ROUND | ROUND_zero | `=ROUND(A7,2)` | 0 | 0 | 未検証 | 0 | 一致 |  |
| ROUND | ROUND_money_10pct | `=ROUND(A2*0.1,0)` | 250 | 250 | 未検証 | 250 | 一致 |  |
| ROUNDUP | ROUNDUP_pos | `=ROUNDUP(2.001,0)` | 3 | 3 | 未検証 | 3 | 一致 |  |
| ROUNDUP | ROUNDUP_neg | `=ROUNDUP(-2.4,0)` | -3 | -3 | 未検証 | -3 | 一致 |  |
| ROUNDUP | ROUNDUP_digits | `=ROUNDUP(A5,1)` | 2.7 | 2.7 | 未検証 | 2.7 | 一致 |  |
| ROUNDUP | ROUNDUP_zero | `=ROUNDUP(A7,0)` | 0 | 0 | 未検証 | 0 | 一致 |  |
| ROUNDDOWN | ROUNDDOWN_pos | `=ROUNDDOWN(2.999,0)` | 2 | 2 | 未検証 | 2 | 一致 |  |
| ROUNDDOWN | ROUNDDOWN_neg | `=ROUNDDOWN(-2.6,0)` | -2 | -2 | 未検証 | -2 | 一致 |  |
| ROUNDDOWN | ROUNDDOWN_shohizei | `=ROUNDDOWN(1980*1.1,0)` | 2178 | 2178 | 未検証 | 2178 | 一致 |  |
| ROUNDDOWN | ROUNDDOWN_digits | `=ROUNDDOWN(A5,2)` | 2.67 | 2.67 | 未検証 | 2.67 | 一致 |  |
| INT | INT_pos | `=INT(2.9)` | 2 | 2 | 未検証 | 2 | 一致 |  |
| INT | INT_neg | `=INT(-2.5)` | -3 | -3 | 未検証 | -2 | 一致 |  |
| INT | INT_neg_cell | `=INT(A8)` | -3 | -3 | 未検証 | -2 | 一致 |  |
| INT | INT_zero | `=INT(A7)` | 0 | 0 | 未検証 | 0 | 一致 |  |
| MOD | MOD_pos | `=MOD(7,3)` | 1 | 1 | 未検証 | 1 | 一致 |  |
| MOD | MOD_neg_dividend | `=MOD(-3,2)` | 1 | 1 | 未検証 | -1 | 一致 |  |
| MOD | MOD_neg_divisor | `=MOD(3,-2)` | -1 | -1 | 未検証 | 1 | 一致 |  |
| MOD | MOD_cell | `=MOD(A3,2)` | 1 | 1 | 未検証 | -1 | 一致 |  |
| MOD | MOD_div0 | `=MOD(3,0)` | #DIV/0! | #DIV/0! | 未検証 | #DIV/0! | 一致 |  |
| MOD | MOD_both_neg | `=MOD(-3,-2)` | -1 | -1 | 未検証 | -1 | 一致 |  |
| MOD | MOD_nested | `=ROUND(MOD(-3,2),0)` | 1 | 1 | 未検証 | -1 | 一致 |  |
| INT | INT_neg_small | `=INT(-0.5)` | -1 | -1 | 未検証 | 0 | 一致 |  |
| INT | INT_nested | `=SUM(INT(-2.5),0)` | -3 | -3 | 未検証 | -2 | 一致 |  |
| ABS | ABS_neg | `=ABS(A3)` | 3 | 3 | 未検証 | 3 | 一致 |  |
| ABS | ABS_pos | `=ABS(A1)` | 1000 | 1000 | 未検証 | 1000 | 一致 |  |
| ABS | ABS_zero | `=ABS(A7)` | 0 | 0 | 未検証 | 0 | 一致 |  |
| MAX | MAX_range | `=MAX(E1:E6)` | 600 | 600 | 未検証 | 600 | 一致 |  |
| MAX | MAX_with_text | `=MAX(A1:B8)` | 2500 | 2500 | 未検証 | 2500 | 一致 |  |
| MAX | MAX_neg_only | `=MAX(A3,A8)` | -2.5 | -2.5 | 未検証 | -2.5 | 一致 |  |
| MIN | MIN_range | `=MIN(E1:E6)` | 100 | 100 | 未検証 | 100 | 一致 |  |
| MIN | MIN_with_text | `=MIN(A1:B8)` | -3 | -3 | 未検証 | -3 | 一致 |  |
| MIN | MIN_empty_cell | `=MIN(G1:G2)` | 0 | 0 | 未検証 | 0 | 一致 |  |
| AVERAGE | AVERAGE_range | `=AVERAGE(E1:E6)` | 350 | 350 | 未検証 | 350 | 一致 |  |
| AVERAGE | AVERAGE_with_text | `=AVERAGE(A1:B8)` | 437.721875 | 437.721875 | 未検証 | 437.721875 | 一致 |  |
| AVERAGE | AVERAGE_blank_skip | `=AVERAGE(G1:G2)` | 0 | 0 | 未検証 | 0 | 一致 |  |
| AVERAGE | AVERAGE_frac | `=ROUND(AVERAGE(C1:C6),6)` | 31 | 31 | 未検証 | 31 | 一致 |  |
| SUM | SUM_range | `=SUM(E1:E6)` | 2100 | 2100 | 未検証 | 2100 | 一致 |  |
| SUM | SUM_text_ignored | `=SUM(A1:B8)` | 3501.775 | 3501.775 | 未検証 | 3501.775 | 一致 |  |
| SUM | SUM_blank | `=SUM(G1:G3)` | 0 | 0 | 未検証 | 0 | 一致 |  |
| SUM | SUM_float | `=SUM(A4,0.2)` | 0.30000000000000004 | 0.30000000000000004 | 未検証 | 0.3 | 一致 |  |
| SUM | SUM_mixed_args | `=SUM(E1:E3,100,A1)` | 1700 | 1700 | 未検証 | 1700 | 一致 |  |
| SUMPRODUCT | SUMPRODUCT_2range | `=SUMPRODUCT(C1:C6,E1:E6)` | 97100 | 97100 | 未検証 | 97100 | 一致 |  |
| SUMPRODUCT | SUMPRODUCT_1range | `=SUMPRODUCT(E1:E6)` | 2100 | 2100 | 未検証 | 2100 | 一致 |  |
| SUMPRODUCT | SUMPRODUCT_cond | `=SUMPRODUCT((D1:D6="A")*E1:E6)` | 1000 | 1000 | 未検証 | #VALUE! | 一致 |  |
| SUMPRODUCT | SUMPRODUCT_len | `=SUMPRODUCT(LEN(B1:B3))` | 11 | 11 | 未検証 | #VALUE! | 一致 |  |
| COUNT | COUNT_numbers | `=COUNT(A1:A8)` | 8 | 8 | 未検証 | 8 | 一致 |  |
| COUNT | COUNT_mixed | `=COUNT(A1:B8)` | 8 | 8 | 未検証 | 8 | 一致 |  |
| COUNT | COUNT_text_number | `=COUNT(B6:B7)` | 0 | 0 | 未検証 | 0 | 一致 |  |
| COUNT | COUNT_blank_range | `=COUNT(G1:G3)` | 1 | 1 | 未検証 | 1 | 一致 |  |
| COUNTA | COUNTA_mixed | `=COUNTA(A1:B8)` | 16 | 16 | 未検証 | 16 | 一致 |  |
| COUNTA | COUNTA_blank | `=COUNTA(G1:G3)` | 2 | 2 | 未検証 | 2 | 一致 |  |
| COUNTA | COUNTA_bool | `=COUNTA(H1:H2)` | 2 | 2 | 未検証 | 2 | 一致 |  |
| COUNTBLANK | COUNTBLANK_range | `=COUNTBLANK(G1:G3)` | 1 | 1 | 未検証 | 1 | 一致 |  |
| COUNTBLANK | COUNTBLANK_none | `=COUNTBLANK(E1:E6)` | 0 | 0 | 未検証 | 0 | 一致 |  |
| COUNTBLANK | COUNTBLANK_wide | `=COUNTBLANK(G1:H2)` | 1 | 1 | 未検証 | 1 | 一致 |  |
| COUNTIF | COUNTIF_ge | `=COUNTIF(C1:C6,">=10")` | 4 | 4 | 未検証 | 4 | 一致 |  |
| COUNTIF | COUNTIF_eq_text | `=COUNTIF(D1:D6,"A")` | 3 | 3 | 未検証 | 3 | 一致 |  |
| COUNTIF | COUNTIF_wildcard | `=COUNTIF(B1:B8,"*ん*")` | 2 | 2 | 未検証 | 2 | 一致 |  |
| COUNTIF | COUNTIF_wild_q | `=COUNTIF(B1:B8,"A-?")` | 1 | 1 | 未検証 | 1 | 一致 |  |
| COUNTIF | COUNTIF_cellref | `=COUNTIF(D1:D6,D1)` | 3 | 3 | 未検証 | 3 | 一致 |  |
| COUNTIF | COUNTIF_zero_blank | `=COUNTIF(G1:G3,0)` | 2 | 2 | 未検証 | 2 | 一致 |  |
| COUNTIFS | COUNTIFS_2cond | `=COUNTIFS(D1:D6,"A",E1:E6,">100")` | 2 | 2 | 未検証 | 2 | 一致 |  |
| COUNTIFS | COUNTIFS_range_num | `=COUNTIFS(C1:C6,">=5",C1:C6,"<=50")` | 4 | 4 | 未検証 | 4 | 一致 |  |
| COUNTIFS | COUNTIFS_nomatch | `=COUNTIFS(D1:D6,"Z")` | 0 | 0 | 未検証 | 0 | 一致 |  |
| COUNTIFS | COUNTIFS_ne | `=COUNTIFS(D1:D6,"<>A")` | 3 | 3 | 未検証 | 3 | 一致 |  |
| SUMIF | SUMIF_ge | `=SUMIF(C1:C6,">=10",E1:E6)` | 1800 | 1800 | 未検証 | 1800 | 一致 |  |
| SUMIF | SUMIF_text_key | `=SUMIF(D1:D6,"B",E1:E6)` | 700 | 700 | 未検証 | 700 | 一致 |  |
| SUMIF | SUMIF_no_sumrange | `=SUMIF(E1:E6,">300")` | 1500 | 1500 | 未検証 | 1500 | 一致 |  |
| SUMIF | SUMIF_wildcard | `=SUMIF(B1:B6,"*ん*",E1:E6)` | 300 | 300 | 未検証 | 300 | 一致 |  |
| SUMIF | SUMIF_nomatch | `=SUMIF(D1:D6,"Z",E1:E6)` | 0 | 0 | 未検証 | 0 | 一致 |  |
| SUMIFS | SUMIFS_2cond | `=SUMIFS(E1:E6,D1:D6,"A",C1:C6,">1")` | 900 | 900 | 未検証 | 900 | 一致 |  |
| SUMIFS | SUMIFS_range_num | `=SUMIFS(E1:E6,C1:C6,">=5",C1:C6,"<=50")` | 1400 | 1400 | 未検証 | 1400 | 一致 |  |
| SUMIFS | SUMIFS_ne | `=SUMIFS(E1:E6,D1:D6,"<>A")` | 1100 | 1100 | 未検証 | 1100 | 一致 |  |
| SUMIFS | SUMIFS_nomatch | `=SUMIFS(E1:E6,D1:D6,"Z")` | 0 | 0 | 未検証 | 0 | 一致 |  |
| SUMIFS | SUMIFS_cellref | `=SUMIFS(E1:E6,D1:D6,D2)` | 700 | 700 | 未検証 | 700 | 一致 |  |
| IF | IF_true | `=IF(A1>500,"大","小")` | 大 | 大 | 未検証 | 大 | 一致 |  |
| IF | IF_false | `=IF(A3>0,"大","小")` | 小 | 小 | 未検証 | 小 | 一致 |  |
| IF | IF_nested | `=IF(A1>2000,"甲",IF(A1>500,"乙","丙"))` | 乙 | 乙 | 未検証 | 乙 | 一致 |  |
| IF | IF_blank_cond | `=IF(G1="","空","有")` | 空 | 空 | 未検証 | 空 | 一致 |  |
| IF | IF_num_result | `=IF(H1,A1,A2)` | 1000 | 1000 | 未検証 | 1000 | 一致 |  |
| IFS | IFS_first | `=IFS(A1>2000,"大",A1>500,"中",TRUE,"小")` | 中 | 中 | 未検証 | #NAME? | 一致 |  |
| IFS | IFS_second | `=IFS(A1>5000,"大",A1>500,"中",TRUE,"小")` | 中 | 中 | 未検証 | #NAME? | 一致 |  |
| IFS | IFS_fallback | `=IFS(A7>5000,"大",A7>500,"中",TRUE,"小")` | 小 | 小 | 未検証 | #NAME? | 一致 |  |
| IFS | IFS_no_match | `=IFERROR(IFS(A7>5000,"大",A7>500,"中"),"NA")` | NA | NA | 未検証 | NA | 一致 |  |
| IFERROR | IFERROR_div0 | `=IFERROR(1/0,"err")` | err | err | 未検証 | err | 一致 |  |
| IFERROR | IFERROR_ok | `=IFERROR(A1/A2,"err")` | 0.4 | 0.4 | 未検証 | 0.4 | 一致 |  |
| IFERROR | IFERROR_na | `=IFERROR(NA(),"err")` | err | err | 未検証 | err | 一致 |  |
| IFERROR | IFERROR_value | `=IFERROR(VALUE("あ"),-1)` | -1 | -1 | 未検証 | -1 | 一致 |  |
| IFNA | IFNA_na | `=IFNA(NA(),"なし")` | なし | なし | 未検証 | なし | 一致 |  |
| IFNA | IFNA_div0_passes | `=IFERROR(IFNA(1/0,"なし"),"div0が素通り")` | div0が素通り | div0が素通り | 未検証 | div0が素通り | 一致 |  |
| IFNA | IFNA_ok | `=IFNA(A1,"なし")` | 1000 | 1000 | 未検証 | 1000 | 一致 |  |
| AND | AND_all_true | `=AND(A1>0,A2>0)` | TRUE | TRUE | 未検証 | TRUE | 一致 |  |
| AND | AND_one_false | `=AND(A1>0,A3>0)` | FALSE | FALSE | 未検証 | FALSE | 一致 |  |
| AND | AND_cellbool | `=AND(H1,H1)` | TRUE | TRUE | 未検証 | TRUE | 一致 |  |
| OR | OR_one_true | `=OR(A3>0,A1>0)` | TRUE | TRUE | 未検証 | TRUE | 一致 |  |
| OR | OR_all_false | `=OR(A3>0,A8>0)` | FALSE | FALSE | 未検証 | FALSE | 一致 |  |
| OR | OR_cellbool | `=OR(H2,H1)` | TRUE | TRUE | 未検証 | TRUE | 一致 |  |
| NOT | NOT_true | `=NOT(TRUE)` | FALSE | FALSE | 未検証 | #NAME? | 一致 |  |
| NOT | NOT_cell | `=NOT(H2)` | TRUE | TRUE | 未検証 | TRUE | 一致 |  |
| NOT | NOT_expr | `=NOT(A1>A2)` | TRUE | TRUE | 未検証 | TRUE | 一致 |  |
| AND | LOGIC_combo | `=IF(AND(OR(A1>0,NOT(TRUE)),A3<0),"y","n")` | y | y | 未検証 | #NAME? | 一致 |  |
| NOT | LOGIC_bool_to_num | `=(A1>0)*1+(A3>0)*1` | 1 | 1 | 未検証 | 1 | 一致 |  |
| VLOOKUP | VLOOKUP_exact | `=VLOOKUP(20,C1:E6,3,FALSE)` | 400 | 400 | 未検証 | #NAME? | 一致 |  |
| VLOOKUP | VLOOKUP_exact_miss | `=IFERROR(VLOOKUP(30,C1:E6,3,FALSE),"NA")` | NA | NA | 未検証 | NA | 一致 |  |
| VLOOKUP | VLOOKUP_approx | `=VLOOKUP(30,C1:E6,3,TRUE)` | 400 | 400 | 未検証 | #NAME? | 一致 |  |
| VLOOKUP | VLOOKUP_approx_low | `=IFERROR(VLOOKUP(0,C1:E6,3,TRUE),"NA")` | NA | NA | 未検証 | NA | 一致 |  |
| VLOOKUP | VLOOKUP_col2 | `=VLOOKUP(50,C1:E6,2,FALSE)` | B | B | 未検証 | #NAME? | 一致 |  |
| VLOOKUP | VLOOKUP_text_key | `=VLOOKUP("C",D1:E6,2,FALSE)` | 400 | 400 | 未検証 | #NAME? | 一致 |  |
| VLOOKUP | VLOOKUP_badcol | `=IFERROR(VLOOKUP(20,C1:E6,9,FALSE),"ERR")` | ERR | ERR | 未検証 | ERR | 一致 |  |
| XLOOKUP | XLOOKUP_basic | `=XLOOKUP(20,C1:C6,E1:E6)` | 400 | 400 | 未検証 | 400 | 一致 |  |
| XLOOKUP | XLOOKUP_notfound | `=XLOOKUP(999,C1:C6,E1:E6,"なし")` | なし | なし | 未検証 | なし | 一致 |  |
| XLOOKUP | XLOOKUP_text | `=XLOOKUP("C",D1:D6,E1:E6,"なし")` | 400 | 400 | 未検証 | 400 | 一致 |  |
| XLOOKUP | XLOOKUP_first_dup | `=XLOOKUP("A",D1:D6,E1:E6,"なし")` | 100 | 100 | 未検証 | 100 | 一致 |  |
| XLOOKUP | XLOOKUP_no_default | `=IFERROR(XLOOKUP(999,C1:C6,E1:E6),"NA")` | NA | NA | 未検証 | NA | 一致 |  |
| INDEX | INDEX_row | `=INDEX(E1:E6,4)` | 400 | 400 | 未検証 | 400 | 一致 |  |
| INDEX | INDEX_2d | `=INDEX(C1:E6,4,3)` | 400 | 400 | 未検証 | 400 | 一致 |  |
| INDEX | INDEX_out_of_range | `=IFERROR(INDEX(E1:E6,99),"ERR")` | ERR | ERR | 未検証 | ERR | 一致 |  |
| INDEX | INDEX_text_range | `=INDEX(B1:B8,8)` | 山田 太郎 | 山田 太郎 | 未検証 | 山田 太郎 | 一致 |  |
| MATCH | MATCH_exact | `=MATCH(20,C1:C6,0)` | 4 | 4 | 未検証 | 4 | 一致 |  |
| MATCH | MATCH_exact_miss | `=IFERROR(MATCH(30,C1:C6,0),"NA")` | NA | NA | 未検証 | NA | 一致 |  |
| MATCH | MATCH_approx_asc | `=MATCH(30,C1:C6,1)` | 4 | 4 | 未検証 | 4 | 一致 |  |
| MATCH | MATCH_text | `=MATCH("C",D1:D6,0)` | 4 | 4 | 未検証 | 4 | 一致 |  |
| MATCH | MATCH_wildcard | `=MATCH("りん*",B1:B8,0)` | 1 | 1 | 未検証 | #N/A | 一致 |  |
| XLOOKUP | XLOOKUP_nested | `=ROUND(XLOOKUP(20,C1:C6,E1:E6),0)` | 400 | 400 | 未検証 | 400 | 一致 |  |
| MATCH | MATCH_nested_wild | `=INDEX(E1:E6,MATCH("りん*",B1:B8,0))` | 100 | 100 | 未検証 | #N/A | 一致 |  |
| INDEX | INDEX_MATCH_combo | `=INDEX(E1:E6,MATCH(20,C1:C6,0))` | 400 | 400 | 未検証 | 400 | 一致 |  |
| INDEX | INDEX_MATCH_text | `=INDEX(E1:E6,MATCH("C",D1:D6,0))` | 400 | 400 | 未検証 | 400 | 一致 |  |
| DATE | DATE_bare | `=DATE(2026,7,31)` | 46234 | 46234 | 未検証 | 46234 | 一致 |  |
| EOMONTH | EOMONTH_bare | `=EOMONTH(F2,0)` | 46081 | 46081 | 未検証 | 46081 | 一致 |  |
| DATE | DATE_basic | `=DATE(2026,7,31)*1` | 46234 | 46234 | 未検証 | 46234 | 一致 |  |
| DATE | DATE_month_over | `=DATE(2026,13,1)*1` | 46388 | 46388 | 未検証 | 46388 | 一致 |  |
| DATE | DATE_day_over | `=DATE(2026,1,32)*1` | 46054 | 46054 | 未検証 | 46054 | 一致 |  |
| DATE | DATE_leap_2024 | `=DATE(2024,2,29)*1` | 45351 | 45351 | 未検証 | 45351 | 一致 |  |
| DATE | DATE_diff_days | `=DATE(2026,7,31)-DATE(2026,7,1)` | 30 | 30 | 未検証 | 30 | 一致 |  |
| YEAR | YEAR_cell | `=YEAR(F5)` | 2026 | 2026 | 未検証 | 2026 | 一致 |  |
| YEAR | YEAR_expr | `=YEAR(DATE(2000,3,1))` | 2000 | 2000 | 未検証 | 2000 | 一致 |  |
| YEAR | YEAR_old | `=YEAR(F6)` | 2000 | 2000 | 未検証 | 2000 | 一致 |  |
| MONTH | MONTH_cell | `=MONTH(F5)` | 7 | 7 | 未検証 | 7 | 一致 |  |
| MONTH | MONTH_jan | `=MONTH(F1)` | 1 | 1 | 未検証 | 1 | 一致 |  |
| MONTH | MONTH_expr | `=MONTH(DATE(2026,12,1))` | 12 | 12 | 未検証 | 12 | 一致 |  |
| DAY | DAY_cell | `=DAY(F5)` | 31 | 31 | 未検証 | 31 | 一致 |  |
| DAY | DAY_month_end | `=DAY(F3)` | 31 | 31 | 未検証 | 31 | 一致 |  |
| DAY | DAY_ymd_combo | `=YEAR(F5)*10000+MONTH(F5)*100+DAY(F5)` | 20260731 | 20260731 | 未検証 | 20260731 | 一致 |  |
| EOMONTH | EOMONTH_0 | `=EOMONTH(F2,0)*1` | 46081 | 46081 | 未検証 | 46081 | 一致 |  |
| EOMONTH | EOMONTH_plus1 | `=EOMONTH(F1,1)*1` | 46081 | 46081 | 未検証 | 46081 | 一致 |  |
| EOMONTH | EOMONTH_minus1 | `=EOMONTH(F3,-1)*1` | 46081 | 46081 | 未検証 | 46081 | 一致 |  |
| EOMONTH | EOMONTH_leap | `=EOMONTH(DATE(2024,2,1),0)*1` | 45351 | 45351 | 未検証 | 45351 | 一致 |  |
| EOMONTH | EOMONTH_year_over | `=EOMONTH(DATE(2026,12,15),1)*1` | 46418 | 46418 | 未検証 | 46418 | 一致 |  |
| DATEDIF | DATEDIF_Y | `=DATEDIF(DATE(2000,3,1),DATE(2026,2,28),"Y")` | 25 | 25 | 未検証 | 25 | 一致 |  |
| DATEDIF | DATEDIF_Y_after | `=DATEDIF(DATE(2000,3,1),DATE(2026,3,1),"Y")` | 26 | 26 | 未検証 | 26 | 一致 |  |
| DATEDIF | DATEDIF_M | `=DATEDIF(DATE(2026,1,31),DATE(2026,3,30),"M")` | 1 | 1 | 未検証 | 1 | 一致 |  |
| DATEDIF | DATEDIF_D | `=DATEDIF(F1,F5,"D")` | 181 | 181 | 未検証 | 181 | 一致 |  |
| DATEDIF | DATEDIF_MD | `=DATEDIF(DATE(2026,1,31),DATE(2026,3,30),"MD")` | 27 | 27 | 未検証 | 27 | 一致 |  |
| DATEDIF | DATEDIF_YM | `=DATEDIF(DATE(2000,3,1),DATE(2026,2,28),"YM")` | 11 | 11 | 未検証 | 11 | 一致 |  |
| WEEKDAY | WEEKDAY_default | `=WEEKDAY(F5)` | 6 | 6 | 未検証 | 6 | 一致 |  |
| WEEKDAY | WEEKDAY_type2 | `=WEEKDAY(F5,2)` | 5 | 5 | 未検証 | 5 | 一致 |  |
| WEEKDAY | WEEKDAY_type3 | `=WEEKDAY(F5,3)` | 4 | 4 | 未検証 | 4 | 一致 |  |
| WEEKDAY | WEEKDAY_sunday | `=WEEKDAY(F4,2)` | 3 | 3 | 未検証 | 3 | 一致 |  |
| TRIM | TRIM_bare | `=TRIM(B4)` | pad | pad | 未検証 | pad | 一致 |  |
| TEXT | TEXT_thousands | `=TEXT(1234.5,"#,##0")` | 1,235 | 1,235 | 未検証 | 1235,##0 | 一致 |  |
| TEXT | TEXT_thousands_dec | `=TEXT(1234.567,"#,##0.00")` | 1,234.57 | 1,234.57 | 未検証 | 1235,##0.00 | 一致 |  |
| TEXT | TEXT_percent | `=TEXT(0.1235,"0.0%")` | 12.4% | 12.4% | 未検証 | 0.1% | 一致 |  |
| TEXT | TEXT_currency_yen | `=TEXT(1234.5,"¥#,##0")` | ¥1,235 | ¥1,235 | 未検証 | ¥1235,##0 | 一致 |  |
| TEXT | TEXT_pad_zero | `=TEXT(7,"0000")` | 0007 | 0007 | 未検証 | 0007 | 一致 |  |
| TEXT | TEXT_date | `=TEXT(F5,"yyyy/mm/dd")` | 2026/07/31 | 2026/07/31 | 未検証 | 2026/07/31 | 一致 |  |
| TEXT | TEXT_date_slash_md | `=TEXT(F1,"m/d")` | 1/31 | 1/31 | 未検証 | 1/31 | 一致 |  |
| TEXT | TEXT_neg_paren | `=TEXT(-1234,"#,##0;(#,##0)")` | (1,234) | (1,234) | 未検証 | -1234,##0;(#,##0) | 一致 |  |
| TEXT | TEXT_cellref | `=TEXT(A2,"#,##0")` | 2,500 | 2,500 | 未検証 | 2500,##0 | 一致 |  |
| LEFT | LEFT_jp | `=LEFT(B1,2)` | りん | りん | 未検証 | りん | 一致 |  |
| LEFT | LEFT_1 | `=LEFT(B3,1)` | a | a | 未検証 | a | 一致 |  |
| LEFT | LEFT_over | `=LEFT(B1,99)` | りんご | りんご | 未検証 | りんご | 一致 |  |
| LEFT | LEFT_zero | `="["&LEFT(B1,0)&"]"` | [] | [] | 未検証 | [] | 一致 |  |
| RIGHT | RIGHT_jp | `=RIGHT(B1,1)` | ご | ご | 未検証 | ご | 一致 |  |
| RIGHT | RIGHT_ascii | `=RIGHT(B3,3)` | ple | ple | 未検証 | ple | 一致 |  |
| RIGHT | RIGHT_over | `=RIGHT(B3,99)` | apple | apple | 未検証 | apple | 一致 |  |
| MID | MID_jp | `=MID(B1,2,2)` | んご | んご | 未検証 | んご | 一致 |  |
| MID | MID_start1 | `=MID(B8,1,2)` | 山田 | 山田 | 未検証 | 山田 | 一致 |  |
| MID | MID_beyond | `="["&MID(B3,99,2)&"]"` | [] | [] | 未検証 | [] | 一致 |  |
| LEN | LEN_jp | `=LEN(B1)` | 3 | 3 | 未検証 | 3 | 一致 |  |
| LEN | LEN_ascii | `=LEN(B3)` | 5 | 5 | 未検証 | 5 | 一致 |  |
| LEN | LEN_space | `=LEN(B4)` | 5 | 5 | 未検証 | 5 | 一致 |  |
| LEN | LEN_num_text | `=LEN(B6)` | 4 | 4 | 未検証 | 4 | 一致 |  |
| LEN | LEN_number_cell | `=LEN(A1)` | 4 | 4 | 未検証 | 4 | 一致 |  |
| FIND | FIND_jp | `=FIND("ん",B2)` | 3 | 3 | 未検証 | 3 | 一致 |  |
| FIND | FIND_ascii | `=FIND("pp",B3)` | 2 | 2 | 未検証 | 2 | 一致 |  |
| FIND | FIND_missing | `=IFERROR(FIND("z",B3),"NA")` | NA | NA | 未検証 | NA | 一致 |  |
| FIND | FIND_start | `=FIND("p",B3,3)` | 3 | 3 | 未検証 | 3 | 一致 |  |
| FIND | FIND_case | `=IFERROR(FIND("A",B3),"NA")` | NA | NA | 未検証 | NA | 一致 |  |
| SUBSTITUTE | SUBSTITUTE_jp | `=SUBSTITUTE(B1,"ご","GO")` | りんGO | りんGO | 未検証 | りんGO | 一致 |  |
| SUBSTITUTE | SUBSTITUTE_all | `=SUBSTITUTE(B3,"p","P")` | aPPle | aPPle | 未検証 | aPPle | 一致 |  |
| SUBSTITUTE | SUBSTITUTE_nth | `=SUBSTITUTE(B3,"p","P",2)` | apPle | apPle | 未検証 | apPle | 一致 |  |
| SUBSTITUTE | SUBSTITUTE_none | `=SUBSTITUTE(B3,"z","Z")` | apple | apple | 未検証 | apple | 一致 |  |
| SUBSTITUTE | SUBSTITUTE_comma | `=SUBSTITUTE(B7,",","")` | 1234 | 1234 | 未検証 | 1234 | 一致 |  |
| CONCATENATE | CONCATENATE_2 | `=CONCATENATE(B1,"-",A1)` | りんご-1000 | りんご-1000 | 未検証 | りんご-1000 | 一致 |  |
| CONCATENATE | CONCATENATE_3 | `=CONCATENATE(B1,B2,B3)` | りんごみかんapple | りんごみかんapple | 未検証 | りんごみかんapple | 一致 |  |
| CONCATENATE | CONCATENATE_blank | `="["&CONCATENATE(G1,B3)&"]"` | [apple] | [apple] | 未検証 | [apple] | 一致 |  |
| & | AMP_text_num | `=B1&"-"&A1` | りんご-1000 | りんご-1000 | 未検証 | りんご-1000 | 一致 |  |
| & | AMP_blank | `="["&G1&"]"` | [] | [] | 未検証 | [] | 一致 |  |
| & | AMP_number_fmt | `="["&A4&"]"` | [0.1] | [0.1] | 未検証 | [0.1] | 一致 |  |
| & | AMP_bool | `="["&H1&"]"` | [TRUE] | [TRUE] | 未検証 | [TRUE] | 一致 |  |
| TRIM | TRIM_pad | `="["&TRIM(B4)&"]"` | [pad] | [pad] | 未検証 | [pad] | 一致 |  |
| TRIM | TRIM_inner | `="["&TRIM(B8)&"]"` | [山田 太郎] | [山田 太郎] | 未検証 | [山田 太郎] | 一致 |  |
| TRIM | TRIM_none | `="["&TRIM(B3)&"]"` | [apple] | [apple] | 未検証 | [apple] | 一致 |  |
| VALUE | VALUE_num_text | `=VALUE(B6)` | 7 | 7 | 未検証 | 7 | 一致 |  |
| VALUE | VALUE_comma | `=IFERROR(VALUE(B7),"NA")` | 1234 | 1234 | 未検証 | NA | 一致 |  |
| VALUE | VALUE_bad | `=IFERROR(VALUE(B1),"NA")` | NA | NA | 未検証 | NA | 一致 |  |
| VALUE | VALUE_zero_text | `=VALUE(G3)` | 0 | 0 | 未検証 | 0 | 一致 |  |
| TEXT | TEXT_pct_zero | `=TEXT(0.5,"0%")` | 50% | 50% | 未検証 | 1% | 一致 |  |
| TEXT | TEXT_dec_round | `=TEXT(2.675,"0.00")` | 2.68 | 2.68 | 未検証 | 2.67 | 一致 |  |
| TEXT | TEXT_neg_number | `=TEXT(-1234.5,"#,##0")` | -1,235 | -1,235 | 未検証 | -1235,##0 | 一致 |  |
| TEXT | TEXT_date_yy | `=TEXT(F5,"yy/m/d")` | 26/7/31 | 26/7/31 | 未検証 | 26/7/31 | 一致 |  |
| TEXT | TEXT_nested | `=LEN(TEXT(1234.5,"#,##0"))` | 5 | 5 | 未検証 | 8 | 一致 |  |
| VALUE | VALUE_empty | `=IFERROR(VALUE(""),"NA")` | NA | NA | 未検証 | NA | 一致 |  |
| VALUE | VALUE_nested_calc | `=IFERROR(VALUE(B7)+1,"NA")` | 1235 | 1235 | 未検証 | NA | 一致 |  |
| TEXTJOIN | TEXTJOIN_nested_len | `=LEN(TEXTJOIN(",",TRUE,D1:D3))` | 5 | 5 | 未検証 | #NAME? | 一致 |  |
| VALUE | VALUE_nested_sum | `=SUM(VALUE(B7),1)` | 1235 | 1235 | 未検証 | #VALUE! | 一致 |  |
| TEXT | TEXT_nested_concat | `="["&TEXT(A2,"#,##0")&"]"` | [2,500] | [2,500] | 未検証 | [2500,##0] | 一致 |  |
| TEXTJOIN | TEXTJOIN_skip | `=TEXTJOIN("\|",TRUE,B1:B3)` | りんご\|みかん\|apple | りんご\|みかん\|apple | 未検証 | #NAME? | 一致 |  |
| TEXTJOIN | TEXTJOIN_delim_comma | `=TEXTJOIN(",",TRUE,"a","b")` | a,b | a,b | 未検証 | #NAME? | 一致 |  |
| TEXTJOIN | TEXTJOIN_nested_fn | `=TEXTJOIN("\|",TRUE,LEFT(B1,1),LEFT(B2,1))` | り\|み | り\|み | 未検証 | #NAME? | 一致 |  |
| TEXTJOIN | TEXTJOIN_mixed | `=TEXTJOIN("-",TRUE,B5,A1,C1:C2)` | A-1-1000-1-5 | A-1-1000-1-5 | 未検証 | #NAME? | 一致 |  |
| TEXTJOIN | TEXTJOIN_keep | `=TEXTJOIN("\|",FALSE,G1:G3)` | \|0\|0 | \|0\|0 | 未検証 | #NAME? | 一致 |  |
| TEXTJOIN | TEXTJOIN_nums | `=TEXTJOIN(",",TRUE,E1:E6)` | 100,200,300,400,500,600 | 100,200,300,400,500,600 | 未検証 | #NAME? | 一致 |  |
| TEXTJOIN | TEXTJOIN_blank_skip | `=TEXTJOIN("\|",TRUE,G1:G3)` | 0\|0 | 0\|0 | 未検証 | #NAME? | 一致 |  |
| TEXT | TEXT_weekday_ja | `=TEXT(F1,"aaa")` | 土 | 土 | 未検証 | aaa | 一致 |  |
| TEXT | TEXT_weekday_ja_long | `=TEXT(F1,"aaaa")` | 土曜日 | 土曜日 | 未検証 | aaaa | 一致 |  |
| TEXT | TEXT_weekday_ja_locale | `=TEXT(F1,"[$-411]aaa")` | 土 | 土 | 未検証 | [$-411]aaa | 一致 |  |
| TEXT | TEXT_weekday_in_date | `=TEXT(F1,"m/d(aaa)")` | 1/31(土) | 1/31(土) | 未検証 | 1/31(aaa) | 一致 |  |
| TEXT | TEXT_weekday_section | `=TEXT(F1,"aaa;@")` | 土 | 土 | 未検証 | aaa;@ | 一致 |  |
| TEXT | TEXT_weekday_en | `=TEXT(F1,"ddd")` | Sat | Sat | 未検証 | 3131 | 一致 |  |
| TEXT | TEXT_weekday_en_long | `=TEXT(F1,"dddd")` | Saturday | Saturday | 未検証 | 3131 | 一致 |  |
| TEXT | TEXT_month_en | `=TEXT(F1,"mmm")` | Jan | Jan | 未検証 | 010 | 一致 |  |
| TEXT | TEXT_month_en_long | `=TEXT(F1,"mmmm")` | January | January | 未検証 | 0100 | 一致 |  |
| TEXT | TEXT_month_initial | `=TEXT(F1,"mmmmm")` | J | J | 未検証 | 01000 | 一致 |  |
| TEXT | TEXT_minute_vs_month | `=TEXT(46053.5,"hh:mm")` | 12:00 | 12:00 | 未検証 | 12:00 | 一致 |  |
| TEXT | TEXT_time_hms | `=TEXT(46053.5,"h:mm:ss")` | 12:00:00 | 12:00:00 | 未検証 | 12:00:00 | 一致 |  |
| TEXT | TEXT_date_and_time | `=TEXT(46053.5,"m/d h:mm")` | 1/31 12:00 | 1/31 12:00 | 未検証 | 1/31 12:00 | 一致 |  |
| TEXT | TEXT_time_ampm | `=TEXT(0.75,"h:mm AM/PM")` | 6:00 PM | 6:00 PM | 未検証 | 6:00 PM | 一致 |  |
| TEXT | TEXT_weekday_nested | `=LEN(TEXT(F1,"aaaa"))` | 3 | 3 | 未検証 | 4 | 一致 |  |
| TEXT | TEXT_era_wareki | `=TEXT(F1,"ge.m.d")` | R8.1.31 | R8.1.31 | 未検証 | ge.1.31 | 一致 |  |
| TEXT | TEXT_elapsed_hours | `=TEXT(F1,"[h]:mm")` | 1105272:00 | 1105272:00 | 未検証 | [0]:00 | 一致 |  |
| TEXT | TEXT_W_19000101_f0 | `=TEXT(DATE(1900,1,1),"g")` | M | M | 未検証 | g | 一致 |  |
| TEXT | TEXT_W_19000101_f1 | `=TEXT(DATE(1900,1,1),"gg")` | 明 | 明 | 未検証 | gg | 一致 |  |
| TEXT | TEXT_W_19000101_f2 | `=TEXT(DATE(1900,1,1),"ggg")` | 明治 | 明治 | 未検証 | ggg | 一致 |  |
| TEXT | TEXT_W_19000101_f3 | `=TEXT(DATE(1900,1,1),"e")` | 33 | 33 | 未検証 | e | 一致 |  |
| TEXT | TEXT_W_19000101_f4 | `=TEXT(DATE(1900,1,1),"ee")` | 33 | 33 | 未検証 | ee | 一致 |  |
| TEXT | TEXT_W_19000101_f5 | `=TEXT(DATE(1900,1,1),"ge.m.d")` | M33.1.2 | M33.1.1 | 未検証 | ge.1.1 | 不一致(既知) | C |
| TEXT | TEXT_W_19000101_f6 | `=TEXT(DATE(1900,1,1),"ggge年m月d日")` | 明治33年1月2日 | 明治33年1月1日 | 未検証 | ggge年1月1日 | 不一致(既知) | C |
| TEXT | TEXT_W_19000101_f7 | `=TEXT(DATE(1900,1,1),"gggee年")` | 明治33年 | 明治33年 | 未検証 | gggee年 | 一致 |  |
| TEXT | TEXT_W_19000101_f8 | `=TEXT(DATE(1900,1,1),"[$-ja-JP]ggge年")` | 明治33年 | 明治33年 | 未検証 | [$-ja-JP]ggge年 | 一致 |  |
| TEXT | TEXT_W_19120729_f0 | `=TEXT(DATE(1912,7,29),"g")` | M | M | 未検証 | g | 一致 |  |
| TEXT | TEXT_W_19120729_f1 | `=TEXT(DATE(1912,7,29),"gg")` | 明 | 明 | 未検証 | gg | 一致 |  |
| TEXT | TEXT_W_19120729_f2 | `=TEXT(DATE(1912,7,29),"ggg")` | 明治 | 明治 | 未検証 | ggg | 一致 |  |
| TEXT | TEXT_W_19120729_f3 | `=TEXT(DATE(1912,7,29),"e")` | 45 | 45 | 未検証 | e | 一致 |  |
| TEXT | TEXT_W_19120729_f4 | `=TEXT(DATE(1912,7,29),"ee")` | 45 | 45 | 未検証 | ee | 一致 |  |
| TEXT | TEXT_W_19120729_f5 | `=TEXT(DATE(1912,7,29),"ge.m.d")` | M45.7.29 | M45.7.29 | 未検証 | ge.7.29 | 一致 |  |
| TEXT | TEXT_W_19120729_f6 | `=TEXT(DATE(1912,7,29),"ggge年m月d日")` | 明治45年7月29日 | 明治45年7月29日 | 未検証 | ggge年7月29日 | 一致 |  |
| TEXT | TEXT_W_19120729_f7 | `=TEXT(DATE(1912,7,29),"gggee年")` | 明治45年 | 明治45年 | 未検証 | gggee年 | 一致 |  |
| TEXT | TEXT_W_19120729_f8 | `=TEXT(DATE(1912,7,29),"[$-ja-JP]ggge年")` | 明治45年 | 明治45年 | 未検証 | [$-ja-JP]ggge年 | 一致 |  |
| TEXT | TEXT_W_19120730_f0 | `=TEXT(DATE(1912,7,30),"g")` | T | T | 未検証 | g | 一致 |  |
| TEXT | TEXT_W_19120730_f1 | `=TEXT(DATE(1912,7,30),"gg")` | 大 | 大 | 未検証 | gg | 一致 |  |
| TEXT | TEXT_W_19120730_f2 | `=TEXT(DATE(1912,7,30),"ggg")` | 大正 | 大正 | 未検証 | ggg | 一致 |  |
| TEXT | TEXT_W_19120730_f3 | `=TEXT(DATE(1912,7,30),"e")` | 1 | 1 | 未検証 | e | 一致 |  |
| TEXT | TEXT_W_19120730_f4 | `=TEXT(DATE(1912,7,30),"ee")` | 01 | 01 | 未検証 | ee | 一致 |  |
| TEXT | TEXT_W_19120730_f5 | `=TEXT(DATE(1912,7,30),"ge.m.d")` | T1.7.30 | T1.7.30 | 未検証 | ge.7.30 | 一致 |  |
| TEXT | TEXT_W_19120730_f6 | `=TEXT(DATE(1912,7,30),"ggge年m月d日")` | 大正1年7月30日 | 大正1年7月30日 | 未検証 | ggge年7月30日 | 一致 |  |
| TEXT | TEXT_W_19120730_f7 | `=TEXT(DATE(1912,7,30),"gggee年")` | 大正01年 | 大正01年 | 未検証 | gggee年 | 一致 |  |
| TEXT | TEXT_W_19120730_f8 | `=TEXT(DATE(1912,7,30),"[$-ja-JP]ggge年")` | 大正1年 | 大正1年 | 未検証 | [$-ja-JP]ggge年 | 一致 |  |
| TEXT | TEXT_W_19261224_f0 | `=TEXT(DATE(1926,12,24),"g")` | T | T | 未検証 | g | 一致 |  |
| TEXT | TEXT_W_19261224_f1 | `=TEXT(DATE(1926,12,24),"gg")` | 大 | 大 | 未検証 | gg | 一致 |  |
| TEXT | TEXT_W_19261224_f2 | `=TEXT(DATE(1926,12,24),"ggg")` | 大正 | 大正 | 未検証 | ggg | 一致 |  |
| TEXT | TEXT_W_19261224_f3 | `=TEXT(DATE(1926,12,24),"e")` | 15 | 15 | 未検証 | e | 一致 |  |
| TEXT | TEXT_W_19261224_f4 | `=TEXT(DATE(1926,12,24),"ee")` | 15 | 15 | 未検証 | ee | 一致 |  |
| TEXT | TEXT_W_19261224_f5 | `=TEXT(DATE(1926,12,24),"ge.m.d")` | T15.12.24 | T15.12.24 | 未検証 | ge.12.24 | 一致 |  |
| TEXT | TEXT_W_19261224_f6 | `=TEXT(DATE(1926,12,24),"ggge年m月d日")` | 大正15年12月24日 | 大正15年12月24日 | 未検証 | ggge年12月24日 | 一致 |  |
| TEXT | TEXT_W_19261224_f7 | `=TEXT(DATE(1926,12,24),"gggee年")` | 大正15年 | 大正15年 | 未検証 | gggee年 | 一致 |  |
| TEXT | TEXT_W_19261224_f8 | `=TEXT(DATE(1926,12,24),"[$-ja-JP]ggge年")` | 大正15年 | 大正15年 | 未検証 | [$-ja-JP]ggge年 | 一致 |  |
| TEXT | TEXT_W_19261225_f0 | `=TEXT(DATE(1926,12,25),"g")` | S | S | 未検証 | g | 一致 |  |
| TEXT | TEXT_W_19261225_f1 | `=TEXT(DATE(1926,12,25),"gg")` | 昭 | 昭 | 未検証 | gg | 一致 |  |
| TEXT | TEXT_W_19261225_f2 | `=TEXT(DATE(1926,12,25),"ggg")` | 昭和 | 昭和 | 未検証 | ggg | 一致 |  |
| TEXT | TEXT_W_19261225_f3 | `=TEXT(DATE(1926,12,25),"e")` | 1 | 1 | 未検証 | e | 一致 |  |
| TEXT | TEXT_W_19261225_f4 | `=TEXT(DATE(1926,12,25),"ee")` | 01 | 01 | 未検証 | ee | 一致 |  |
| TEXT | TEXT_W_19261225_f5 | `=TEXT(DATE(1926,12,25),"ge.m.d")` | S1.12.25 | S1.12.25 | 未検証 | ge.12.25 | 一致 |  |
| TEXT | TEXT_W_19261225_f6 | `=TEXT(DATE(1926,12,25),"ggge年m月d日")` | 昭和1年12月25日 | 昭和1年12月25日 | 未検証 | ggge年12月25日 | 一致 |  |
| TEXT | TEXT_W_19261225_f7 | `=TEXT(DATE(1926,12,25),"gggee年")` | 昭和01年 | 昭和01年 | 未検証 | gggee年 | 一致 |  |
| TEXT | TEXT_W_19261225_f8 | `=TEXT(DATE(1926,12,25),"[$-ja-JP]ggge年")` | 昭和1年 | 昭和1年 | 未検証 | [$-ja-JP]ggge年 | 一致 |  |
| TEXT | TEXT_W_19890107_f0 | `=TEXT(DATE(1989,1,7),"g")` | S | S | 未検証 | g | 一致 |  |
| TEXT | TEXT_W_19890107_f1 | `=TEXT(DATE(1989,1,7),"gg")` | 昭 | 昭 | 未検証 | gg | 一致 |  |
| TEXT | TEXT_W_19890107_f2 | `=TEXT(DATE(1989,1,7),"ggg")` | 昭和 | 昭和 | 未検証 | ggg | 一致 |  |
| TEXT | TEXT_W_19890107_f3 | `=TEXT(DATE(1989,1,7),"e")` | 64 | 64 | 未検証 | e | 一致 |  |
| TEXT | TEXT_W_19890107_f4 | `=TEXT(DATE(1989,1,7),"ee")` | 64 | 64 | 未検証 | ee | 一致 |  |
| TEXT | TEXT_W_19890107_f5 | `=TEXT(DATE(1989,1,7),"ge.m.d")` | S64.1.7 | S64.1.7 | 未検証 | ge.1.7 | 一致 |  |
| TEXT | TEXT_W_19890107_f6 | `=TEXT(DATE(1989,1,7),"ggge年m月d日")` | 昭和64年1月7日 | 昭和64年1月7日 | 未検証 | ggge年1月7日 | 一致 |  |
| TEXT | TEXT_W_19890107_f7 | `=TEXT(DATE(1989,1,7),"gggee年")` | 昭和64年 | 昭和64年 | 未検証 | gggee年 | 一致 |  |
| TEXT | TEXT_W_19890107_f8 | `=TEXT(DATE(1989,1,7),"[$-ja-JP]ggge年")` | 昭和64年 | 昭和64年 | 未検証 | [$-ja-JP]ggge年 | 一致 |  |
| TEXT | TEXT_W_19890108_f0 | `=TEXT(DATE(1989,1,8),"g")` | H | H | 未検証 | g | 一致 |  |
| TEXT | TEXT_W_19890108_f1 | `=TEXT(DATE(1989,1,8),"gg")` | 平 | 平 | 未検証 | gg | 一致 |  |
| TEXT | TEXT_W_19890108_f2 | `=TEXT(DATE(1989,1,8),"ggg")` | 平成 | 平成 | 未検証 | ggg | 一致 |  |
| TEXT | TEXT_W_19890108_f3 | `=TEXT(DATE(1989,1,8),"e")` | 1 | 1 | 未検証 | e | 一致 |  |
| TEXT | TEXT_W_19890108_f4 | `=TEXT(DATE(1989,1,8),"ee")` | 01 | 01 | 未検証 | ee | 一致 |  |
| TEXT | TEXT_W_19890108_f5 | `=TEXT(DATE(1989,1,8),"ge.m.d")` | H1.1.8 | H1.1.8 | 未検証 | ge.1.8 | 一致 |  |
| TEXT | TEXT_W_19890108_f6 | `=TEXT(DATE(1989,1,8),"ggge年m月d日")` | 平成1年1月8日 | 平成1年1月8日 | 未検証 | ggge年1月8日 | 一致 |  |
| TEXT | TEXT_W_19890108_f7 | `=TEXT(DATE(1989,1,8),"gggee年")` | 平成01年 | 平成01年 | 未検証 | gggee年 | 一致 |  |
| TEXT | TEXT_W_19890108_f8 | `=TEXT(DATE(1989,1,8),"[$-ja-JP]ggge年")` | 平成1年 | 平成1年 | 未検証 | [$-ja-JP]ggge年 | 一致 |  |
| TEXT | TEXT_W_20190430_f0 | `=TEXT(DATE(2019,4,30),"g")` | H | H | 未検証 | g | 一致 |  |
| TEXT | TEXT_W_20190430_f1 | `=TEXT(DATE(2019,4,30),"gg")` | 平 | 平 | 未検証 | gg | 一致 |  |
| TEXT | TEXT_W_20190430_f2 | `=TEXT(DATE(2019,4,30),"ggg")` | 平成 | 平成 | 未検証 | ggg | 一致 |  |
| TEXT | TEXT_W_20190430_f3 | `=TEXT(DATE(2019,4,30),"e")` | 31 | 31 | 未検証 | e | 一致 |  |
| TEXT | TEXT_W_20190430_f4 | `=TEXT(DATE(2019,4,30),"ee")` | 31 | 31 | 未検証 | ee | 一致 |  |
| TEXT | TEXT_W_20190430_f5 | `=TEXT(DATE(2019,4,30),"ge.m.d")` | H31.4.30 | H31.4.30 | 未検証 | ge.4.30 | 一致 |  |
| TEXT | TEXT_W_20190430_f6 | `=TEXT(DATE(2019,4,30),"ggge年m月d日")` | 平成31年4月30日 | 平成31年4月30日 | 未検証 | ggge年4月30日 | 一致 |  |
| TEXT | TEXT_W_20190430_f7 | `=TEXT(DATE(2019,4,30),"gggee年")` | 平成31年 | 平成31年 | 未検証 | gggee年 | 一致 |  |
| TEXT | TEXT_W_20190430_f8 | `=TEXT(DATE(2019,4,30),"[$-ja-JP]ggge年")` | 平成31年 | 平成31年 | 未検証 | [$-ja-JP]ggge年 | 一致 |  |
| TEXT | TEXT_W_20190501_f0 | `=TEXT(DATE(2019,5,1),"g")` | R | R | 未検証 | g | 一致 |  |
| TEXT | TEXT_W_20190501_f1 | `=TEXT(DATE(2019,5,1),"gg")` | 令 | 令 | 未検証 | gg | 一致 |  |
| TEXT | TEXT_W_20190501_f2 | `=TEXT(DATE(2019,5,1),"ggg")` | 令和 | 令和 | 未検証 | ggg | 一致 |  |
| TEXT | TEXT_W_20190501_f3 | `=TEXT(DATE(2019,5,1),"e")` | 1 | 1 | 未検証 | e | 一致 |  |
| TEXT | TEXT_W_20190501_f4 | `=TEXT(DATE(2019,5,1),"ee")` | 01 | 01 | 未検証 | ee | 一致 |  |
| TEXT | TEXT_W_20190501_f5 | `=TEXT(DATE(2019,5,1),"ge.m.d")` | R1.5.1 | R1.5.1 | 未検証 | ge.5.1 | 一致 |  |
| TEXT | TEXT_W_20190501_f6 | `=TEXT(DATE(2019,5,1),"ggge年m月d日")` | 令和1年5月1日 | 令和1年5月1日 | 未検証 | ggge年5月1日 | 一致 |  |
| TEXT | TEXT_W_20190501_f7 | `=TEXT(DATE(2019,5,1),"gggee年")` | 令和01年 | 令和01年 | 未検証 | gggee年 | 一致 |  |
| TEXT | TEXT_W_20190501_f8 | `=TEXT(DATE(2019,5,1),"[$-ja-JP]ggge年")` | 令和1年 | 令和1年 | 未検証 | [$-ja-JP]ggge年 | 一致 |  |
| TEXT | TEXT_W_20200101_f0 | `=TEXT(DATE(2020,1,1),"g")` | R | R | 未検証 | g | 一致 |  |
| TEXT | TEXT_W_20200101_f1 | `=TEXT(DATE(2020,1,1),"gg")` | 令 | 令 | 未検証 | gg | 一致 |  |
| TEXT | TEXT_W_20200101_f2 | `=TEXT(DATE(2020,1,1),"ggg")` | 令和 | 令和 | 未検証 | ggg | 一致 |  |
| TEXT | TEXT_W_20200101_f3 | `=TEXT(DATE(2020,1,1),"e")` | 2 | 2 | 未検証 | e | 一致 |  |
| TEXT | TEXT_W_20200101_f4 | `=TEXT(DATE(2020,1,1),"ee")` | 02 | 02 | 未検証 | ee | 一致 |  |
| TEXT | TEXT_W_20200101_f5 | `=TEXT(DATE(2020,1,1),"ge.m.d")` | R2.1.1 | R2.1.1 | 未検証 | ge.1.1 | 一致 |  |
| TEXT | TEXT_W_20200101_f6 | `=TEXT(DATE(2020,1,1),"ggge年m月d日")` | 令和2年1月1日 | 令和2年1月1日 | 未検証 | ggge年1月1日 | 一致 |  |
| TEXT | TEXT_W_20200101_f7 | `=TEXT(DATE(2020,1,1),"gggee年")` | 令和02年 | 令和02年 | 未検証 | gggee年 | 一致 |  |
| TEXT | TEXT_W_20200101_f8 | `=TEXT(DATE(2020,1,1),"[$-ja-JP]ggge年")` | 令和2年 | 令和2年 | 未検証 | [$-ja-JP]ggge年 | 一致 |  |
| TEXT | TEXT_W_20260131_f0 | `=TEXT(DATE(2026,1,31),"g")` | R | R | 未検証 | g | 一致 |  |
| TEXT | TEXT_W_20260131_f1 | `=TEXT(DATE(2026,1,31),"gg")` | 令 | 令 | 未検証 | gg | 一致 |  |
| TEXT | TEXT_W_20260131_f2 | `=TEXT(DATE(2026,1,31),"ggg")` | 令和 | 令和 | 未検証 | ggg | 一致 |  |
| TEXT | TEXT_W_20260131_f3 | `=TEXT(DATE(2026,1,31),"e")` | 8 | 8 | 未検証 | e | 一致 |  |
| TEXT | TEXT_W_20260131_f4 | `=TEXT(DATE(2026,1,31),"ee")` | 08 | 08 | 未検証 | ee | 一致 |  |
| TEXT | TEXT_W_20260131_f5 | `=TEXT(DATE(2026,1,31),"ge.m.d")` | R8.1.31 | R8.1.31 | 未検証 | ge.1.31 | 一致 |  |
| TEXT | TEXT_W_20260131_f6 | `=TEXT(DATE(2026,1,31),"ggge年m月d日")` | 令和8年1月31日 | 令和8年1月31日 | 未検証 | ggge年1月31日 | 一致 |  |
| TEXT | TEXT_W_20260131_f7 | `=TEXT(DATE(2026,1,31),"gggee年")` | 令和08年 | 令和08年 | 未検証 | gggee年 | 一致 |  |
| TEXT | TEXT_W_20260131_f8 | `=TEXT(DATE(2026,1,31),"[$-ja-JP]ggge年")` | 令和8年 | 令和8年 | 未検証 | [$-ja-JP]ggge年 | 一致 |  |
| TEXT | TEXT_K_0_f0 | `=TEXT(0,"[h]:mm")` | 0:00 | 0:00 | 未検証 | [0]:00 | 一致 |  |
| TEXT | TEXT_K_0_f1 | `=TEXT(0,"[hh]:mm")` | 00:00 | 00:00 | 未検証 | 00:00 | 一致 |  |
| TEXT | TEXT_K_0_f2 | `=TEXT(0,"[h]:mm:ss")` | 0:00:00 | 0:00:00 | 未検証 | [0]:00:00 | 一致 |  |
| TEXT | TEXT_K_0_f3 | `=TEXT(0,"[m]:ss")` | 0:00 | 0:00 | 未検証 | [12]:00 | 一致 |  |
| TEXT | TEXT_K_0_f4 | `=TEXT(0,"[mm]:ss")` | 00:00 | 00:00 | 未検証 | 00:00 | 一致 |  |
| TEXT | TEXT_K_0_f5 | `=TEXT(0,"[s]")` | 0 | 0 | 未検証 | [0] | 一致 |  |
| TEXT | TEXT_K_0_f6 | `=TEXT(0,"[h]")` | 0 | 0 | 未検証 | [0] | 一致 |  |
| TEXT | TEXT_K_0_f7 | `=TEXT(0,"[h]:mm:ss.00")` | 0:00:00.00 | 0:00:00.00 | 未検証 | [0]:00:00 | 一致 |  |
| TEXT | TEXT_K_0p5_f0 | `=TEXT(0.5,"[h]:mm")` | 12:00 | 12:00 | 未検証 | [12]:00 | 一致 |  |
| TEXT | TEXT_K_0p5_f1 | `=TEXT(0.5,"[hh]:mm")` | 12:00 | 12:00 | 未検証 | 12:00 | 一致 |  |
| TEXT | TEXT_K_0p5_f2 | `=TEXT(0.5,"[h]:mm:ss")` | 12:00:00 | 12:00:00 | 未検証 | [12]:00:00 | 一致 |  |
| TEXT | TEXT_K_0p5_f3 | `=TEXT(0.5,"[m]:ss")` | 720:00 | 720:00 | 未検証 | [12]:00 | 一致 |  |
| TEXT | TEXT_K_0p5_f4 | `=TEXT(0.5,"[mm]:ss")` | 720:00 | 720:00 | 未検証 | 720:00 | 一致 |  |
| TEXT | TEXT_K_0p5_f5 | `=TEXT(0.5,"[s]")` | 43200 | 43200 | 未検証 | [0] | 一致 |  |
| TEXT | TEXT_K_0p5_f6 | `=TEXT(0.5,"[h]")` | 12 | 12 | 未検証 | [12] | 一致 |  |
| TEXT | TEXT_K_0p5_f7 | `=TEXT(0.5,"[h]:mm:ss.00")` | 12:00:00.00 | 12:00:00.00 | 未検証 | [12]:00:00 | 一致 |  |
| TEXT | TEXT_K_1p5_f0 | `=TEXT(1.5,"[h]:mm")` | 36:00 | 36:00 | 未検証 | [12]:00 | 一致 |  |
| TEXT | TEXT_K_1p5_f1 | `=TEXT(1.5,"[hh]:mm")` | 36:00 | 36:00 | 未検証 | 36:00 | 一致 |  |
| TEXT | TEXT_K_1p5_f2 | `=TEXT(1.5,"[h]:mm:ss")` | 36:00:00 | 36:00:00 | 未検証 | [12]:00:00 | 一致 |  |
| TEXT | TEXT_K_1p5_f3 | `=TEXT(1.5,"[m]:ss")` | 2160:00 | 2160:00 | 未検証 | [12]:00 | 一致 |  |
| TEXT | TEXT_K_1p5_f4 | `=TEXT(1.5,"[mm]:ss")` | 2160:00 | 2160:00 | 未検証 | 2160:00 | 一致 |  |
| TEXT | TEXT_K_1p5_f5 | `=TEXT(1.5,"[s]")` | 129600 | 129600 | 未検証 | [0] | 一致 |  |
| TEXT | TEXT_K_1p5_f6 | `=TEXT(1.5,"[h]")` | 36 | 36 | 未検証 | [12] | 一致 |  |
| TEXT | TEXT_K_1p5_f7 | `=TEXT(1.5,"[h]:mm:ss.00")` | 36:00:00.00 | 36:00:00.00 | 未検証 | [12]:00:00 | 一致 |  |
| TEXT | TEXT_K_46234_f0 | `=TEXT(46234,"[h]:mm")` | 1109616:00 | 1109616:00 | 未検証 | [0]:00 | 一致 |  |
| TEXT | TEXT_K_46234_f1 | `=TEXT(46234,"[hh]:mm")` | 1109616:00 | 1109616:00 | 未検証 | 1109616:00 | 一致 |  |
| TEXT | TEXT_K_46234_f2 | `=TEXT(46234,"[h]:mm:ss")` | 1109616:00:00 | 1109616:00:00 | 未検証 | [0]:00:00 | 一致 |  |
| TEXT | TEXT_K_46234_f3 | `=TEXT(46234,"[m]:ss")` | 66576960:00 | 66576960:00 | 未検証 | [7]:00 | 一致 |  |
| TEXT | TEXT_K_46234_f4 | `=TEXT(46234,"[mm]:ss")` | 66576960:00 | 66576960:00 | 未検証 | 66576960:00 | 一致 |  |
| TEXT | TEXT_K_46234_f5 | `=TEXT(46234,"[s]")` | 3994617600 | 3994617600 | 未検証 | [0] | 一致 |  |
| TEXT | TEXT_K_46234_f6 | `=TEXT(46234,"[h]")` | 1109616 | 1109616 | 未検証 | [0] | 一致 |  |
| TEXT | TEXT_K_46234_f7 | `=TEXT(46234,"[h]:mm:ss.00")` | 1109616:00:00.00 | 1109616:00:00.00 | 未検証 | [0]:00:00 | 一致 |  |
| TEXT | TEXT_K_half_sec_f0 | `=TEXT(1/24/60*59.5/60,"[h]:mm")` | 0:00 | 0:00 | 未検証 | [0]:00 | 一致 |  |
| TEXT | TEXT_K_half_sec_f1 | `=TEXT(1/24/60*59.5/60,"[hh]:mm")` | 00:00 | 00:00 | 未検証 | 00:00 | 一致 |  |
| TEXT | TEXT_K_half_sec_f2 | `=TEXT(1/24/60*59.5/60,"[h]:mm:ss")` | 0:00:59 | 0:00:59 | 未検証 | [0]:00:59 | 一致 |  |
| TEXT | TEXT_K_half_sec_f3 | `=TEXT(1/24/60*59.5/60,"[m]:ss")` | 0:59 | 0:59 | 未検証 | [12]:59 | 一致 |  |
| TEXT | TEXT_K_half_sec_f4 | `=TEXT(1/24/60*59.5/60,"[mm]:ss")` | 00:59 | 00:59 | 未検証 | 00:59 | 一致 |  |
| TEXT | TEXT_K_half_sec_f5 | `=TEXT(1/24/60*59.5/60,"[s]")` | 59 | 59 | 未検証 | [59] | 一致 |  |
| TEXT | TEXT_K_half_sec_f6 | `=TEXT(1/24/60*59.5/60,"[h]")` | 0 | 0 | 未検証 | [0] | 一致 |  |
| TEXT | TEXT_K_half_sec_f7 | `=TEXT(1/24/60*59.5/60,"[h]:mm:ss.00")` | 0:00:59.50 | 0:00:59.50 | 未検証 | [0]:00:59.5 | 一致 |  |
| TEXT | TEXT_K_0p999999_f0 | `=TEXT(0.999999,"[h]:mm")` | 24:00 | 24:00 | 未検証 | [23]:59 | 一致 |  |
| TEXT | TEXT_K_0p999999_f1 | `=TEXT(0.999999,"[hh]:mm")` | 24:00 | 24:00 | 未検証 | 23:59 | 一致 |  |
| TEXT | TEXT_K_0p999999_f2 | `=TEXT(0.999999,"[h]:mm:ss")` | 24:00:00 | 24:00:00 | 未検証 | [23]:59:59 | 一致 |  |
| TEXT | TEXT_K_0p999999_f3 | `=TEXT(0.999999,"[m]:ss")` | 1440:00 | 1440:00 | 未検証 | [12]:59 | 一致 |  |
| TEXT | TEXT_K_0p999999_f4 | `=TEXT(0.999999,"[mm]:ss")` | 1440:00 | 1440:00 | 未検証 | 1439:59 | 一致 |  |
| TEXT | TEXT_K_0p999999_f5 | `=TEXT(0.999999,"[s]")` | 86400 | 86400 | 未検証 | [59] | 一致 |  |
| TEXT | TEXT_K_0p999999_f6 | `=TEXT(0.999999,"[h]")` | 24 | 24 | 未検証 | [23] | 一致 |  |
| TEXT | TEXT_K_0p999999_f7 | `=TEXT(0.999999,"[h]:mm:ss.00")` | 23:59:59.91 | 23:59:59.91 | 未検証 | [23]:59:59.91 | 一致 |  |
| TEXT | TEXT_K_m1p5_f0 | `=TEXT(-1.5,"[h]:mm")` | #VALUE! | #VALUE! | 未検証 | [-12]:00 | 一致 |  |
| TEXT | TEXT_K_m1p5_f1 | `=TEXT(-1.5,"[hh]:mm")` | #VALUE! | #VALUE! | 未検証 | -36:00 | 一致 |  |
| TEXT | TEXT_K_m1p5_f2 | `=TEXT(-1.5,"[h]:mm:ss")` | #VALUE! | #VALUE! | 未検証 | [-12]:00:00 | 一致 |  |
| TEXT | TEXT_K_m1p5_f3 | `=TEXT(-1.5,"[m]:ss")` | #VALUE! | #VALUE! | 未検証 | [12]:00 | 一致 |  |
| TEXT | TEXT_K_m1p5_f4 | `=TEXT(-1.5,"[mm]:ss")` | #VALUE! | #VALUE! | 未検証 | -2160:00 | 一致 |  |
| TEXT | TEXT_K_m1p5_f5 | `=TEXT(-1.5,"[s]")` | -129600 | -129600 | 未検証 | [0] | 一致 |  |
| TEXT | TEXT_K_m1p5_f6 | `=TEXT(-1.5,"[h]")` | -36 | -36 | 未検証 | [-12] | 一致 |  |
| TEXT | TEXT_K_m1p5_f7 | `=TEXT(-1.5,"[h]:mm:ss.00")` | #VALUE! | #VALUE! | 未検証 | [-12]:00:00 | 一致 |  |
| TEXT | TEXT_E_1em7_f0 | `=TEXT(1E-7,"[h]:mm:ss")` | 0:00:00 | 0:00:00 | 未検証 | #ERROR | 一致 |  |
| TEXT | TEXT_E_1em7_f1 | `=TEXT(1E-7,"[h]:mm:ss.00")` | 0:00:00.01 | 0:00:00.01 | 未検証 | #ERROR | 一致 |  |
| TEXT | TEXT_E_1em7_f2 | `=TEXT(1E-7,"[s]")` | 0 | 0 | 未検証 | #ERROR | 一致 |  |
| TEXT | TEXT_E_1em7_f3 | `=TEXT(1E-7,"ge.m.d")` | M33.1.0 | M33.1.0 | 未検証 | #ERROR | 一致 |  |
| TEXT | TEXT_E_m555em17_f0 | `=TEXT(-5.55E-17,"[h]:mm:ss")` | #VALUE! | #VALUE! | 未検証 | #ERROR | 一致 |  |
| TEXT | TEXT_E_m555em17_f1 | `=TEXT(-5.55E-17,"[h]:mm:ss.00")` | #VALUE! | #VALUE! | 未検証 | #ERROR | 一致 |  |
| TEXT | TEXT_E_m555em17_f2 | `=TEXT(-5.55E-17,"[s]")` | -0 | -0 | 未検証 | #ERROR | 一致 |  |
| TEXT | TEXT_E_m555em17_f3 | `=TEXT(-5.55E-17,"ge.m.d")` | #VALUE! | #VALUE! | 未検証 | #ERROR | 一致 |  |
| TEXT | TEXT_E_1ep21_f0 | `=TEXT(1E+21,"[h]:mm:ss")` | #VALUE! | #VALUE! | 未検証 | #ERROR | 一致 |  |
| TEXT | TEXT_E_1ep21_f1 | `=TEXT(1E+21,"[h]:mm:ss.00")` | #VALUE! | #VALUE! | 未検証 | #ERROR | 一致 |  |
| TEXT | TEXT_E_1ep21_f2 | `=TEXT(1E+21,"[s]")` | 86400000000000000000000000 | 86400000000000000000000000 | 未検証 | #ERROR | 一致 |  |
| TEXT | TEXT_E_1ep21_f3 | `=TEXT(1E+21,"ge.m.d")` | #VALUE! | #VALUE! | 未検証 | #ERROR | 一致 |  |
| TEXT | TEXT_E_hashuu_f0 | `=TEXT(0.1+0.2-0.3,"[h]:mm:ss")` | 0:00:00 | 0:00:00 | 未検証 | [0]:00:00 | 一致 |  |
| TEXT | TEXT_E_hashuu_f1 | `=TEXT(0.1+0.2-0.3,"[h]:mm:ss.00")` | 0:00:00.00 | 0:00:00.00 | 未検証 | [0]:00:00 | 一致 |  |
| TEXT | TEXT_E_hashuu_f2 | `=TEXT(0.1+0.2-0.3,"[s]")` | 0 | 0 | 未検証 | [0] | 一致 |  |
| TEXT | TEXT_E_hashuu_f3 | `=TEXT(0.1+0.2-0.3,"ge.m.d")` | M33.1.0 | M33.1.0 | 未検証 | ge.12.30 | 一致 |  |
| TEXT | TEXT_B_2958465_f0 | `=TEXT(2958465,"ge.m.d")` | R7981.12.31 | R7981.12.31 | 未検証 | ge.12.31 | 一致 |  |
| TEXT | TEXT_B_2958465_f1 | `=TEXT(2958465,"ggge年m月d日")` | 令和7981年12月31日 | 令和7981年12月31日 | 未検証 | ggge年12月31日 | 一致 |  |
| TEXT | TEXT_B_2958465_f2 | `=TEXT(2958465,"[h]:mm:ss")` | 71003160:00:00 | 71003160:00:00 | 未検証 | [0]:00:00 | 一致 |  |
| TEXT | TEXT_B_2958465_f3 | `=TEXT(2958465,"[h]")` | 71003160 | 71003160 | 未検証 | [0] | 一致 |  |
| TEXT | TEXT_B_2958465_f4 | `=TEXT(2958465,"[s]")` | 255611376000 | 255611376000 | 未検証 | [0] | 一致 |  |
| TEXT | TEXT_B_2958465_f5 | `=TEXT(2958465,"yyyy/m/d")` | 9999/12/31 | 9999/12/31 | 未検証 | 9999/12/31 | 一致 |  |
| TEXT | TEXT_B_2958465p99999_f0 | `=TEXT(2958465.99999,"ge.m.d")` | R7981.12.31 | R7981.12.31 | 未検証 | ge.12.31 | 一致 |  |
| TEXT | TEXT_B_2958465p99999_f1 | `=TEXT(2958465.99999,"ggge年m月d日")` | 令和7981年12月31日 | 令和7981年12月31日 | 未検証 | ggge年12月31日 | 一致 |  |
| TEXT | TEXT_B_2958465p99999_f2 | `=TEXT(2958465.99999,"[h]:mm:ss")` | 71003183:59:59 | 71003183:59:59 | 未検証 | [23]:59:59 | 一致 |  |
| TEXT | TEXT_B_2958465p99999_f3 | `=TEXT(2958465.99999,"[h]")` | 71003183 | 71003183 | 未検証 | [23] | 一致 |  |
| TEXT | TEXT_B_2958465p99999_f4 | `=TEXT(2958465.99999,"[s]")` | 255611462399 | 255611462399 | 未検証 | [59] | 一致 |  |
| TEXT | TEXT_B_2958465p99999_f5 | `=TEXT(2958465.99999,"yyyy/m/d")` | 9999/12/31 | 9999/12/31 | 未検証 | 9999/12/31 | 一致 |  |
| TEXT | TEXT_B_2958466_f0 | `=TEXT(2958466,"ge.m.d")` | #VALUE! | #VALUE! | 未検証 | ge.1.1 | 一致 |  |
| TEXT | TEXT_B_2958466_f1 | `=TEXT(2958466,"ggge年m月d日")` | #VALUE! | #VALUE! | 未検証 | ggge年1月1日 | 一致 |  |
| TEXT | TEXT_B_2958466_f2 | `=TEXT(2958466,"[h]:mm:ss")` | #VALUE! | #VALUE! | 未検証 | [0]:00:00 | 一致 |  |
| TEXT | TEXT_B_2958466_f3 | `=TEXT(2958466,"[h]")` | 71003184 | 71003184 | 未検証 | [0] | 一致 |  |
| TEXT | TEXT_B_2958466_f4 | `=TEXT(2958466,"[s]")` | 255611462400 | 255611462400 | 未検証 | [0] | 一致 |  |
| TEXT | TEXT_B_2958466_f5 | `=TEXT(2958466,"yyyy/m/d")` | #VALUE! | #VALUE! | 未検証 | 10000/1/1 | 一致 |  |
| TEXT | TEXT_B_2958466p5_f0 | `=TEXT(2958466.5,"ge.m.d")` | #VALUE! | #VALUE! | 未検証 | ge.1.1 | 一致 |  |
| TEXT | TEXT_B_2958466p5_f1 | `=TEXT(2958466.5,"ggge年m月d日")` | #VALUE! | #VALUE! | 未検証 | ggge年1月1日 | 一致 |  |
| TEXT | TEXT_B_2958466p5_f2 | `=TEXT(2958466.5,"[h]:mm:ss")` | #VALUE! | #VALUE! | 未検証 | [12]:00:00 | 一致 |  |
| TEXT | TEXT_B_2958466p5_f3 | `=TEXT(2958466.5,"[h]")` | 71003196 | 71003196 | 未検証 | [12] | 一致 |  |
| TEXT | TEXT_B_2958466p5_f4 | `=TEXT(2958466.5,"[s]")` | 255611505600 | 255611505600 | 未検証 | [0] | 一致 |  |
| TEXT | TEXT_B_2958466p5_f5 | `=TEXT(2958466.5,"yyyy/m/d")` | #VALUE! | #VALUE! | 未検証 | 10000/1/1 | 一致 |  |
| SORT | SORT_bare | `=SORT(E1:E6)` | 100 | 100 | 未検証 | 100 | 一致 |  |
| UNIQUE | UNIQUE_bare | `=UNIQUE(D1:D6)` | A | A | 未検証 | A | 一致 |  |
| FILTER | FILTER_bare | `=FILTER(E1:E6,D1:D6="A")` | 100 | 100 | 未検証 | 100 | 一致 |  |
| SORT | SORT_empty | `=SORT(Z1:Z3)` | 0 | 0 | 未検証 | (null) | 一致 |  |
| UNIQUE | UNIQUE_empty | `=UNIQUE(Z1:Z3)` | 0 | 0 | 未検証 | (null) | 一致 |  |
| FILTER | FILTER_nomatch | `=IFERROR(FILTER(E1:E6,E1:E6>9999),"NA")` | NA | NA | 未検証 | NA | 一致 |  |
| SORT | SORT_nested_sum | `=SUM(SORT(E1:E6))` | 2100 | 2100 | 未検証 | 2100 | 一致 |  |
| SORT | SORT_nested_index2 | `=INDEX(SORT(E1:E6,1,-1),2)` | 500 | 500 | 未検証 | 500 | 一致 |  |
| UNIQUE | UNIQUE_nested_sum | `=SUM(UNIQUE(E1:E6))` | 2100 | 2100 | 未検証 | 2100 | 一致 |  |
| UNIQUE | UNIQUE_nested_count | `=COUNT(UNIQUE(D1:D6))` | 0 | 0 | 未検証 | 0 | 一致 |  |
| FILTER | FILTER_nested_count | `=COUNT(FILTER(E1:E6,D1:D6="A"))` | 3 | 3 | 未検証 | 3 | 一致 |  |
| SORT | SORT_asc_join | `=TEXTJOIN(",",TRUE,SORT(E1:E6))` | 100,200,300,400,500,600 | 100,200,300,400,500,600 | 未検証 | #NAME? | 一致 |  |
| SORT | SORT_desc_join | `=TEXTJOIN(",",TRUE,SORT(E1:E6,1,-1))` | 600,500,400,300,200,100 | 600,500,400,300,200,100 | 未検証 | #NAME? | 一致 |  |
| SORT | SORT_first | `=INDEX(SORT(E1:E6,1,-1),1)` | 600 | 600 | 未検証 | 600 | 一致 |  |
| SORT | SORT_text | `=TEXTJOIN(",",TRUE,SORT(D1:D6))` | A,A,A,B,B,C | A,A,A,B,B,C | 未検証 | #NAME? | 一致 |  |
| SORT | SORT_count | `=COUNTA(SORT(E1:E6))` | 6 | 6 | 未検証 | 6 | 一致 |  |
| UNIQUE | UNIQUE_join | `=TEXTJOIN(",",TRUE,UNIQUE(D1:D6))` | A,B,C | A,B,C | 未検証 | #NAME? | 一致 |  |
| UNIQUE | UNIQUE_count | `=COUNTA(UNIQUE(D1:D6))` | 3 | 3 | 未検証 | 3 | 一致 |  |
| UNIQUE | UNIQUE_nums | `=TEXTJOIN(",",TRUE,UNIQUE(E1:E6))` | 100,200,300,400,500,600 | 100,200,300,400,500,600 | 未検証 | #NAME? | 一致 |  |
| UNIQUE | UNIQUE_all_same | `=COUNTA(UNIQUE(D1:D1))` | 1 | 1 | 未検証 | 1 | 一致 |  |
| FILTER | FILTER_sum | `=SUM(FILTER(E1:E6,D1:D6="A"))` | 1000 | 1000 | 未検証 | 1000 | 一致 |  |
| FILTER | FILTER_join | `=TEXTJOIN(",",TRUE,FILTER(E1:E6,D1:D6="A"))` | 100,300,600 | 100,300,600 | 未検証 | #NAME? | 一致 |  |
| FILTER | FILTER_num_cond | `=TEXTJOIN(",",TRUE,FILTER(E1:E6,C1:C6>=10))` | 300,400,500,600 | 300,400,500,600 | 未検証 | #NAME? | 一致 |  |
| FILTER | FILTER_empty | `=IFERROR(TEXTJOIN(",",TRUE,FILTER(E1:E6,D1:D6="Z")),"NA")` | NA | NA | 未検証 | NA | 一致 |  |
| FILTER | FILTER_if_empty | `=TEXTJOIN(",",TRUE,FILTER(E1:E6,D1:D6="Z","なし"))` | なし | なし | 未検証 | #NAME? | 一致 |  |
| SORT | ARRAY_sort_unique | `=TEXTJOIN(",",TRUE,SORT(UNIQUE(D1:D6)))` | A,B,C | A,B,C | 未検証 | #NAME? | 一致 |  |
| 配列演算 | ARR_mul_text | `=D1:D6*E1:E6` | #VALUE! | #VALUE! | 未検証 | #VALUE! | 一致 |  |
| 配列演算 | ARR_mul_num | `=C1:C6*E1:E6` | 100 | 100 | 未検証 | #VALUE! | 一致 |  |
| 配列演算 | ARR_if_cond | `=IF(D1:D6="A",E1:E6,0)` | 100 | 100 | 未検証 | #VALUE! | 一致 |  |
| 配列演算 | ARR_sum_mul_text | `=IFERROR(SUM(D1:D6*E1:E6),"ERR")` | ERR | ERR | 未検証 | ERR | 一致 |  |
| 配列演算 | ARR_sum_mul_num | `=SUM(C1:C6*E1:E6)` | 97100 | 97100 | 未検証 | #VALUE! | 一致 |  |
| 配列演算 | ARR_sum_if | `=SUM(IF(D1:D6="A",E1:E6,0))` | 1000 | 1000 | 未検証 | #VALUE! | 一致 |  |
| 配列演算 | ARR_count_if | `=COUNT(IF(D1:D6="A",E1:E6,0))` | 6 | 6 | 未検証 | 0 | 一致 |  |
| 配列演算 | ARR_count_mul | `=COUNT(C1:C6*E1:E6)` | 6 | 6 | 未検証 | 0 | 一致 |  |
| CONCAT | CONCAT_range | `=CONCAT(D1:D3)` | ABA | ABA | 未検証 | #NAME? | 一致 |  |
| CONCAT | CONCAT_mixed | `=CONCAT(B1,"-",A1)` | りんご-1000 | りんご-1000 | 未検証 | #NAME? | 一致 |  |
| CONCAT | CONCAT_blank | `="["&CONCAT(G1:G3)&"]"` | [00] | [00] | 未検証 | #NAME? | 一致 |  |
| CONCAT | CONCAT_number_fmt | `=CONCAT(A4)` | 0.1 | 0.1 | 未検証 | #NAME? | 一致 |  |
| CONCAT | CONCAT_nested | `=LEN(CONCAT(D1:D6))` | 6 | 6 | 未検証 | #NAME? | 一致 |  |
| LOOKUP | LOOKUP_exact | `=LOOKUP(20,C1:C6,E1:E6)` | 400 | 400 | 未検証 | #NAME? | 一致 |  |
| LOOKUP | LOOKUP_between | `=LOOKUP(30,C1:C6,E1:E6)` | 400 | 400 | 未検証 | #NAME? | 一致 |  |
| LOOKUP | LOOKUP_below_all | `=IFERROR(LOOKUP(0,C1:C6,E1:E6),"NA")` | NA | NA | 未検証 | NA | 一致 |  |
| LOOKUP | LOOKUP_above_all | `=LOOKUP(999,C1:C6,E1:E6)` | 600 | 600 | 未検証 | #NAME? | 一致 |  |
| LOOKUP | LOOKUP_text | `=LOOKUP("B",{"A","B","C"},{100,200,300})` | 200 | 200 | 未検証 | #NAME? | 一致 |  |
| LOOKUP | LOOKUP_nested | `=ROUND(LOOKUP(30,C1:C6,E1:E6),0)` | 400 | 400 | 未検証 | #NAME? | 一致 |  |
| XMATCH | XMATCH_exact | `=XMATCH(20,C1:C6)` | 4 | 4 | 未検証 | #NAME? | 一致 |  |
| XMATCH | XMATCH_miss | `=IFERROR(XMATCH(30,C1:C6),"NA")` | NA | NA | 未検証 | NA | 一致 |  |
| XMATCH | XMATCH_next_smaller | `=XMATCH(30,C1:C6,-1)` | 4 | 4 | 未検証 | #NAME? | 一致 |  |
| XMATCH | XMATCH_next_larger | `=XMATCH(30,C1:C6,1)` | 5 | 5 | 未検証 | #NAME? | 一致 |  |
| XMATCH | XMATCH_wildcard | `=XMATCH("りん*",B1:B8,2)` | 1 | 1 | 未検証 | #NAME? | 一致 |  |
| XMATCH | XMATCH_reverse | `=XMATCH("A",D1:D6,0,-1)` | 6 | 6 | 未検証 | #NAME? | 一致 |  |
| XMATCH | XMATCH_nested | `=INDEX(E1:E6,XMATCH(20,C1:C6))` | 400 | 400 | 未検証 | #NAME? | 一致 |  |
| INDIRECT | INDIRECT_cell | `=INDIRECT("E1")` | 100 | 100 | 未検証 | #NAME? | 一致 |  |
| INDIRECT | INDIRECT_from_cell | `=INDIRECT(B12)` | 200 | 200 | 未検証 | #NAME? | 一致 |  |
| INDIRECT | INDIRECT_range_sum | `=SUM(INDIRECT("E1:E6"))` | 2100 | 2100 | 未検証 | #NAME? | 一致 |  |
| INDIRECT | INDIRECT_bad | `=IFERROR(INDIRECT("あ"),"NA")` | NA | NA | 未検証 | NA | 一致 |  |
| INDIRECT | INDIRECT_nested | `=ROUND(SUM(INDIRECT("E1:E6"))/2,0)` | 1050 | 1050 | 未検証 | #NAME? | 一致 |  |
| DATEVALUE | DATEVALUE_slash | `=DATEVALUE("2026/7/31")` | 46234 | 46234 | 未検証 | #VALUE! | 一致 |  |
| DATEVALUE | DATEVALUE_hyphen | `=DATEVALUE("2026-07-31")` | 46234 | 46234 | 未検証 | #VALUE! | 一致 |  |
| DATEVALUE | DATEVALUE_jp | `=IFERROR(DATEVALUE("2026年7月31日"),"NA")` | 46234 | 46234 | 未検証 | NA | 一致 |  |
| DATEVALUE | DATEVALUE_bad | `=IFERROR(DATEVALUE("あ"),"NA")` | NA | NA | 未検証 | NA | 一致 |  |
| DATEVALUE | DATEVALUE_nested | `=YEAR(DATEVALUE("2026-07-31"))` | 2026 | 2026 | 未検証 | #VALUE! | 一致 |  |
| NUMBERVALUE | NUMBERVALUE_plain | `=NUMBERVALUE("1.5")` | 1.5 | 1.5 | 未検証 | #NAME? | 一致 |  |
| NUMBERVALUE | NUMBERVALUE_sep | `=NUMBERVALUE("1,234.5")` | 1234.5 | 1234.5 | 未検証 | #NAME? | 一致 |  |
| NUMBERVALUE | NUMBERVALUE_custom | `=NUMBERVALUE("1.234,5",",",".")` | 1234.5 | 1234.5 | 未検証 | #NAME? | 一致 |  |
| NUMBERVALUE | NUMBERVALUE_bad | `=IFERROR(NUMBERVALUE("あ"),"NA")` | NA | NA | 未検証 | NA | 一致 |  |
| NUMBERVALUE | NUMBERVALUE_nested | `=SUM(NUMBERVALUE("1,234"),1)` | 1235 | 1235 | 未検証 | #NAME? | 一致 |  |
| FIXED | FIXED_default | `=FIXED(1234.567)` | 1,234.57 | 1,234.57 | 未検証 | #NAME? | 一致 |  |
| FIXED | FIXED_digits0 | `=FIXED(1234.5,0)` | 1,235 | 1,235 | 未検証 | #NAME? | 一致 |  |
| FIXED | FIXED_nocomma | `=FIXED(1234.567,2,TRUE)` | 1234.57 | 1234.57 | 未検証 | #NAME? | 一致 |  |
| FIXED | FIXED_negative_digits | `=FIXED(1234.5,-2)` | 1,200 | 1,200 | 未検証 | #NAME? | 一致 |  |
| FIXED | FIXED_is_text | `=LEN(FIXED(1234.5,0))` | 5 | 5 | 未検証 | #NAME? | 一致 |  |
| ASC | ASC_alnum | `=ASC(B9)` | ABC123 | ABC123 | 未検証 | #NAME? | 一致 |  |
| ASC | ASC_kana | `=ASC("アイウ")` | ｱｲｳ | ｱｲｳ | 未検証 | #NAME? | 一致 |  |
| ASC | ASC_mixed | `=ASC("Ａ亜１")` | A亜1 | A亜1 | 未検証 | #NAME? | 一致 |  |
| DBCS | DBCS_alnum | `=DBCS("ABC123")` | ＡＢＣ１２３ | ＡＢＣ１２３ | 未検証 | #NAME? | 一致 |  |
| DBCS | DBCS_kana | `=DBCS(B10)` | アイウ | アイウ | 未検証 | #NAME? | 一致 |  |
| DBCS | DBCS_dakuten | `=DBCS(B11)` | ガギ | ガギ | 未検証 | #NAME? | 一致 |  |
| DBCS | DBCS_nested | `=LEN(DBCS(B10))` | 3 | 3 | 未検証 | #NAME? | 一致 |  |
| ASC | ASC_nested | `=LEN(ASC(B9))` | 6 | 6 | 未検証 | #NAME? | 一致 |  |
| TEXTBEFORE | TEXTBEFORE_first | `=TEXTBEFORE("007-1234","-")` | 007 | 007 | 未検証 | #NAME? | 一致 |  |
| TEXTAFTER | TEXTAFTER_first | `=TEXTAFTER("007-1234","-")` | 1234 | 1234 | 未検証 | #NAME? | 一致 |  |
| TEXTBEFORE | TEXTBEFORE_nth | `=TEXTBEFORE("a-b-c","-",2)` | a-b | a-b | 未検証 | #NAME? | 一致 |  |
| TEXTAFTER | TEXTAFTER_negative | `=TEXTAFTER("a-b-c","-",-1)` | c | c | 未検証 | #NAME? | 一致 |  |
| TEXTBEFORE | TEXTBEFORE_missing | `=IFERROR(TEXTBEFORE("abc","-"),"NA")` | NA | NA | 未検証 | NA | 一致 |  |
| TEXTBEFORE | TEXTBEFORE_ifmissing | `=TEXTBEFORE("abc","-",1,0,0,"なし")` | なし | なし | 未検証 | #NAME? | 一致 |  |
| TEXTAFTER | TEXTAFTER_nested | `=LEN(TEXTAFTER("007-1234","-"))` | 4 | 4 | 未検証 | #NAME? | 一致 |  |
| DOLLAR | DOLLAR_default | `=DOLLAR(1234.567)` | ¥1,235 | ¥1,235 | 未検証 | #NAME? | 一致 |  |
| DOLLAR | DOLLAR_digits0 | `=DOLLAR(1234.5,0)` | ¥1,235 | ¥1,235 | 未検証 | #NAME? | 一致 |  |
| DOLLAR | DOLLAR_digits1 | `=DOLLAR(1234.5,1)` | ¥1,234.5 | ¥1,234.5 | 未検証 | #NAME? | 一致 |  |
| DOLLAR | DOLLAR_digits2 | `=DOLLAR(1234.567,2)` | ¥1,234.57 | ¥1,234.57 | 未検証 | #NAME? | 一致 |  |
| DOLLAR | DOLLAR_negative | `=DOLLAR(-1234.5)` | ¥-1,235 | ¥-1,235 | 未検証 | #NAME? | 一致 |  |
| DOLLAR | DOLLAR_negative2 | `=DOLLAR(-1234.5,2)` | ¥-1,234.50 | ¥-1,234.50 | 未検証 | #NAME? | 一致 |  |
| DOLLAR | DOLLAR_neg_digits | `=DOLLAR(1234.5,-2)` | ¥1,200 | ¥1,200 | 未検証 | #NAME? | 一致 |  |
| DOLLAR | DOLLAR_nested | `=LEN(DOLLAR(1000,0))` | 6 | 6 | 未検証 | #NAME? | 一致 |  |
| TYPE | TYPE_number | `=TYPE(A1)` | 1 | 1 | 未検証 | #NAME? | 一致 |  |
| TYPE | TYPE_text | `=TYPE(B1)` | 2 | 2 | 未検証 | #NAME? | 一致 |  |
| TYPE | TYPE_logical | `=TYPE(H1)` | 4 | 4 | 未検証 | #NAME? | 一致 |  |
| TYPE | TYPE_error | `=TYPE(1/0)` | 16 | 16 | 未検証 | #NAME? | 一致 |  |
| TYPE | TYPE_array | `=TYPE({1,2})` | 64 | 64 | 未検証 | #NAME? | 一致 |  |
| TYPE | TYPE_nested | `=TYPE(A1)+TYPE(B1)` | 3 | 3 | 未検証 | #NAME? | 一致 |  |
| AGGREGATE | AGG_sum | `=AGGREGATE(9,0,E1:E6)` | 2100 | 2100 | 未検証 | #NAME? | 一致 |  |
| AGGREGATE | AGG_average | `=AGGREGATE(1,0,E1:E6)` | 350 | 350 | 未検証 | #NAME? | 一致 |  |
| AGGREGATE | AGG_max | `=AGGREGATE(4,0,E1:E6)` | 600 | 600 | 未検証 | #NAME? | 一致 |  |
| AGGREGATE | AGG_count | `=AGGREGATE(2,0,E1:E6)` | 6 | 6 | 未検証 | #NAME? | 一致 |  |
| AGGREGATE | AGG_large_k | `=AGGREGATE(14,0,E1:E6,2)` | 500 | 500 | 未検証 | #NAME? | 一致 |  |
| AGGREGATE | AGG_small_k | `=AGGREGATE(15,0,E1:E6,2)` | 200 | 200 | 未検証 | #NAME? | 一致 |  |
| AGGREGATE | AGG_ignore_err | `=AGGREGATE(9,6,C1:C6)` | 186 | 186 | 未検証 | #NAME? | 一致 |  |
| AGGREGATE | AGG_median | `=AGGREGATE(12,0,E1:E6)` | 350 | 350 | 未検証 | #NAME? | 一致 |  |
| AGGREGATE | AGG_nested | `=ROUND(AGGREGATE(1,0,E1:E6),0)` | 350 | 350 | 未検証 | #NAME? | 一致 |  |
| LENB | LENB_jp | `=LENB(B1)` | 6 | 6 | 未検証 | #NAME? | 一致 |  |
| LENB | LENB_ascii | `=LENB(B3)` | 5 | 5 | 未検証 | #NAME? | 一致 |  |
| LENB | LENB_mixed | `=LENB(B8)` | 9 | 9 | 未検証 | #NAME? | 一致 |  |
| LENB | LENB_halfkana | `=LENB(B10)` | 3 | 3 | 未検証 | #NAME? | 一致 |  |
| LENB | LENB_nested | `=LENB(B1)+LENB(B3)` | 11 | 11 | 未検証 | #NAME? | 一致 |  |
| LEFTB | LEFTB_jp | `=LEFTB(B1,2)` | り | り | 未検証 | #NAME? | 一致 |  |
| LEFTB | LEFTB_odd | `=LEFTB(B1,3)` | り  | り  | 未検証 | #NAME? | 一致 |  |
| LEFTB | LEFTB_zero | `="["&LEFTB(B1,0)&"]"` | [] | [] | 未検証 | #NAME? | 一致 |  |
| LEFTB | LEFTB_nested | `=LEN(LEFTB(B1,4))` | 2 | 2 | 未検証 | #NAME? | 一致 |  |
| RIGHTB | RIGHTB_jp | `=RIGHTB(B1,2)` | ご | ご | 未検証 | #NAME? | 一致 |  |
| RIGHTB | RIGHTB_odd | `=RIGHTB(B1,3)` |  ご |  ご | 未検証 | #NAME? | 一致 |  |
| RIGHTB | RIGHTB_ascii | `=RIGHTB(B3,3)` | ple | ple | 未検証 | #NAME? | 一致 |  |
| MIDB | MIDB_jp | `=MIDB(B1,3,2)` | ん | ん | 未検証 | #NAME? | 一致 |  |
| MIDB | MIDB_odd | `=MIDB(B1,2,2)` |    |    | 未検証 | #NAME? | 一致 |  |
| MIDB | MIDB_nested | `=LEN(MIDB(B1,1,4))` | 2 | 2 | 未検証 | #NAME? | 一致 |  |
| RANK | RANK_desc | `=RANK(20,C1:C6,0)` | 3 | 3 | 未検証 | #NAME? | 一致 |  |
| RANK | RANK_asc | `=RANK(20,C1:C6,1)` | 4 | 4 | 未検証 | #NAME? | 一致 |  |
| RANK | RANK_default_order | `=RANK(20,C1:C6)` | 3 | 3 | 未検証 | #NAME? | 一致 |  |
| RANK.EQ | RANK_eq | `=RANK.EQ(20,C1:C6,0)` | 3 | 3 | 未検証 | #NAME? | 一致 |  |
| RANK.AVG | RANK_avg | `=RANK.AVG(20,C1:C6,0)` | 3 | 3 | 未検証 | #NAME? | 一致 |  |
| RANK | RANK_miss | `=IFERROR(RANK(30,C1:C6,0),"NA")` | NA | NA | 未検証 | NA | 一致 |  |
| RANK | RANK_nested | `=INDEX(E1:E6,RANK(20,C1:C6,1))` | 400 | 400 | 未検証 | #NAME? | 一致 |  |
| VALUETOTEXT | VTT_number | `=VALUETOTEXT(A1)` | 1000 | 1000 | 未検証 | #NAME? | 一致 |  |
| VALUETOTEXT | VTT_text | `=VALUETOTEXT(B1)` | りんご | りんご | 未検証 | #NAME? | 一致 |  |
| VALUETOTEXT | VTT_strict | `=VALUETOTEXT(B1,1)` | "りんご" | "りんご" | 未検証 | #NAME? | 一致 |  |
| VALUETOTEXT | VTT_logical | `=VALUETOTEXT(H1)` | TRUE | TRUE | 未検証 | #NAME? | 一致 |  |
| VALUETOTEXT | VTT_nested | `=LEN(VALUETOTEXT(A1))` | 4 | 4 | 未検証 | #NAME? | 一致 |  |
| ENCODEURL | ENC_space | `=ENCODEURL("a b")` | a%20b | a%20b | 未検証 | #NAME? | 一致 |  |
| ENCODEURL | ENC_jp | `=ENCODEURL(B1)` | %E3%82%8A%E3%82%93%E3%81%94 | %E3%82%8A%E3%82%93%E3%81%94 | 未検証 | #NAME? | 一致 |  |
| ENCODEURL | ENC_symbols | `=ENCODEURL("a/b?c=1&d")` | a%2Fb%3Fc%3D1%26d | a%2Fb%3Fc%3D1%26d | 未検証 | #NAME? | 一致 |  |
| ENCODEURL | ENC_specials | `=ENCODEURL("!'()*-_.~")` | %21%27%28%29%2A-_.%7E | %21%27%28%29%2A-_.%7E | 未検証 | #NAME? | 一致 |  |
| ENCODEURL | ENC_nested | `=LEN(ENCODEURL("a b"))` | 5 | 5 | 未検証 | #NAME? | 一致 |  |
| AGGREGATE | AGG_mode | `=IFERROR(AGGREGATE(13,0,E1:E6),"NA")` | NA | NA | 未検証 | NA | 一致 |  |
| MODE | MODE_bare | `=MODE(I1:I6)` | 4 | 4 | 未検証 | #NAME? | 一致 |  |
| MODE | MODE_none | `=IFERROR(MODE(C1:C6),"NA")` | NA | NA | 未検証 | NA | 一致 |  |
| MODE.SNGL | MODE_sngl | `=MODE.SNGL(I1:I6)` | 4 | 4 | 未検証 | #NAME? | 一致 |  |
| MODE | MODE_nested | `=ROUND(MODE(I1:I6),0)` | 4 | 4 | 未検証 | #NAME? | 一致 |  |
| TRIMMEAN | TRIMMEAN_bare | `=TRIMMEAN(I1:I6,0.4)` | 4.25 | 4.25 | 未検証 | #NAME? | 一致 |  |
| TRIMMEAN | TRIMMEAN_zero | `=TRIMMEAN(I1:I6,0)` | 4.666666666666667 | 4.666666666666667 | 未検証 | #NAME? | 一致 |  |
| TRIMMEAN | TRIMMEAN_nested | `=ROUND(TRIMMEAN(I1:I6,0.4),2)` | 4.25 | 4.25 | 未検証 | #NAME? | 一致 |  |
| PERCENTRANK | PERCENTRANK_bare | `=PERCENTRANK(I1:I6,4)` | 0.2 | 0.2 | 未検証 | #NAME? | 一致 |  |
| PERCENTRANK | PERCENTRANK_top | `=PERCENTRANK(I1:I6,9)` | 1 | 1 | 未検証 | #NAME? | 一致 |  |
| PERCENTRANK | PERCENTRANK_sig | `=PERCENTRANK(I1:I6,4,5)` | 0.2 | 0.2 | 未検証 | #NAME? | 一致 |  |
| PERCENTRANK | PERCENTRANK_nested | `=ROUND(PERCENTRANK(I1:I6,4)*100,0)` | 20 | 20 | 未検証 | #NAME? | 一致 |  |
| KURT | KURT_bare | `=KURT(I1:I6)` | 3.2079119571683474 | 3.2079119571683474 | 未検証 | #NAME? | 一致 |  |
| KURT | KURT_nested | `=ROUND(KURT(I1:I6),4)` | 3.2079 | 3.2079 | 未検証 | #NAME? | 一致 |  |
| INTERCEPT | INTERCEPT_bare | `=INTERCEPT(E1:E6,C1:C6)` | 213.36088154269973 | 213.36088154269973 | 未検証 | #NAME? | 一致 |  |
| INTERCEPT | INTERCEPT_nested | `=ROUND(INTERCEPT(E1:E6,C1:C6),4)` | 213.3609 | 213.3609 | 未検証 | #NAME? | 一致 |  |
| FORECAST | FORECAST_bare | `=FORECAST(30,E1:E6,C1:C6)` | 345.59228650137743 | 345.59228650137743 | 未検証 | #NAME? | 一致 |  |
| FORECAST.LINEAR | FORECAST_linear | `=FORECAST.LINEAR(30,E1:E6,C1:C6)` | 345.59228650137743 | 345.59228650137743 | 未検証 | #NAME? | 一致 |  |
| FORECAST | FORECAST_nested | `=ROUND(FORECAST(30,E1:E6,C1:C6),2)` | 345.59 | 345.59 | 未検証 | #NAME? | 一致 |  |
| IRR | IRR_bare | `=IRR(J1:J5)` | 0.15322137877181535 | 0.15322137877181552 | 未検証 | 0.15322137877 | 一致 |  |
| IRR | IRR_guess | `=IRR(J1:J5,0.2)` | 0.15322137877181524 | 0.15322137876909325 | 未検証 | 0.15322137877 | 一致 |  |
| IRR | IRR_nested | `=ROUND(IRR(J1:J5)*100,2)` | 15.32 | 15.32 | 未検証 | 15.32 | 一致 |  |
| PERMUT | PERMUT_bare | `=PERMUT(6,3)` | 120 | 120 | 未検証 | #NAME? | 一致 |  |
| PERMUT | PERMUT_k0 | `=PERMUT(6,0)` | 1 | 1 | 未検証 | #NAME? | 一致 |  |
| PERMUT | PERMUT_nested | `=PERMUT(6,3)/6` | 20 | 20 | 未検証 | #NAME? | 一致 |  |
| PERMUTATIONA | PERMUTATIONA_bare | `=PERMUTATIONA(6,3)` | 216 | 216 | 未検証 | #NAME? | 一致 |  |
| PERMUTATIONA | PERMUTATIONA_nested | `=PERMUTATIONA(6,3)/6` | 36 | 36 | 未検証 | #NAME? | 一致 |  |
| MDETERM | MDETERM_bare | `=MDETERM(K1:L2)` | -2 | -2 | 未検証 | #NAME? | 一致 |  |
| MDETERM | MDETERM_nested | `=ROUND(MDETERM(K1:L2),0)` | -2 | -2 | 未検証 | #NAME? | 一致 |  |
| GESTEP | GESTEP_bare | `=GESTEP(5,4)` | 1 | 1 | 未検証 | #NAME? | 一致 |  |
| GESTEP | GESTEP_below | `=GESTEP(3,4)` | 0 | 0 | 未検証 | #NAME? | 一致 |  |
| GESTEP | GESTEP_nostep | `=GESTEP(5)` | 1 | 1 | 未検証 | #NAME? | 一致 |  |
| GESTEP | GESTEP_nested | `=SUM(GESTEP(5,4),GESTEP(3,4))` | 1 | 1 | 未検証 | #NAME? | 一致 |  |
