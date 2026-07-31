# 版対応 検証ハーネス レポート

- 生成日: 2026-07-31
- 真値: **O365HomePremRetail 16.0.20228.20124** (x64 / Current Channel)
- ロケール: UI=1041 / 国=81 / 小数点='.' 桁区切り=',' / 日付システム=1900
- 計算経路: **book.html の setCellFormula(本番と同じ)**。生の HyperFormula ではない。

> **★これは「今この瞬間の本番グリッドのお金バグ」である。** 下の「不一致(既知)」のうち money_impact の付いた項目は、
> 金額の端数・表示に直接効く。ハーネスを先に作る判断のため **このコミットでは直していない**（区分Aとして期限付きで台帳に登録）。

## 集計

| 判定 | 件数 |
|---|---|
| 一致 | 221 |
| 不一致(既知) | 14 |
| 不一致(新規) | 0 |
| 未検証 | 0 |
| 揮発性 | 2 |
| **合計** | 237 |

※ 「未検証」は緑ではない。その版の真値がまだ無い、という意味。

## 版ごとの状態

| 版 | 状態 |
|---|---|
| Excel 365 (16.0.20228.20124) | **真値**（このリポジトリの基準） |
| LibreOffice | **未検証**（goldenが無い。CIの別ジョブで生成する） |
| Excel 2016 / 2019 / Mac | **未検証**（実機が無い。golden/RECIPE.md の手順でその環境で1回走らせれば埋まる） |

## 揮発性の関数（別扱い）

TODAY / NOW は毎回答えが変わるため **golden突合の対象外**。固定値と比べると必ず腐るので、実行時点との一致だけを見る。

| ケース | 式 | 見方 | 期待 | 実際 | 判定 |
|---|---|---|---|---|---|
| TODAY_serial | `=TODAY()*1` | 実行時点の日付シリアルと一致するか | 46234 | 46234 | OK |
| NOW_int_is_today | `=INT(NOW())-TODAY()` | NOWの整数部がTODAYと一致するか | 0 | 0 | OK |

## 不一致（新規）＝赤

なし。

## 不一致（既知＝台帳にあり・緑だが必ず全件出す）

| 区分 | 関数 | ケース | 式 | Exally | Excel真値 | 中身と期限 |
|---|---|---|---|---|---|---|
| A | SUMPRODUCT | SUMPRODUCT_cond | `=SUMPRODUCT((D1:D6="A")*E1:E6)` | #VALUE! | 1000 | 条件×金額の定番形が動かない(SUMIFSの代わりに広く使われる形) / 期限 2026-08-31 |
| A | SUMPRODUCT | SUMPRODUCT_len | `=SUMPRODUCT(LEN(B1:B3))` | #VALUE! | 11 | 配列を返す関数を引数に取れない / 期限 2026-08-31 |
| A | MATCH | MATCH_wildcard | `=MATCH("りん*",B1:B8,0)` | #N/A | 1 | 部分一致の検索ができない / 期限 2026-08-31 |
| A | TEXT | TEXT_nested | `=LEN(TEXT(1234.5,"#,##0"))` | 8 | 5 | ★=LEN(TEXT(1234.5,"#,##0")) のように入れ子だと独自層が届かず、書式が効かないHFの結果(『1234.5』の8文字)になる。INT/MODは式に書き換えて根治したが、TEXTは書式処理を式で書けない。第2波で仕組みごと直す / 期限 2026-08-31 |
| A | SORT | SORT_asc_join | `=TEXTJOIN(",",TRUE,SORT(E1:E6))` | #NAME? | 100,200,300,400,500,600 | SORTが無い。★第1波でTEXTJOINの引数分解を直したので、壊れた文字列ではなくエラーが出るようになった / 期限 2026-08-31 |
| A | SORT | SORT_desc_join | `=TEXTJOIN(",",TRUE,SORT(E1:E6,1,-1))` | #NAME? | 600,500,400,300,200,100 | 同上(以前は『1,-1)』という無意味な文字列を黙って返していた) / 期限 2026-08-31 |
| A | SORT | SORT_first | `=INDEX(SORT(E1:E6,1,-1),1)` | #NAME? | 600 | SORTが無い / 期限 2026-08-31 |
| A | SORT | SORT_text | `=TEXTJOIN(",",TRUE,SORT(D1:D6))` | #NAME? | A,A,A,B,B,C | 同上 / 期限 2026-08-31 |
| A | SORT | SORT_count | `=COUNTA(SORT(E1:E6))` | 1 | 6 | SORTが無く、COUNTAがエラー値を1個数えている / 期限 2026-08-31 |
| A | UNIQUE | UNIQUE_join | `=TEXTJOIN(",",TRUE,UNIQUE(D1:D6))` | #NAME? | A,B,C | UNIQUEが無い / 期限 2026-08-31 |
| A | UNIQUE | UNIQUE_count | `=COUNTA(UNIQUE(D1:D6))` | 1 | 3 | 同上(COUNTAがエラー値を1個数えている) / 期限 2026-08-31 |
| A | UNIQUE | UNIQUE_nums | `=TEXTJOIN(",",TRUE,UNIQUE(E1:E6))` | #NAME? | 100,200,300,400,500,600 | 同上 / 期限 2026-08-31 |
| A | FILTER | FILTER_if_empty | `=TEXTJOIN(",",TRUE,FILTER(E1:E6,D1:D6="Z","なし"))` | #N/A | なし | 該当なしの時に返す第3引数に未対応 / 期限 2026-08-31 |
| A | SORT | ARRAY_sort_unique | `=TEXTJOIN(",",TRUE,SORT(UNIQUE(D1:D6)))` | #NAME? | A,B,C | SORTとUNIQUEの両方が要る / 期限 2026-08-31 |

## 入力の型が保たれるか（別枠）

セルに打ち込んだ文字をどう解釈するか。関数の検証とは別の話なので、混ぜずにここで見る。
期待値は想像ではなく、**実Excelの標準書式セルに同じ文字を打ち込んだ実測値**。

| ケース | 打ち込んだ値 | Exally | Excel(標準書式セル) | Excel型 | 判定 | 中身 |
|---|---|---|---|---|---|---|
| INPUT_typed_0007 | `0007` | 7 | 7 | n | 一致 |  |
| INPUT_typed_comma | `1,234` | 1,234 | 1234 | n | 不一致(既知) | 『1,234』と打っても数値にならない=以降の合計に入らない |
| INPUT_typed_datelike | `2026-07-31` | 2026-07-31 | 46234 | n | 不一致(既知) | 『2026-07-31』と打っても日付にならない=日付計算に使えない |
| INPUT_typed_code | `007-1234` | 007-1234 | 007-1234 | s | 一致 |  |

## ★独自層が生HFより悪くしていないか

独自層に関数を足すと、HyperFormula なら合っていた物を壊すことがある（TEXTの日付書式で実際に起きた）。
**独自層に関数を足す時は、必ずここが増えていないことを確認してから足す。** 台帳に載っていない劣化は赤。

劣化なし。

## 経路の固定（将来 生HF に落ちたら気付くための錠）

- 独自層(_jsComputeFormula)が答えたケース: **46件**
- 生HFと本番経路で答えが違うケース: **52件** … この差が消えたら「素通りに落ちた」ということ
- 独自層の入口: {"jsSetCount":1,"entryPoints":1}（1つだけであること）

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
| SUM | SUM_float | `=SUM(A4,0.2)` | 0.3 | 0.30000000000000004 | 未検証 | 0.3 | 一致 |  |
| SUM | SUM_mixed_args | `=SUM(E1:E3,100,A1)` | 1700 | 1700 | 未検証 | 1700 | 一致 |  |
| SUMPRODUCT | SUMPRODUCT_2range | `=SUMPRODUCT(C1:C6,E1:E6)` | 97100 | 97100 | 未検証 | 97100 | 一致 |  |
| SUMPRODUCT | SUMPRODUCT_1range | `=SUMPRODUCT(E1:E6)` | 2100 | 2100 | 未検証 | 2100 | 一致 |  |
| SUMPRODUCT | SUMPRODUCT_cond | `=SUMPRODUCT((D1:D6="A")*E1:E6)` | #VALUE! | 1000 | 未検証 | #VALUE! | 不一致(既知) | A |
| SUMPRODUCT | SUMPRODUCT_len | `=SUMPRODUCT(LEN(B1:B3))` | #VALUE! | 11 | 未検証 | #VALUE! | 不一致(既知) | A |
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
| AND | AND_all_true | `=AND(A1>0,A2>0)` | true | TRUE | 未検証 | TRUE | 一致 |  |
| AND | AND_one_false | `=AND(A1>0,A3>0)` | false | FALSE | 未検証 | FALSE | 一致 |  |
| AND | AND_cellbool | `=AND(H1,H1)` | true | TRUE | 未検証 | TRUE | 一致 |  |
| OR | OR_one_true | `=OR(A3>0,A1>0)` | true | TRUE | 未検証 | TRUE | 一致 |  |
| OR | OR_all_false | `=OR(A3>0,A8>0)` | false | FALSE | 未検証 | FALSE | 一致 |  |
| OR | OR_cellbool | `=OR(H2,H1)` | true | TRUE | 未検証 | TRUE | 一致 |  |
| NOT | NOT_true | `=NOT(TRUE)` | false | FALSE | 未検証 | #NAME? | 一致 |  |
| NOT | NOT_cell | `=NOT(H2)` | true | TRUE | 未検証 | TRUE | 一致 |  |
| NOT | NOT_expr | `=NOT(A1>A2)` | true | TRUE | 未検証 | TRUE | 一致 |  |
| AND | LOGIC_combo | `=IF(AND(OR(A1>0,NOT(TRUE)),A3<0),"y","n")` | y | y | 未検証 | #NAME? | 一致 |  |
| NOT | LOGIC_bool_to_num | `=(A1>0)*1+(A3>0)*1` | 1 | 1 | 未検証 | 1 | 一致 |  |
| VLOOKUP | VLOOKUP_exact | `=VLOOKUP(20,C1:E6,3,FALSE)` | 400 | 400 | 未検証 | #NAME? | 一致 |  |
| VLOOKUP | VLOOKUP_exact_miss | `=IFERROR(VLOOKUP(30,C1:E6,3,FALSE),"NA")` | NA | NA | 未検証 | NA | 一致 |  |
| VLOOKUP | VLOOKUP_approx | `=VLOOKUP(30,C1:E6,3,TRUE)` | 400 | 400 | 未検証 | #NAME? | 一致 |  |
| VLOOKUP | VLOOKUP_approx_low | `=IFERROR(VLOOKUP(0,C1:E6,3,TRUE),"NA")` | NA | NA | 未検証 | NA | 一致 |  |
| VLOOKUP | VLOOKUP_col2 | `=VLOOKUP(50,C1:E6,2,FALSE)` | B | B | 未検証 | #NAME? | 一致 |  |
| VLOOKUP | VLOOKUP_text_key | `=VLOOKUP("C",D1:E6,2,FALSE)` | 400 | 400 | 未検証 | #NAME? | 一致 |  |
| VLOOKUP | VLOOKUP_badcol | `=IFERROR(VLOOKUP(20,C1:E6,9,FALSE),"ERR")` | ERR | ERR | 未検証 | ERR | 一致 |  |
| XLOOKUP | XLOOKUP_basic | `=XLOOKUP(20,C1:C6,E1:E6)` | 400 | 400 | 未検証 | #NAME? | 一致 |  |
| XLOOKUP | XLOOKUP_notfound | `=XLOOKUP(999,C1:C6,E1:E6,"なし")` | なし | なし | 未検証 | #NAME? | 一致 |  |
| XLOOKUP | XLOOKUP_text | `=XLOOKUP("C",D1:D6,E1:E6,"なし")` | 400 | 400 | 未検証 | #NAME? | 一致 |  |
| XLOOKUP | XLOOKUP_first_dup | `=XLOOKUP("A",D1:D6,E1:E6,"なし")` | 100 | 100 | 未検証 | #NAME? | 一致 |  |
| XLOOKUP | XLOOKUP_no_default | `=IFERROR(XLOOKUP(999,C1:C6,E1:E6),"NA")` | NA | NA | 未検証 | NA | 一致 |  |
| INDEX | INDEX_row | `=INDEX(E1:E6,4)` | 400 | 400 | 未検証 | 400 | 一致 |  |
| INDEX | INDEX_2d | `=INDEX(C1:E6,4,3)` | 400 | 400 | 未検証 | 400 | 一致 |  |
| INDEX | INDEX_out_of_range | `=IFERROR(INDEX(E1:E6,99),"ERR")` | ERR | ERR | 未検証 | ERR | 一致 |  |
| INDEX | INDEX_text_range | `=INDEX(B1:B8,8)` | 山田 太郎 | 山田 太郎 | 未検証 | 山田 太郎 | 一致 |  |
| MATCH | MATCH_exact | `=MATCH(20,C1:C6,0)` | 4 | 4 | 未検証 | 4 | 一致 |  |
| MATCH | MATCH_exact_miss | `=IFERROR(MATCH(30,C1:C6,0),"NA")` | NA | NA | 未検証 | NA | 一致 |  |
| MATCH | MATCH_approx_asc | `=MATCH(30,C1:C6,1)` | 4 | 4 | 未検証 | 4 | 一致 |  |
| MATCH | MATCH_text | `=MATCH("C",D1:D6,0)` | 4 | 4 | 未検証 | 4 | 一致 |  |
| MATCH | MATCH_wildcard | `=MATCH("りん*",B1:B8,0)` | #N/A | 1 | 未検証 | #N/A | 不一致(既知) | A |
| INDEX | INDEX_MATCH_combo | `=INDEX(E1:E6,MATCH(20,C1:C6,0))` | 400 | 400 | 未検証 | 400 | 一致 |  |
| INDEX | INDEX_MATCH_text | `=INDEX(E1:E6,MATCH("C",D1:D6,0))` | 400 | 400 | 未検証 | 400 | 一致 |  |
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
| VALUE | VALUE_num_text | `=VALUE(B6)` | 7 | 7 | 未検証 | #NAME? | 一致 |  |
| VALUE | VALUE_comma | `=IFERROR(VALUE(B7),"NA")` | 1234 | 1234 | 未検証 | NA | 一致 |  |
| VALUE | VALUE_bad | `=IFERROR(VALUE(B1),"NA")` | NA | NA | 未検証 | NA | 一致 |  |
| VALUE | VALUE_zero_text | `=VALUE(G3)` | 0 | 0 | 未検証 | #NAME? | 一致 |  |
| TEXT | TEXT_pct_zero | `=TEXT(0.5,"0%")` | 50% | 50% | 未検証 | 1% | 一致 |  |
| TEXT | TEXT_dec_round | `=TEXT(2.675,"0.00")` | 2.68 | 2.68 | 未検証 | 2.67 | 一致 |  |
| TEXT | TEXT_neg_number | `=TEXT(-1234.5,"#,##0")` | -1,235 | -1,235 | 未検証 | -1235,##0 | 一致 |  |
| TEXT | TEXT_date_yy | `=TEXT(F5,"yy/m/d")` | 26/7/31 | 26/7/31 | 未検証 | 26/7/31 | 一致 |  |
| TEXT | TEXT_nested | `=LEN(TEXT(1234.5,"#,##0"))` | 8 | 5 | 未検証 | 8 | 不一致(既知) | A |
| VALUE | VALUE_empty | `=IFERROR(VALUE(""),"NA")` | NA | NA | 未検証 | NA | 一致 |  |
| VALUE | VALUE_nested_calc | `=IFERROR(VALUE(B7)+1,"NA")` | 1235 | 1235 | 未検証 | NA | 一致 |  |
| TEXTJOIN | TEXTJOIN_skip | `=TEXTJOIN("\|",TRUE,B1:B3)` | りんご\|みかん\|apple | りんご\|みかん\|apple | 未検証 | #NAME? | 一致 |  |
| TEXTJOIN | TEXTJOIN_delim_comma | `=TEXTJOIN(",",TRUE,"a","b")` | a,b | a,b | 未検証 | #NAME? | 一致 |  |
| TEXTJOIN | TEXTJOIN_nested_fn | `=TEXTJOIN("\|",TRUE,LEFT(B1,1),LEFT(B2,1))` | り\|み | り\|み | 未検証 | #NAME? | 一致 |  |
| TEXTJOIN | TEXTJOIN_mixed | `=TEXTJOIN("-",TRUE,B5,A1,C1:C2)` | A-1-1000-1-5 | A-1-1000-1-5 | 未検証 | #NAME? | 一致 |  |
| TEXTJOIN | TEXTJOIN_keep | `=TEXTJOIN("\|",FALSE,G1:G3)` | \|0\|0 | \|0\|0 | 未検証 | #NAME? | 一致 |  |
| TEXTJOIN | TEXTJOIN_nums | `=TEXTJOIN(",",TRUE,E1:E6)` | 100,200,300,400,500,600 | 100,200,300,400,500,600 | 未検証 | #NAME? | 一致 |  |
| TEXTJOIN | TEXTJOIN_blank_skip | `=TEXTJOIN("\|",TRUE,G1:G3)` | 0\|0 | 0\|0 | 未検証 | #NAME? | 一致 |  |
| SORT | SORT_asc_join | `=TEXTJOIN(",",TRUE,SORT(E1:E6))` | #NAME? | 100,200,300,400,500,600 | 未検証 | #NAME? | 不一致(既知) | A |
| SORT | SORT_desc_join | `=TEXTJOIN(",",TRUE,SORT(E1:E6,1,-1))` | #NAME? | 600,500,400,300,200,100 | 未検証 | #NAME? | 不一致(既知) | A |
| SORT | SORT_first | `=INDEX(SORT(E1:E6,1,-1),1)` | #NAME? | 600 | 未検証 | #NAME? | 不一致(既知) | A |
| SORT | SORT_text | `=TEXTJOIN(",",TRUE,SORT(D1:D6))` | #NAME? | A,A,A,B,B,C | 未検証 | #NAME? | 不一致(既知) | A |
| SORT | SORT_count | `=COUNTA(SORT(E1:E6))` | 1 | 6 | 未検証 | 1 | 不一致(既知) | A |
| UNIQUE | UNIQUE_join | `=TEXTJOIN(",",TRUE,UNIQUE(D1:D6))` | #NAME? | A,B,C | 未検証 | #NAME? | 不一致(既知) | A |
| UNIQUE | UNIQUE_count | `=COUNTA(UNIQUE(D1:D6))` | 1 | 3 | 未検証 | 1 | 不一致(既知) | A |
| UNIQUE | UNIQUE_nums | `=TEXTJOIN(",",TRUE,UNIQUE(E1:E6))` | #NAME? | 100,200,300,400,500,600 | 未検証 | #NAME? | 不一致(既知) | A |
| UNIQUE | UNIQUE_all_same | `=COUNTA(UNIQUE(D1:D1))` | 1 | 1 | 未検証 | 1 | 一致 |  |
| FILTER | FILTER_sum | `=SUM(FILTER(E1:E6,D1:D6="A"))` | 1000 | 1000 | 未検証 | 1000 | 一致 |  |
| FILTER | FILTER_join | `=TEXTJOIN(",",TRUE,FILTER(E1:E6,D1:D6="A"))` | 100,300,600 | 100,300,600 | 未検証 | #NAME? | 一致 |  |
| FILTER | FILTER_num_cond | `=TEXTJOIN(",",TRUE,FILTER(E1:E6,C1:C6>=10))` | 300,400,500,600 | 300,400,500,600 | 未検証 | #NAME? | 一致 |  |
| FILTER | FILTER_empty | `=IFERROR(TEXTJOIN(",",TRUE,FILTER(E1:E6,D1:D6="Z")),"NA")` | NA | NA | 未検証 | NA | 一致 |  |
| FILTER | FILTER_if_empty | `=TEXTJOIN(",",TRUE,FILTER(E1:E6,D1:D6="Z","なし"))` | #N/A | なし | 未検証 | #NAME? | 不一致(既知) | A |
| SORT | ARRAY_sort_unique | `=TEXTJOIN(",",TRUE,SORT(UNIQUE(D1:D6)))` | #NAME? | A,B,C | 未検証 | #NAME? | 不一致(既知) | A |
