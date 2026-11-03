# toru-isomitted.ps1 — ★ISOMITTED の「本当の 使い方」を 実Excel に 聞く★（2026-09-08）
#
#  ★★なぜ 要るか（指示役 2026-09-08）★★
#    `golden-346-2026-09-08.tsv` の ISOMITTED は ★16本 とも LAMBDA の 外★で、★全部 False★。
#    ⇒★LAMBDA の 中で 使った 物が ★1本も 在りません★★
#    ⇒ このまま「引数が 在れば FALSE」に 直すと
#      ★16本 とも 合う＝金の紙は 全部 緑★／★でも 本当の 仕事は 1つも できていない★
#    ⇒★★これは 半分 合う 計算★＝ISREF の「A:A が 黙って 逆」と 同じ 型★★
#
#  ★★だから 先に 実Excel に 聞く★★
#    ・『空の マス』と『省かれた 引数』が ★別物★だと ★数で★ 出す
#
#  ★司さんの 実物には 触りません★＝★新しい ブック★を 開いて 使い、保存せずに 閉じる
#
#  使い方: powershell -File docs/measured/kansuu46/toru-isomitted.ps1

$ErrorActionPreference = 'Stop'

# ══ ★★2つ目の 窓（2026-09-08 に 足した）★★ ══════════════════
#  ★物差しの 欠陥★ .Value2 は ★0 で ない 値に 0 を 返す★
#    =0.1+0.2-0.3    … .Value2 ★0★ ／ =(式)=0 ★False★ ／ (式)*1e17 5.55
#    =11.1+22.2-33.3 … .Value2 0   ／ =(式)=0 ★True★  ／ (式)*1e17 0
#    ⇒★.Value2 では この 2つが どちらも 0 に 見える★
#    正体 …★最後の 演算が ＋か− の 時だけ 実Excel が ★見せる 時に★ 0 に する★
#  ★もう1つ★ =DEC2BIN(0.5) は ★文字列の "0"★＝数の 0 では ない
#    ⇒ ="0"=0 は FALSE ⇒★見せかけの 0 と 同じ 顔★⇒★型を 見ないと 分けられない★
#  ⇒★見張り tests/monosashi-mado.test.mjs が これを 入れて いない 道具を 赤に する★
function 窓２_型($v) {
  if ($null -eq $v) { return 'Empty' }
  if ($v -is [string]) { return 'String' }
  if ($v -is [bool]) { return 'Boolean' }
  if ($v -is [double] -or $v -is [int] -or $v -is [long]) { return 'Number' }
  return 'Other'
}
function 窓２_本当にゼロか($sh, [string]$式) {
  # ★『0』が 出た 時だけ 呼ぶ★ … =(式)=0 の 真偽を 返す
  $中 = $式 -replace '^=\s*', ''
  try {
    $sh.Range('BZ1').Clear() | Out-Null
    $sh.Range('BZ1').Formula = ('=(' + $中 + ')=0')
    $z = $sh.Range('BZ1').Value2
    $sh.Range('BZ1').Clear() | Out-Null
    if ($z -is [bool]) { return $(if ($z) { 'TRUE' } else { 'FALSE' }) }
    return '★判じられない★'
  } catch { return '★判じられない★' }
}

$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-isomitted-2026-09-08.tsv'

# ★誤りの 番号 → 名前★（★2026-09-08 に 新しい 6種を 足した★
#   ⇒ 表に 無い 番号は ★「数」として 紙に 載ってしまう★＝偽の 赤に なる）
$誤りの番号 = @{
  -2146826288 = '#NULL!'; -2146826281 = '#DIV/0!'; -2146826273 = '#VALUE!';
  -2146826265 = '#REF!';  -2146826259 = '#NAME?';  -2146826252 = '#NUM!';
  -2146826246 = '#N/A';
  -2146826243 = '#SPILL!';   -2146826242 = '#CONNECT!'; -2146826241 = '#BLOCKED!';
  -2146826240 = '#UNKNOWN!'; -2146826239 = '#FIELD!';   -2146826238 = '#CALC!'; -2146826237 = '#BUSY!'
}

# ★聞く 式★（★人が 決めるのは ここだけ★）
#   ★『空の マス』と『省かれた 引数』を 分けて 見る為の 組★
$候補 = @(
  @{ 札 = 'LAMBDA の 中・渡した';        式 = '=LAMBDA(x,y,ISOMITTED(y))(1,2)' },
  @{ 札 = 'LAMBDA の 中・★省いた★';      式 = '=LAMBDA(x,y,ISOMITTED(y))(1)' },
  @{ 札 = 'LAMBDA の 中・1つ 渡した';    式 = '=LAMBDA(x,ISOMITTED(x))(1)' },
  @{ 札 = 'LAMBDA の 中・★空のマス★';    式 = '=LAMBDA(x,ISOMITTED(x))(C1)' },
  @{ 札 = 'LAMBDA の 中・値の在るマス';  式 = '=LAMBDA(x,ISOMITTED(x))(A1)' },
  @{ 札 = 'LAMBDA の 中・省いて IF';     式 = '=LAMBDA(x,y,IF(ISOMITTED(y),"ナシ",y))(1)' },
  @{ 札 = 'LAMBDA の 中・渡して IF';     式 = '=LAMBDA(x,y,IF(ISOMITTED(y),"ナシ",y))(1,9)' },
  @{ 札 = 'LET の 中';                   式 = '=LET(x,1,ISOMITTED(x))' },
  @{ 札 = '裸（引数なし）';              式 = '=ISOMITTED()' },
  @{ 札 = 'LAMBDA の 外・空のマス';      式 = '=ISOMITTED(C1)' },
  @{ 札 = 'LAMBDA の 外・値のマス';      式 = '=ISOMITTED(A1)' },
  # ★★2回目（1回目で 止めない）★★
  #   1回目 … `=LAMBDA(x,y,ISOMITTED(y))(1)` は ★#VALUE!★＝★TRUE では なかった★
  #   ⇒★引数の 数が 足りない 呼び方は 実Excel が ★呼び出しの 所で★ 断っている★
  #   ⇒ 実Excel の「省く」は ★カンマを 置く★形かもしれない ⇒ そこを 聞く
  @{ 札 = '★カンマで 省く★ (1,)';         式 = '=LAMBDA(x,y,ISOMITTED(y))(1,)' },
  @{ 札 = '★カンマで 省く★ (1,) と IF';    式 = '=LAMBDA(x,y,IF(ISOMITTED(y),"ナシ",y))(1,)' },
  @{ 札 = '★カンマで 省く★ (,2) 前を';     式 = '=LAMBDA(x,y,ISOMITTED(x))(,2)' },
  @{ 札 = '★カンマで 省く★ (,2) 後を';     式 = '=LAMBDA(x,y,ISOMITTED(y))(,2)' },
  @{ 札 = '★3つ中 真ん中を 省く★';         式 = '=LAMBDA(x,y,z,ISOMITTED(y))(1,,3)' }
)

Write-Host '★実Excel を 開きます（新しい ブック・保存しません）★'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Worksheets.Item(1)

  # ★金の紙を 取った 時と 同じ 土台★（C 列は ★空のまま★）
  $sh.Range('A1').Value2 = 1
  $sh.Range('A2').Value2 = 2
  $sh.Range('A3').Value2 = 3
  $sh.Range('A4').Value2 = 4
  $sh.Range('A5').Value2 = 5
  $sh.Range('B1').Value2 = 2
  $sh.Range('B2').Value2 = 4
  $sh.Range('B3').Value2 = 6
  $sh.Range('B4').Value2 = 8
  $sh.Range('B5').Value2 = 10
  $sh.Range('D1').Formula = '=DATE(2024,1,1)'
  $sh.Range('D2').Formula = '=DATE(2026,1,1)'

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★ISOMITTED の「本当の 使い方」を 実Excel に 聞いた★（2026-09-08）')
  $行.Add('# ★金の紙 16本は 全部 LAMBDA の 外＝本当の 仕事を 1本も 測っていなかった★（指示役）')
  # ★★1列目は ★関数名★（他の 20枚と 同じ 形）★★
  #   ⇒ 2026-09-08 に 直した … 前は 1列目が ★札★で、
  #     ★grep で 数える 道具から この 紙が 16行 全部 漏れた★
  #   ⇒★自分で 作った 紙が 自分の 掃除から 漏れる★＝作った その場で 通す
  #   ⇒ 札は ★消さず 5列目へ★（指示役「中身を 残して 形を 揃える」）
  $行.Add('# 関数' + "`t" + '式' + "`t" + '実Excel の 答え' + "`t" + '型' + "`t" + '★札★')

  foreach ($c in $候補) {
    $値 = $null
    $型 = ''
    try {
      $sh.Range('H1').Formula = $c.式
      $v = $sh.Range('H1').Value2
      if ($null -eq $v) {
        $値 = ''; $型 = 'Empty'
      } elseif ($v -is [int] -or $v -is [long]) {
        if ($誤りの番号.ContainsKey([int]$v)) { $値 = $誤りの番号[[int]$v]; $型 = 'Error' }
        else { $値 = [string]$v; $型 = 'Int32' }
      } elseif ($v -is [double]) {
        $値 = [string]$v; $型 = 'Double'
      } elseif ($v -is [bool]) {
        $値 = if ($v) { 'True' } else { 'False' }
        $型 = 'Boolean'
      } else {
        $値 = [string]$v; $型 = 'String'
      }
    } catch {
      $値 = '★実Excel が 受け付けない★'; $型 = 'Rejected'
    }
    $行.Add('ISOMITTED' + "`t" + $c.式 + "`t" + $値 + "`t" + $型 + "`t" + $c.札)
    Write-Host ('  ' + $c.札.PadRight(26) + $c.式.PadRight(40) + ' → ' + $値 + '  (' + $型 + ')')
  }

  [System.IO.File]::WriteAllLines($出, $行, (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ''
  Write-Host ('★書いた … ' + $出 + '★')

  $bk.Close($false)
} finally {
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
