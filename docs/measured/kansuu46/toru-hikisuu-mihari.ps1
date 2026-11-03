# toru-hikisuu-mihari.ps1 — ★引数の 見張りの 境目を 実Excel に 聞く★（2026-09-08）
#
#  ★★なぜ★★
#    書式の 付いた 数を 直した 後、まだ 16本 食い違う。
#    その うち ★危ない側（実Excel は 断るのに うちは 答えてしまう）が 9本★。
#      =TOCOL(A1:B5,DATE(...))   実Excel #VALUE! ／ うち 1
#      =FINDB("b","abc",TIME(...)) 実Excel #VALUE! ／ うち 2
#      =REPLACEB("abcde",TIME(...),1,"x") 実Excel #VALUE! ／ うち xbcde
#    ⇒★引数の 見張りが 無い／緩い★
#    ⇒★★どこから 断るかは ★私が 決めません＝実Excel に 聞きます★★★
#
#  ★★物差し★★ 表を 返す 式は ★=INDEX(式,1,1)★（★溢れさせない★）
#    ⇒ そのまま 打つと `.Value2` が 溢れの 産物を 返す（2026-09-08 に DROP で 踏んだ）
#
#  ★司さんの 実物には 触りません★＝新しい ブック・保存せず
#
#  使い方: powershell -File docs/measured/kansuu46/toru-hikisuu-mihari.ps1

$ErrorActionPreference = 'Stop'

# ══ ★★2つ目の 窓（2026-09-09 に 足した）★★ ══════════════════
#  ★なぜ 後から 足したか★
#    この 道具は #52 で 作りました。その 時 まだ ★2つ目の 窓の 決まりが 無かった★。
#    #54（物差しの 直し）が main に 入って ★見張りが 出来た★ので、
#    ★#52 と #54 は 別々では 緑・一緒に すると 赤★に なりました。
#    ⇒★どちらの PR でも 見えない 形★＝入ってから 出た
#
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
$出 = Join-Path $ここ 'golden-hikisuu-mihari-2026-09-08.tsv'

$誤りの番号 = @{
  -2146826288 = '#NULL!'; -2146826281 = '#DIV/0!'; -2146826273 = '#VALUE!';
  -2146826265 = '#REF!';  -2146826259 = '#NAME?';  -2146826252 = '#NUM!';
  -2146826246 = '#N/A';
  -2146826243 = '#SPILL!';   -2146826242 = '#CONNECT!'; -2146826241 = '#BLOCKED!';
  -2146826240 = '#UNKNOWN!'; -2146826239 = '#FIELD!';   -2146826238 = '#CALC!'
}

# ★渡す 数（★境目の 前後を 必ず 両方★）★
$数ら = @('-2', '-1', '-0.5', '0', '0.5', '1', '1.5', '2', '2.5', '3', '3.5', '4', '5', '6', '45294')

# ★見張りたい 引数（★名簿から 選んだ★＝16本の 食い違いに 出てきた 式だけ）★
#   `{N}` が 上の 数に 置き換わる
$候補 = @(
  @{ 名 = 'TOCOL の 2番目（無視する 印）';     式 = '=TOCOL(A1:B5,{N})';            表 = $true },
  @{ 名 = 'TOROW の 2番目（無視する 印）';     式 = '=TOROW(A1:B5,{N})';            表 = $true },
  @{ 名 = 'FINDB の 3番目（何文字目から）';    式 = '=FINDB("b","abc",{N})';        表 = $false },
  @{ 名 = 'SEARCHB の 3番目（何文字目から）';  式 = '=SEARCHB("b","abc",{N})';      表 = $false },
  @{ 名 = 'REPLACEB の 2番目（何文字目から）'; 式 = '=REPLACEB("abcde",{N},1,"x")'; 表 = $false },
  @{ 名 = 'REPLACEB の 3番目（何文字ぶん）';   式 = '=REPLACEB("abcde",1,{N},"x")'; 表 = $false },
  @{ 名 = 'CHOOSEROWS の 2番目（何行目）';     式 = '=CHOOSEROWS(A1:A5,{N})';       表 = $true },
  @{ 名 = 'CHOOSECOLS の 2番目（何列目）';     式 = '=CHOOSECOLS(A1:B5,{N})';       表 = $true },
  @{ 名 = 'TAKE の 2番目（何行）';             式 = '=TAKE(A1:A5,{N})';             表 = $true },
  @{ 名 = 'DROP の 2番目（何行 落とす）';      式 = '=DROP(A1:A5,{N})';             表 = $true },
  @{ 名 = 'WRAPROWS の 2番目（何個で 折る）';  式 = '=WRAPROWS(A1:A5,{N})';         表 = $true },
  @{ 名 = 'WRAPCOLS の 2番目（何個で 折る）';  式 = '=WRAPCOLS(A1:A5,{N})';         表 = $true }
)

Write-Host ('★数 ' + $数ら.Count + '通り × 引数 ' + $候補.Count + '個 ＝ ' + ($数ら.Count * $候補.Count) + '本 打ちます★')

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Worksheets.Item(1)
  for ($r = 1; $r -le 5; $r++) {
    $sh.Cells.Item($r, 1).Value2 = $r
    $sh.Cells.Item($r, 2).Value2 = $r * 2
  }

  function 押して字に($式) {
    $sh.Range('T1:Z40').Clear() | Out-Null
    try {
      $sh.Range('T1').Formula = $式
      $v = $sh.Range('T1').Value2
      if ($null -eq $v) { return @{ 値 = '(空)'; 型 = 'Empty' } }
      if (($v -is [int] -or $v -is [long]) -and $誤りの番号.ContainsKey([int]$v)) {
        return @{ 値 = $誤りの番号[[int]$v]; 型 = 'Error' }
      }
      if ($v -is [bool]) { return @{ 値 = $(if ($v) { 'True' } else { 'False' }); 型 = 'Boolean' } }
      # ★★『0』が 出た 時だけ 2つ目の 窓を 開ける★★（.Value2 の 0 を そのまま 信じない）
      #   ★型も 一緒に 出す★＝★見せかけの 0★と★字の "0"★は .Value2 では 同じ 顔
      if ($v -is [string] -and $v -eq '0') {
        return @{ 値 = '0'; 型 = '字の"0"'; 窓2 = (窓２_本当にゼロか $sh $式); 窓2型 = (窓２_型 $v) }
      }
      if ($v -is [double]) {
        if ($v -eq 0) {
          return @{ 値 = '0'; 型 = 'Double'; 窓2 = (窓２_本当にゼロか $sh $式); 窓2型 = (窓２_型 $v) }
        }
        return @{ 値 = $v.ToString('R'); 型 = 'Double' }
      }
      if ($v -is [int] -or $v -is [long]) { return @{ 値 = [string]$v; 型 = 'Int32' } }
      return @{ 値 = [string]$v; 型 = 'String' }
    } catch { return @{ 値 = '★受け付けない★'; 型 = 'Rejected' } }
  }

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★引数の 見張りの 境目を 実Excel に 聞いた★（2026-09-08）')
  $行.Add('# ★どこから 断るかは ★私が 決めない＝実Excel の 答えの 通り★')
  $行.Add('# ★境目の 前後を 必ず 両方★ … -2 -1 -0.5 0 0.5 1 1.5 2 2.5 3 3.5 4 5 6 45294')
  $行.Add('# ★物差し★ 表を 返す 式は =INDEX(式,1,1)（★溢れさせない★）')
  $行.Add('# ★土台★ A1:A5=1..5 / B1:B5=2,4,6,8,10')
  $行.Add('# 関数' + "`t" + '式' + "`t" + '実Excel の 答え' + "`t" + '型' + "`t" + '渡した数' + "`t" + 'どの引数')

  $本数 = 0
  foreach ($k in $候補) {
    $ならび = @()
    foreach ($n in $数ら) {
      $式 = $k.式.Replace('{N}', $n)
      if ($k.表) {
        $中 = $式 -replace '^=\s*', ''
        $r = 押して字に ('=INDEX(' + $中 + ',1,1)')
      } else {
        $r = 押して字に $式
      }
      $名 = (($式 -replace '^=\s*', '') -replace '[^A-Za-z0-9._].*$', '').ToUpper()
      $行.Add($名 + "`t" + $式 + "`t" + $r.値 + "`t" + $r.型 + "`t" + $n + "`t" + $k.名)
      $ならび += ($n + '=' + $r.値)
      $本数++
    }
    Write-Host ('  ' + $k.名)
    Write-Host ('    ' + ($ならび -join '  '))
  }

  # ★★書き戻しは 必ず LF（2026-09-09 に 直した）★★
  #   `WriteAllLines` は Windows では ★CRLF★ で 書きます。
  #   ⇒ この 紙は CRLF に なって いました（★repo の 決まりは LF★・`tests/eol.test.mjs`）
  #   ⇒ 行を 自分で つないで `WriteAllText` で 書く
  [System.IO.File]::WriteAllText($出, ($行 -join "`n") + "`n", (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ''
  Write-Host ('★★書けた … ' + $本数 + '本★★')
  Write-Host ('★書いた … ' + $出 + '★')

  $bk.Close($false)
} finally {
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
