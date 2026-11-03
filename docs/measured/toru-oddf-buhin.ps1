# toru-oddf-buhin.ps1 — ★ODDFPRICE の ★部品★を 実Excel に 直に 聞く★（2026-09-16）
#
#  ★★なぜ 部品を 聞くか★★
#    2026-09-16 に ODDFPRICE の ★答えを 当てて 6通り 試し★、★全部 元より 悪く★ なった。
#    ⇒★答えを 当てるのは 数字合わせ★＝やめた（棚63 に 書いた）
#    ⇒★★答えでは なく 部品を 聞く★★
#       ODDFPRICE は ★端数期間の 長さ★を 数えて 割り引く 式。
#       その ★日数の 数え方★は `COUPDAYBS` `COUPDAYS` `COUPDAYSNC` `COUPNUM` で
#       ★実Excel に 直に 聞けます★（どれも 実Excel に 在る 関数）。
#    ⇒★basis ごとに どう 数えるかが 分かれば ODDFPRICE の 中身が 決まります★
#
#  ★★決まり★★
#    ・2つの 道具で 数えてから／新しい 空の ブックだけ／Visible=$false
#    ・finally で Quit／消えるまで 待って 秒数を 出す／★BOM 付き★
#    ・★2つ目の 窓（=(式)=0）と 型★も 取る（`.Value2` は 0 で ない 値に 0 を 返す）
#    ・★司さんの 実物は 開きません★
#
#  使い方: pwsh -NoProfile -File docs/measured/toru-oddf-buhin.ps1

$ErrorActionPreference = 'Stop'
$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-oddf-buhin-2026-09-16.tsv'

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★前に 居た Excel … Get-Process ' + $数1 + '個 ／ Win32_Process ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Error '★★Excel が 動いて います＝走らせません★★'; exit 3 }

# ★ODDFPRICE の 紙と ★同じ 日付★★（golden-kane-2026-09-07.tsv の 6組）
$組 = @(
  @{ 名 = 'A'; 決 = 'DATE(2008,11,11)'; 満 = 'DATE(2021,3,1)';  発 = 'DATE(2008,10,15)'; 初 = 'DATE(2009,3,1)';  f = 2 },
  @{ 名 = 'B'; 決 = 'DATE(2008,1,15)';  満 = 'DATE(2012,6,30)'; 発 = 'DATE(2007,11,20)'; 初 = 'DATE(2008,6,30)'; f = 2 },
  @{ 名 = 'C'; 決 = 'DATE(2009,3,10)';  満 = 'DATE(2012,8,31)'; 発 = 'DATE(2008,12,5)';  初 = 'DATE(2009,8,31)'; f = 4 },
  @{ 名 = 'D'; 決 = 'DATE(2009,6,1)';   満 = 'DATE(2012,1,1)';  発 = 'DATE(2009,1,1)';   初 = 'DATE(2010,1,1)'; f = 2 },
  @{ 名 = 'E'; 決 = 'DATE(2009,6,1)';   満 = 'DATE(2012,1,1)';  発 = 'DATE(2008,7,1)';   初 = 'DATE(2010,1,1)'; f = 2 },
  @{ 名 = 'F'; 決 = 'DATE(2009,6,1)';   満 = 'DATE(2012,1,1)';  発 = 'DATE(2009,4,1)';   初 = 'DATE(2010,1,1)'; f = 2 }
)

$xl = New-Object -ComObject Excel.Application
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Sheets.Item(1)

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★ODDFPRICE の 部品を 実Excel に 直に 聞いた★（2026-09-16）')
  $行.Add('#')
  $行.Add('# ★なぜ★ 答えを 当てて 6通り 試して 全部 悪く なった（棚63）')
  $行.Add('#   ⇒★答えでは なく 部品（日数の 数え方）を 聞く★')
  $行.Add('#')
  $行.Add('# ★聞く 部品★')
  $行.Add('#   COUPDAYBS  … 前の 利払日から 決済までの 日数')
  $行.Add('#   COUPDAYS   … 決済が 入って いる 利払期間の 日数')
  $行.Add('#   COUPDAYSNC … 決済から 次の 利払日までの 日数')
  $行.Add('#   COUPNUM    … 決済から 満期までの 利払回数')
  $行.Add('#   YEARFRAC   … 2つの 日の 間の 年数（basis ごと）')
  $行.Add('#')
  $行.Add('# ★2つ目の 窓★ `=(式)=0` … `.Value2` は 0 で ない 値に 0 を 返す')
  $行.Add('#')
  $行.Add('# ★どの Excel か★ … 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('#')
  $行.Add('# 組' + "`t" + 'basis' + "`t" + '部品' + "`t" + '式' + "`t" + '答え' + "`t" + '=(式)=0' + "`t" + '型')

  $r = 1
  foreach ($g in $組) {
    foreach ($b in 0..4) {
      $部品たち = @(
        @{ 名 = 'COUPDAYBS';  式 = '=COUPDAYBS('  + $g.決 + ',' + $g.満 + ',' + $g.f + ',' + $b + ')' },
        @{ 名 = 'COUPDAYS';   式 = '=COUPDAYS('   + $g.決 + ',' + $g.満 + ',' + $g.f + ',' + $b + ')' },
        @{ 名 = 'COUPDAYSNC'; 式 = '=COUPDAYSNC(' + $g.決 + ',' + $g.満 + ',' + $g.f + ',' + $b + ')' },
        @{ 名 = 'COUPNUM';    式 = '=COUPNUM('    + $g.決 + ',' + $g.満 + ',' + $g.f + ',' + $b + ')' },
        # ★端数期間そのもの★ … 発行 → 初回利払
        @{ 名 = 'YF_発初';    式 = '=YEARFRAC('   + $g.発 + ',' + $g.初 + ',' + $b + ')' },
        # ★発行 → 決済★（経過利息の 分）
        @{ 名 = 'YF_発決';    式 = '=YEARFRAC('   + $g.発 + ',' + $g.決 + ',' + $b + ')' },
        # ★決済 → 初回利払★
        @{ 名 = 'YF_決初';    式 = '=YEARFRAC('   + $g.決 + ',' + $g.初 + ',' + $b + ')' }
      )
      foreach ($p in $部品たち) {
        $c = $sh.Range('A' + $r)
        $c.Formula = $p.式
        $v = $c.Value2
        $答 = if ($null -eq $v) { '(空)' } elseif ($v -is [double]) { $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
        $型 = if ($null -eq $v) { '(空)' } elseif ($v -is [double]) { 'Double' } elseif ($v -is [string]) { 'String' } else { 'Other' }
        $w = $sh.Range('B' + $r)
        $w.Formula = '=(' + $p.式.Substring(1) + ')=0'
        $ゼロか = [string]$w.Value2
        $行.Add($g.名 + "`t" + $b + "`t" + $p.名 + "`t" + $p.式 + "`t" + $答 + "`t" + $ゼロか + "`t" + $型)
        $r++
      }
    }
    Write-Host ('  組 ' + $g.名 + ' … 済')
  }

  [System.IO.File]::WriteAllText($出, ($行 -join "`n") + "`n", (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ''
  Write-Host ('★書いた … ' + $出 + '（' + ($行.Count - 17) + '行）★')
  $bk.Close($false)
} finally {
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  $t = [Diagnostics.Stopwatch]::StartNew()
  while ((@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -gt 0) -and ($t.Elapsed.TotalSeconds -lt 180)) {
    Start-Sleep -Milliseconds 500
  }
  $t.Stop()
  Write-Host ('★Excel が 消えるまで ' + [math]::Round($t.Elapsed.TotalSeconds, 1) + '秒 ／ 残り ' +
    @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count + '個★')
}
