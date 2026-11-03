# toru-yearfrac-basis1.ps1 — ★YEARFRAC の basis 1 は 何で 割って いるか★（2026-09-16）
#
#  ★★なぜ 要るか★★
#    ODDFPRICE の 部品を 実Excel に 聞いたら
#    COUPDAYBS/COUPDAYS/COUPDAYSNC/COUPNUM は ★120本 全部 合った★のに
#    ★YEARFRAC の basis 1 だけ 30本中 7本 違った★。
#    うちは ★いつでも「年の 平均日数」で 割って います★（365.5 など）。
#    実Excel は ★365 と 366 を 使い分けて います★。
#    ⇒★★分かれ道を 切り分ける 組を 実Excel に 打たせて 測ります★★
#      （★答えを 当てない★＝棚63 で 当てて 6通り 全部 悪く なった）
#
#  ★YEARFRAC は お客さんが 直に 打つ 関数★＝これ 自体が ★客に 出る 欠陥★。
#
#  ★切り分ける 分かれ道★
#    ①同じ 年（うるう年 / 平年）
#    ②年を またぐが 1年 以内 … 始めの 年が うるう年（2/29 の 前 / 後）
#    ③年を またぐが 1年 以内 … 終わりの 年が うるう年（2/29 の 前 / 後）
#    ④ちょうど 2/29 で 始まる / 終わる
#    ⑤ちょうど 1年（同じ 月日）と その 1日 先
#    ⑥1年 超え（★何年を 平均するのか★）
#
#  ★★決まり★★
#    ・2つの 道具で 数えてから／新しい 空の ブックだけ／Visible=$false
#    ・finally で Quit／消えるまで 待って 秒数を 出す／★BOM 付き★
#    ・★2つ目の 窓（=(式)=0）と 型★も 取る
#    ・★司さんの 実物は 開きません★
#
#  使い方: pwsh -NoProfile -File docs/measured/toru-yearfrac-basis1.ps1

$ErrorActionPreference = 'Stop'
$ここ = Split-Path -Parent $MyInvocation.MyCommand.Path
$出 = Join-Path $ここ 'golden-yearfrac-basis1-2026-09-16.tsv'

$数1 = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
$数2 = @(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue).Count
Write-Host ('★前に 居た Excel … Get-Process ' + $数1 + '個 ／ Win32_Process ' + $数2 + '個★')
if ($数1 -ne 0 -or $数2 -ne 0) { Write-Error '★★Excel が 動いて います＝走らせません★★'; exit 3 }

# ★組 ＝ @(始, 終, なぜ この 組か)★
$組 = @(
  @('2009-01-15', '2009-06-30', '①同じ年・平年'),
  @('2008-01-15', '2008-06-30', '①同じ年・うるう年・2/29 を またぐ'),
  @('2008-03-01', '2008-12-31', '①同じ年・うるう年・2/29 を またぎません'),
  @('2008-01-01', '2008-02-01', '①同じ年・うるう年・2/29 より 前だけ'),
  @('2009-01-01', '2009-12-31', '①同じ年・平年・ほぼ 1年'),
  @('2008-01-01', '2008-12-31', '①同じ年・うるう年・ほぼ 1年'),

  @('2008-01-15', '2009-01-10', '②またぐ・始が うるう年・始は 2/29 より 前'),
  @('2008-10-15', '2009-03-01', '②またぐ・始が うるう年・始は 2/29 より 後（測済 365）'),
  @('2008-02-29', '2009-01-15', '④始が ちょうど 2/29'),
  @('2008-02-28', '2009-01-15', '④始が 2/28（上の 1日 前）'),

  @('2007-11-20', '2008-06-30', '③またぐ・終が うるう年・終は 2/29 より 後（測済 366）'),
  @('2007-11-20', '2008-01-15', '③またぐ・終が うるう年・終は 2/29 より 前（測済 365）'),
  @('2007-11-20', '2008-02-29', '④終が ちょうど 2/29'),
  @('2007-11-20', '2008-02-28', '④終が 2/28（上の 1日 前）'),
  @('2007-12-31', '2008-03-01', '③またぐ・終が うるう年・2/29 を またぐ'),

  @('2008-03-01', '2009-03-01', '⑤ちょうど 1年（同じ 月日）'),
  @('2008-03-01', '2009-03-02', '⑤ちょうど 1年 ＋1日'),
  @('2008-03-01', '2009-02-28', '⑤ちょうど 1年 ー1日'),
  @('2009-03-01', '2010-03-01', '⑤ちょうど 1年（どちらも 平年）'),
  @('2009-03-01', '2010-03-02', '⑤ちょうど 1年 ＋1日（どちらも 平年）'),

  @('2007-01-01', '2010-01-01', '⑥年越え 3年（何年を 平均するか）'),
  @('2007-06-01', '2011-06-01', '⑥年越え 4年'),
  @('2008-01-01', '2009-12-31', '⑥年越え 2年（前が うるう年）'),
  @('2009-01-01', '2010-12-31', '⑥年越え 2年（どちらも 平年）'),
  @('2010-01-01', '2013-01-01', '⑥年越え 3年（真ん中に 2012 うるう年）'),
  @('2011-01-01', '2012-01-01', '⑤ちょうど 1年（終が うるう年の 1/1）'),
  @('2012-01-01', '2013-01-01', '⑤ちょうど 1年（始が うるう年の 1/1）'),
  @('2012-03-01', '2013-01-01', '②またぐ・始が うるう年・2/29 より 後'),
  @('2012-01-01', '2012-12-31', '①同じ年・うるう年 2012'),
  @('2000-01-01', '2000-12-31', '①同じ年・2000（400で 割れる＝うるう年）'),
  @('1900-01-01', '1900-12-31', '①同じ年・1900（★Excel は うるう年だと 思って いる★）'),
  @('2100-01-01', '2100-12-31', '①同じ年・2100（100で 割れる＝平年）')
)

$xl = New-Object -ComObject Excel.Application
try {
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  $bk = $xl.Workbooks.Add()
  $sh = $bk.Sheets.Item(1)

  $行 = New-Object System.Collections.Generic.List[string]
  $行.Add('# ★YEARFRAC の basis 1 は 何で 割って いるか★（2026-09-16）')
  $行.Add('#')
  $行.Add('# ★なぜ★ うちは いつでも「年の 平均日数」で 割って いた（365.5 など）')
  $行.Add('#   ★実Excel は 365 と 366 を 使い分けて いる★＝ODDFPRICE の 部品で 7本 違った')
  $行.Add('#')
  $行.Add('# ★分母★ … 日数 ÷ 答え（★これを 見れば 365 か 366 か 分かります★）')
  $行.Add('#   ★実Excel に 割らせて います★（うちで 割ると うちの 丸めが 混ざる）')
  $行.Add('#')
  $行.Add('# ★2つ目の 窓★ `=(式)=0` … `.Value2` は 0 で ない 値に 0 を 返す')
  $行.Add('#')
  $行.Add('# ★どの Excel か★ … 版 ' + $xl.Version + ' ／ build ' + $xl.Build)
  $行.Add('#')
  $行.Add('# 始' + "`t" + '終' + "`t" + '日数' + "`t" + '答え' + "`t" + '分母' + "`t" + '=(式)=0' + "`t" + '型' + "`t" + 'なぜ この 組か')

  $r = 1
  foreach ($g in $組) {
    $始 = $g[0]; $終 = $g[1]; $訳 = $g[2]
    $d1 = 'DATE(' + $始.Substring(0, 4) + ',' + [int]$始.Substring(5, 2) + ',' + [int]$始.Substring(8, 2) + ')'
    $d2 = 'DATE(' + $終.Substring(0, 4) + ',' + [int]$終.Substring(5, 2) + ',' + [int]$終.Substring(8, 2) + ')'
    $式 = '=YEARFRAC(' + $d1 + ',' + $d2 + ',1)'

    $c = $sh.Range('A' + $r)
    $c.Formula = $式
    $v = $c.Value2
    $答 = if ($null -eq $v) { '(空)' } elseif ($v -is [double]) { $v.ToString('R', [Globalization.CultureInfo]::InvariantCulture) } else { [string]$v }
    $型 = if ($null -eq $v) { '(空)' } elseif ($v -is [double]) { 'Double' } elseif ($v -is [string]) { 'String' } else { 'Other' }

    # ★日数は 実Excel に 引かせる★（うちで 数えない＝向こうの 通し番号で）
    $c2 = $sh.Range('C' + $r)
    $c2.Formula = '=' + $d2 + '-' + $d1
    $日数 = [string]$c2.Value2

    # ★分母★ … 日数 ÷ 答え（実Excel に 割らせる）
    $c3 = $sh.Range('D' + $r)
    $c3.Formula = '=(' + $d2 + '-' + $d1 + ')/' + $式.Substring(1)
    $v3 = $c3.Value2
    $分母 = if ($v3 -is [double]) { $v3.ToString('R', [Globalization.CultureInfo]::InvariantCulture) } else { [string]$v3 }

    # ★2つ目の 窓★
    $w = $sh.Range('B' + $r)
    $w.Formula = '=(' + $式.Substring(1) + ')=0'
    $ゼロか = [string]$w.Value2

    $行.Add($始 + "`t" + $終 + "`t" + $日数 + "`t" + $答 + "`t" + $分母 + "`t" + $ゼロか + "`t" + $型 + "`t" + $訳)
    $r++
  }

  [System.IO.File]::WriteAllText($出, ($行 -join "`n") + "`n", (New-Object System.Text.UTF8Encoding($false)))
  Write-Host ('★書いた … ' + $出 + '（' + $組.Count + '行）★')
  $bk.Close($false)
} finally {
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  $t = [Diagnostics.Stopwatch]::StartNew()
  while ((@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -gt 0) -and ($t.Elapsed.TotalSeconds -lt 300)) {
    Start-Sleep -Milliseconds 500
  }
  $t.Stop()
  Write-Host ('★Excel が 消えるまで ' + [math]::Round($t.Elapsed.TotalSeconds, 1) + '秒 ／ 残り ' +
    @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count + '個★')
}
