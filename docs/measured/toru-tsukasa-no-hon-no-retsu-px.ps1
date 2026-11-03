# toru-tsukasa-no-hon-no-retsu-px.ps1 -- ★司さんの本の 列の px を 実Excel で 取る★（135）（2026-10-04・経営者）
#   ★中身の 字は 読まない★（列幅と px と 隠し だけ）／%TEMP% に 写して 読み取り専用で 開き 保存しない・写しは 消す
#   ★0 の 窓★ ... この 道具は マスの 値を 1つも 読まない（列の 幅だけ）＝tests/monosashi-mado.test.mjs の 免除の 条件「数の 答えを 取らない」に 当たる
#   使い方: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <この道具> -道 <xlsb> -出す先 <tsv>
param([string]$道 = '', [string]$出す先 = '')
# 司さんの本を %TEMP% に写して 読むだけで開き、各板の 使っている列の 幅（字）と px（.Width×96/72）だけを書く。中身の字は読まない。
$写し = Join-Path $env:TEMP 'exally-retsu-px.xlsb'
Copy-Item -LiteralPath $道 -Destination $写し -Force
$xl = New-Object -ComObject Excel.Application
$bk = $null; $sh = $null; $u = $null; $col = $null; $st = $null
$行 = New-Object System.Collections.ArrayList
try {
  $xl.Visible = $false; $xl.DisplayAlerts = $false
  $bk = $xl.Workbooks.Open($写し, $false, $true)
  $st = $bk.Styles.Item('Normal')
  Write-Host ('本の 既定の 字体 ' + [string]$st.Font.Name + ' ' + [string]$st.Font.Size); $st = $null
  $n = [int]$bk.Sheets.Count
  for ($i = 1; $i -le $n; $i++) {
    $sh = $bk.Sheets.Item($i)
    $u = $sh.UsedRange
    $c0 = [int]$u.Column; $cn = [int]$u.Columns.Count
    $u = $null
    for ($c = 1; $c -le ($c0 + $cn - 1); $c++) {
      $col = $sh.Columns.Item($c)
      [void]$行.Add(([string]$i + "`t" + ($c - 1) + "`t" + [string]$col.ColumnWidth + "`t" + [string]([math]::Round([double]$col.Width * 96 / 72, 2)) + "`t" + [string]$col.Hidden))
      $col = $null
    }
    $sh = $null
  }
  $bk.Close($false)
} finally { $col = $null; $u = $null; $sh = $null; $st = $null; $bk = $null; $xl.Quit(); $xl = $null }
Remove-Item -LiteralPath $写し -Force
$書 = New-Object System.IO.StreamWriter($出す先, $false, (New-Object System.Text.UTF8Encoding($false)))
$書.NewLine = "`n"
$書.WriteLine("板(1から)`t列(0から)`t列幅(字)`tpx`t隠し")
foreach ($l in $行) { $書.WriteLine($l) }
$書.Close()
Write-Host ('★書いた 列★ ' + $行.Count)
