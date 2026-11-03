# ★実Excel に BAHTTEXT の 答えを 出させる★（読むだけ・新規ブック・保存しない）
#   ★自分で 考えた 規則で 作らない★＝実物の 答えを 台帳に する
param([string]$Out)
$ErrorActionPreference='Stop'
$数 = @(
  0, 1, 2, 3, 4, 5, 6, 7, 8, 9,
  10, 11, 12, 20, 21, 22, 25, 30, 31, 41, 51, 91,
  100, 101, 111, 120, 121, 199, 200, 201,
  1000, 1001, 1011, 1100, 1111, 2000, 2001,
  10000, 10001, 11000, 21000, 100000, 100001, 110000,
  1000000, 1000001, 1000011, 2000000, 1234567, 12345678, 123456789,
  1000000000, 1000000000000,
  0.5, 0.25, 0.01, 0.05, 0.1, 0.99, 1.5, 1.25, 1.75, 21.21,
  123.45, 1000.5, 1234.56, 999999.99,
  -1, -0.5, -21, -1234.56, -1000000,
  0.001, 0.005, 0.994, 0.995, 1.005, 2.675
)
$xl=$null
try{
  $xl=New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
  $wb=$xl.Workbooks.Add(); $sh=$wb.Worksheets.Item(1)
  $行=1
  $出=@()
  foreach($n in $数){
    $sh.Cells($行,1).Value2 = [double]$n
    $sh.Cells($行,2).Formula = "=BAHTTEXT(A$行)"
    $行++
  }
  Start-Sleep -Milliseconds 300
  $行=1
  foreach($n in $数){
    $t = $sh.Cells($行,2).Text
    $出 += [pscustomobject]@{ n = $n; text = $t }
    $行++
  }
  $出 | ConvertTo-Json -Depth 3 | Out-File -FilePath $Out -Encoding utf8
  "書いた … $Out（$($出.Count)件）"
  $出 | Select-Object -First 6 | ForEach-Object { "  $($_.n) → $($_.text)" }
}catch{ "★落ちた★ $($_.Exception.Message)" }
finally{ if($xl){ try{foreach($z in @($xl.Workbooks)){$z.Close($false)}}catch{}; try{$xl.Quit()}catch{} } }
