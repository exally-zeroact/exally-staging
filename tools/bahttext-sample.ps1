# ★BAHTTEXT の 見本を 実Excel に 作らせる★（式のまま 保存＝うちで 計算し直せる）
param([string]$Out)
$ErrorActionPreference='Stop'
$数 = @(123.45, -1234.56, 1.005, 0, 1000000, 0.5, 21)
$xl=$null
try{
  $xl=New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
  $wb=$xl.Workbooks.Add(); $sh=$wb.Worksheets.Item(1); $sh.Name='タイ語'
  $r=1
  foreach($n in $数){
    $sh.Cells($r,1).Value2=[double]$n
    $sh.Cells($r,2).Formula="=BAHTTEXT(A$r)"   # ★うちが 計算し直す 式★
    $sh.Cells($r,3).Value2=$sh.Cells($r,2).Text # ★実Excel が 出した 答え（字のまま）★
    $r++
  }
  $sh.Cells($r,2).Formula="=BAHTTEXT(A99)"     # ★空の セル★
  $sh.Cells($r,3).Value2=$sh.Cells($r,2).Text
  $wb.SaveAs($Out, 51)
  "保存した … $Out（$($数.Count+1)行）"
}catch{ "★落ちた★ $($_.Exception.Message)" }
finally{ if($xl){ try{foreach($z in @($xl.Workbooks)){$z.Close($false)}}catch{}; try{$xl.Quit()}catch{} } }
