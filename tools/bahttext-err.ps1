# ★変な物を 渡した時に 実Excel が 何を 返すか★（読むだけ・保存しない）
param([string]$Out)
$ErrorActionPreference='Stop'
$xl=$null
try{
  $xl=New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
  $wb=$xl.Workbooks.Add(); $sh=$wb.Worksheets.Item(1)
  $組 = @(
    @{ な='文字';       式='=BAHTTEXT("abc")' },
    @{ な='数の字';     式='=BAHTTEXT("123")' },
    @{ な='空の セル';  式='=BAHTTEXT(Z99)' },
    @{ な='空の 字';    式='=BAHTTEXT("")' },
    @{ な='TRUE';       式='=BAHTTEXT(TRUE)' },
    @{ な='FALSE';      式='=BAHTTEXT(FALSE)' },
    @{ な='エラー';     式='=BAHTTEXT(1/0)' },
    @{ な='1e20';       式='=BAHTTEXT(1E+20)' },
    @{ な='1e21';       式='=BAHTTEXT(1E+21)' }
  )
  $行=1; foreach($k in $組){ $sh.Cells($行,1).Formula = $k.式; $行++ }
  Start-Sleep -Milliseconds 300
  $行=1; $出=@()
  foreach($k in $組){ $出 += [pscustomobject]@{ な=$k.な; 式=$k.式; text=$sh.Cells($行,1).Text }; $行++ }
  $出 | ConvertTo-Json -Depth 3 | Out-File -FilePath $Out -Encoding utf8
  $出 | ForEach-Object { "  $($_.な.PadRight(10)) $($_.式.PadRight(24)) → $($_.text)" }
}catch{ "★落ちた★ $($_.Exception.Message)" }
finally{ if($xl){ try{foreach($z in @($xl.Workbooks)){$z.Close($false)}}catch{}; try{$xl.Quit()}catch{} } }
