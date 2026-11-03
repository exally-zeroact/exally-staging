# 余白の 決まった形（標準・広い・狭い）を 実Excel から 測る
# ★COM に「標準/広い/狭い」を 直に 呼ぶ 道が 無い★ので、
#   ExecuteExcel4Macro や リボンの 既定値は 読めない。
#   代わりに ★新しいブックの 既定（＝標準）★を 測り、他は 出せない事を 書き残す。
$ErrorActionPreference='Stop'
$xl=New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
$wb=$xl.Workbooks.Add(); $ps=$wb.Worksheets.Item(1).PageSetup
"標準（新しいブックの 既定）… 上={0}pt 下={1}pt 左={2}pt 右={3}pt ヘッダー={4}pt フッター={5}pt" -f `
  $ps.TopMargin,$ps.BottomMargin,$ps.LeftMargin,$ps.RightMargin,$ps.HeaderMargin,$ps.FooterMargin
"  cm に すると … 上={0} 下={1} 左={2} 右={3}" -f `
  [math]::Round($ps.TopMargin/28.3464566929134,3), [math]::Round($ps.BottomMargin/28.3464566929134,3), `
  [math]::Round($ps.LeftMargin/28.3464566929134,3), [math]::Round($ps.RightMargin/28.3464566929134,3)
try { $xl.ExecuteExcel4Macro('PAGE.SETUP()') ; "  ExecuteExcel4Macro は 呼べた" }
catch { "  ★ExecuteExcel4Macro からは 読めない★" }
$wb.Close($false); $xl.Quit()
