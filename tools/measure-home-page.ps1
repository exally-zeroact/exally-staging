# ホーム／ページレイアウトの 真値を 実Excel で 測る（読むだけ・保存しない）
$ErrorActionPreference='Stop'
$xl=New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
$wb=$xl.Workbooks.Add(); $ws=$wb.Worksheets.Item(1); $ps=$ws.PageSetup
"── フォント ──"
"既定の フォント = '{0}' / 大きさ {1}" -f $xl.StandardFont, $xl.StandardFontSize
"A1 の フォント  = '{0}' / {1}" -f $ws.Range('A1').Font.Name, $ws.Range('A1').Font.Size
"── 文字の 向き（Orientation）──"
foreach($o in 0,45,90,-45,-90,-4166,-4171){
  try { $ws.Range('A1').Orientation=$o; "  {0} を入れたら {1}" -f $o,$ws.Range('A1').Orientation }
  catch { "  {0} は ★入らない★" -f $o }
}
$ws.Range('A1').Orientation=0
"  （-4166=縦書き / -4171=横書き）"
"── 余白（既定と 決まった 形）──"
"既定 … 上={0} 下={1} 左={2} 右={3} ヘッダー={4} フッター={5}（pt）" -f `
  $ps.TopMargin,$ps.BottomMargin,$ps.LeftMargin,$ps.RightMargin,$ps.HeaderMargin,$ps.FooterMargin
"  1cm = {0} pt" -f $xl.CentimetersToPoints(1)
"  1inch = {0} pt" -f $xl.InchesToPoints(1)
"── 用紙 ──"
"既定の PaperSize = {0}（9=A4 / 8=A3 / 11=A5 / 1=Letter / 13=B5）" -f $ps.PaperSize
foreach($p in 9,8,11,1,13){ try { $ps.PaperSize=$p; "  {0} を入れたら {1}" -f $p,$ps.PaperSize } catch { "  {0} は 入らない" -f $p } }
$ps.PaperSize=9
"── 印刷範囲 ──"
"既定の PrintArea = '{0}'" -f $ps.PrintArea
$ps.PrintArea = 'A1:C10'
"  入れた後 = '{0}'" -f $ps.PrintArea
$ps.PrintArea = ''
"  空に した後 = '{0}'" -f $ps.PrintArea
"── シートのオプション（表示 と 印刷）──"
"枠線 … 表示={0} 印刷={1}" -f $xl.ActiveWindow.DisplayGridlines, $ps.PrintGridlines
"見出し … 表示={0} 印刷={1}" -f $xl.ActiveWindow.DisplayHeadings, $ps.PrintHeadings
"── 拡大縮小印刷 ──"
"Zoom={0} FitToPagesWide={1} FitToPagesTall={2}" -f $ps.Zoom, $ps.FitToPagesWide, $ps.FitToPagesTall
$wb.Close($false); $xl.Quit()
