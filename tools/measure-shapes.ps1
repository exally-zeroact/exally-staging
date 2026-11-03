# 図形・テキストボックス・画像の 既定を 実Excel で 測る（読むだけ・保存しない）
$ErrorActionPreference='Stop'
$xl=New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
$wb=$xl.Workbooks.Add(); $ws=$wb.Worksheets.Item(1)
"── 図形（四角）──"
$sh = $ws.Shapes.AddShape(1, 10, 10, 100, 50)   # 1 = msoShapeRectangle
"  名前='{0}' 大きさ={1}x{2} 場所={3},{4}" -f $sh.Name,$sh.Width,$sh.Height,$sh.Left,$sh.Top
"  塗り={0} 線の色={1} 線の太さ={2}" -f $sh.Fill.ForeColor.RGB, $sh.Line.ForeColor.RGB, $sh.Line.Weight
"  文字は 入るか = {0}" -f ($null -ne $sh.TextFrame2)
"  重なりの 順（ZOrderPosition）= {0}" -f $sh.ZOrderPosition
"── テキスト ボックス ──"
$tb = $ws.Shapes.AddTextbox(1, 10, 80, 120, 30)  # 1 = msoTextOrientationHorizontal
"  名前='{0}' 大きさ={1}x{2}" -f $tb.Name,$tb.Width,$tb.Height
"  塗り={0}（塗りが 見えるか={1}）／線が 見えるか={2}" -f $tb.Fill.ForeColor.RGB, $tb.Fill.Visible, $tb.Line.Visible
"  重なりの 順 = {0}" -f $tb.ZOrderPosition
"── 重なりを 動かす ──"
$sh.ZOrder(0) | Out-Null    # 0 = msoBringToFront
"  四角を 前面へ … 四角={0} テキスト={1}" -f $sh.ZOrderPosition, $tb.ZOrderPosition
$sh.ZOrder(1) | Out-Null    # 1 = msoSendToBack
"  四角を 背面へ … 四角={0} テキスト={1}" -f $sh.ZOrderPosition, $tb.ZOrderPosition
"── 図形の 種類（番号）──"
"  四角=1 / 角丸四角=5 / 円=9 / 三角=7 / 矢印(右)=33 / 星5=92 / 線=(AddLine)"
"── 全部の 数 ──"
"  Shapes.Count = {0}" -f $ws.Shapes.Count
$wb.Close($false); $xl.Quit()
