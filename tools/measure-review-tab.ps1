# 校閲タブの 真値を 実Excel で 測る（読むだけ・保存しない）
$ErrorActionPreference='Stop'
$xl=New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
$wb=$xl.Workbooks.Add(); $ws=$wb.Worksheets.Item(1)
$ws.Range('A1').Value2='あいうえお'; $ws.Range('A2').Value2=123
$ws.Range('B1').Value2='hello world'; $ws.Range('B2').Formula='=A2*2'
"── ブックの保護 ──"
"既定 … 構造={0} 窓={1}" -f $wb.ProtectStructure, $wb.ProtectWindows
$wb.Protect($null,$true,$false)
"守った後 … 構造={0} 窓={1}" -f $wb.ProtectStructure, $wb.ProtectWindows
try { $wb.Worksheets.Add() | Out-Null; "  ★守っていても シートが 足せた★" }
catch { "  ★守ると シートが 足せない★" }
$wb.Unprotect()
"外した後 … 構造={0}" -f $wb.ProtectStructure
try { $wb.Worksheets.Add() | Out-Null; "  外したら シートが 足せた" } catch { "  外しても 足せない" }
"── メモ（旧コメント）──"
$c = $ws.Range('A1').AddComment('これは メモ')
"  作った … 作者='{0}' 中身='{1}'" -f $c.Author, $c.Text()
"  見えているか = {0}（既定は 隠れている）" -f $c.Visible
"  形の大きさ = {0} x {1}" -f [math]::Round($c.Shape.Width,1), [math]::Round($c.Shape.Height,1)
$ws.Range('A1').ClearComments()
"── 範囲の編集を許可する ──"
"  既定の 数 = {0}" -f $ws.Protection.AllowEditRanges.Count
$ws.Protection.AllowEditRanges.Add('はんい1', $ws.Range('A1:B2')) | Out-Null
"  足した後 = {0}（名前='{1}' 範囲={2}）" -f $ws.Protection.AllowEditRanges.Count, `
   $ws.Protection.AllowEditRanges.Item(1).Title, $ws.Protection.AllowEditRanges.Item(1).Range.Address(0,0)
"── ブックの統計情報 ──"
"  シート数 = {0}" -f $wb.Worksheets.Count
"  中身の在るセル = {0}" -f $ws.UsedRange.Cells.Count
"  式のセル = {0}" -f $(try { $ws.Cells.SpecialCells(-4123).Count } catch { 0 })
"  字のセル = {0}" -f $(try { $ws.Cells.SpecialCells(2,2).Count } catch { 0 })
"  数のセル = {0}" -f $(try { $ws.Cells.SpecialCells(2,1).Count } catch { 0 })
$wb.Close($false); $xl.Quit()
