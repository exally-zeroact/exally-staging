Attribute VB_Name = "Module1"
Option Explicit

' 毎月の締め：先月ぶんを 別のシートに 写して 合計を 出す
Sub 月次締め()
    Dim 最終行 As Long
    最終行 = Cells(Rows.Count, 1).End(xlUp).Row
    Sheets("集計").Range("A1:D" & 最終行).Value = Sheets("明細").Range("A1:D" & 最終行).Value
    Sheets("集計").Range("E1").Formula = "=SUM(D2:D" & 最終行 & ")"
    MsgBox "締めました"
End Sub

Sub 印刷()
    Sheets("集計").PrintOut Copies:=1
End Sub
