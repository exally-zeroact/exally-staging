Attribute VB_Name = "Module3"
Option Explicit

' 毎月：名前で 並べ替えて 税込の列を 足し、控えを 右へ 写す
Sub 月次の並べ替え()
    Range("A1:D100").Sort Key1:=Range("B2"), Order1:=xlDescending
    Range("E1").Value = "税込"
    Range("E2:E100").Formula = "=D2*1.1"
    Range("A1:E100").Copy Destination:=Range("H1")
End Sub
