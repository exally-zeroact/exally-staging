Attribute VB_Name = "Sheet1"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Column = 4 Then
        Target.Offset(0, 1).Value = Target.Value * 1.1
    End If
End Sub
