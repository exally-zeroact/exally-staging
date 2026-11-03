Attribute VB_Name = "Module2"
Option Explicit

' CSVÇ ì«Ç›çûÇÒÇ≈ ì\ÇÈ
Sub CSVéÊÇËçûÇ›()
    Dim f As String
    f = Application.GetOpenFilename("CSV,*.csv")
    If f = "False" Then Exit Sub
    Workbooks.Open Filename:=f
    ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Sheets("éÊçû").Range("A1")
    ActiveWorkbook.Close SaveChanges:=False
End Sub
