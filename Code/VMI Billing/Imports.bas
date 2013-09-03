Attribute VB_Name = "Imports"
Option Explicit

Sub ImportMaster()
    Dim sPath As String
    Dim PrevDispAlert As Boolean

    sPath = "\\br3615gaps\gaps\Duke\VMI Master.xlsx"
    PrevDispAlert = Application.DisplayAlerts

    On Error GoTo OPEN_FAILED
    Workbooks.Open sPath
    On Error GoTo 0
    ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Sheets("Master").Range("A1")

    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = PrevDispAlert
    Exit Sub

OPEN_FAILED:
    'File Not Found
    Err.Raise 53
End Sub
