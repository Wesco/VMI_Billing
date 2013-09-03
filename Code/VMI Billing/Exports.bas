Attribute VB_Name = "Exports"
Option Explicit

Sub SaveThisBook()
    Dim s As Variant
    Dim Wkbk As Variant

    For Each s In ThisWorkbook.Sheets
        If s.Name <> "Drop In" And _
           s.Name <> "Macro" And _
           s.Name <> "PivotTable" And _
           s.Name <> "Info" And _
           s.Name <> "VMI eStock" And _
           s.Name <> "Master" Then

            If TypeName(Wkbk) = "Empty" Then
                s.Copy
                Set Wkbk = ActiveWorkbook
            Else
                s.Copy After:=Wkbk.Sheets(Wkbk.Sheets.Count)
            End If
        End If
    Next
End Sub

Sub SaveCombinedBilling()
    Dim Result As Long
    Dim StartTime As Double
    Dim PrevDispAlert As Boolean
    Dim Complete As Boolean
    Dim SaveDialog As FileDialog
    Dim dt As Date

    Set SaveDialog = Application.FileDialog(msoFileDialogSaveAs)
    dt = DateAdd("m", -1, Date)
    PrevDispAlert = Application.DisplayAlerts
    Application.DisplayAlerts = False
    StartTime = Timer
    Result = vbNo

    Result = MsgBox("Save combined billing info?", vbYesNo, "Save Sheet")
    If Result = vbYes Then
        Sheets("Drop In").Copy

        SaveDialog.InitialFileName = "ALLDATA_" & UCase(Format(dt, "mmm")) & "_" & Format(dt, "yyyy")
        SaveDialog.Show

        If SaveDialog.SelectedItems.Count > 0 Then
            ActiveWorkbook.SaveAs SaveDialog.SelectedItems.Item(1), xlOpenXMLWorkbook
            Complete = True
        Else
            Complete = False
        End If

        ActiveWorkbook.Close
    Else
        Complete = False
    End If

    If Complete = True Then
        FillInfo FunctionName:="SaveCombinedBilling", _
                 FileDate:="", _
                 Parameters:="", _
                 ExecutionTime:=Timer - StartTime, _
                 Result:="Complete"
    Else
        FillInfo FunctionName:="SaveCombinedBilling", _
                 FileDate:="", _
                 Parameters:="", _
                 ExecutionTime:=Timer - StartTime, _
                 Result:="User Canceled"
    End If

    Application.DisplayAlerts = PrevDispAlert
End Sub

