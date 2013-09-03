Attribute VB_Name = "Program"
Option Explicit

Sub Main()
    Dim NuclearResult As VbMsgBoxResult
    Dim FossilHydroResult As VbMsgBoxResult
    Dim iRows As Long   'TotalRows
    Dim s As Variant

    On Error GoTo FAILED_IMPORT_MASTER
    ImportMaster
    On Error GoTo 0

    Sheets("Drop In").Select

    On Error GoTo FAILED_IMPORT
    'Import first Billing Info sheet
    NuclearResult = MsgBox("Open the Nuclear Billing Info file", vbYesNo, "File Import")
    If NuclearResult = vbYes Then
        UserImportFile Sheets("Drop In").Range("A1")
        iRows = ActiveSheet.UsedRange.Rows.Count
    End If

    'Import second Billing Info sheet
    FossilHydroResult = MsgBox("Open the Fossil/Hydro Billing Info file", vbYesNo, "File Import")
    If FossilHydroResult = vbYes Then
        If NuclearResult = vbYes Then
            UserImportFile Sheets("Drop In").Cells(iRows + 1, 1)
            Rows(iRows + 1).Delete
        Else
            UserImportFile Sheets("Drop In").Range("A1")
            iRows = ActiveSheet.UsedRange.Rows.Count
        End If
    End If

    'Import VMI eStock Cost Data
    MsgBox "Open the VMI eStock Data file"
    UserImportFile Sheets("VMI eStock").Range("A1")
    On Error GoTo 0

    SaveCombinedBilling
    CreatePivTables

    For Each s In ThisWorkbook.Sheets
        If s.Name <> "Drop In" And _
           s.Name <> "PivotTable" And _
           s.Name <> "Info" And _
           s.Name <> "Macro" And _
           s.Name <> "VMI eStock" And _
           s.Name <> "Master" Then
            Template s.Name
        End If
    Next

    Exit Sub

FAILED_IMPORT:
    MsgBox "User cancelled file import. Macro aborted.", vbOKOnly, "Error"
    Exit Sub

FAILED_IMPORT_MASTER:
    MsgBox "Unable to import VMI Master. Macro aborted.", vbOKOnly, "Error"
End Sub

Sub CleanUp()
    Dim s As Variant

    Application.DisplayAlerts = False
    For Each s In ThisWorkbook.Sheets
        If s.Name <> "Drop In" And _
           s.Name <> "PivotTable" And _
           s.Name <> "Info" And _
           s.Name <> "Macro" And _
           s.Name <> "VMI eStock" And _
           s.Name <> "Master" Then
            s.Delete
        End If
    Next
    Application.DisplayAlerts = True

    Sheets("Drop In").Cells.Delete
    Sheets("PivotTable").Cells.Delete
    Sheets("Info").Cells.Delete
    Sheets("VMI eStock").Cells.Delete
    ActiveWorkbook.Save
End Sub
