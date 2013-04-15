Attribute VB_Name = "Module1"
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

Sub CreatePivTables()
    Dim vCell As Variant
    Dim TotalRows As Long
    Dim TotalCols As Integer
    Dim SheetName As String
    Dim StartTime As Double

    StartTime = Timer
    Sheets("Drop In").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    TotalCols = ActiveSheet.UsedRange.Columns.Count

    ActiveWorkbook.PivotCaches.Create( _
            SourceType:=xlDatabase, _
            SourceData:=Sheets("Drop In").Range(Cells(1, 1), Cells(TotalRows, 15)), _
            Version:=xlPivotTableVersion14).CreatePivotTable _
            TableDestination:="PivotTable!R1C1", _
            TableName:="PivotTable1", _
            DefaultVersion:=xlPivotTableVersion14

    Sheets("PivotTable").Select
    With ActiveSheet.PivotTables("PivotTable1")
        .PivotFields("Plant").Orientation = xlRowField
        .PivotFields("Plant").Position = 1
        .AddDataField .PivotFields("Extended Price"), "Sum of Extended Price", xlSum
        .ColumnGrand = False
    End With

    For Each vCell In ActiveSheet.PivotTables(1).DataBodyRange
        SheetName = vCell.Offset(0, -1).Text

        vCell.ShowDetail = True
        ActiveSheet.Name = SheetName
        Sheets("PivotTable").Select
    Next


End Sub

Sub Template(SheetName As String)
    Dim iCols As Integer
    Dim iRows As Long
    Dim StartTime As Double
    Dim aHeaders As Variant
    Dim aFields As Variant
    Dim Rng As Variant
    Dim dt As Date
    Dim i As Long

    StartTime = Timer
    dt = DateAdd("m", -1, Date)
    aHeaders = Array( _
               "Plant", _
               "Vendor Code", _
               "2nd Tier Supplier", _
               "Invoice Date", _
               "VMI Order #", _
               "Order Line", _
               "Stock Code", _
               "Description", _
               "Qty", _
               "Price", _
               "Extended Price", _
               "Invoice Number", _
               "Supplier Inv#", _
               "Supplier Inv Date", _
               "Packing Slip#")

    aFields = Array( _
              "Period Covered", _
              "Total", _
              "PO Number", _
              "Route Code", _
              "Invoice Number")

    Sheets(SheetName).Select
    ActiveSheet.ListObjects(1).Unlist
    Rows(1).Delete

    Rows("1:7").EntireRow.Insert

    'Add Header Fields
    Range(Cells(2, 2), Cells(UBound(aFields) + 2, 2)) = WorksheetFunction.Transpose(aFields)
    For Each Rng In Range("B1:C6")
        Rng.BorderAround xlContinuous
    Next

    'Vendor ID
    With Range("H1:H2")
        .Font.Name = "Arial"
        .Font.Size = 14
        .Font.Bold = True
        .NumberFormat = "@"
        .Interior.Color = 65535
        .BorderAround xlContinuous, xlMedium
    End With
    Range("H1").Value = "Vendor ID"
    Range("H2").Value = "000132199002"

    'Plant Name
    With Range("B1")
        .Formula = "=IFERROR(VLOOKUP(A8,Master!A:D,2,FALSE),"""")"
        .Font.Name = "Arial"
        .Font.Size = 14
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
    End With
    Range("B1").Value = Range("B1").Value
    Range("B1:C1").Merge

    'Period Covered
    With Range("B2:C2")
        .Font.Name = "Arial"
        .Font.Size = 9
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
    End With
    Range("C2").Value = Format(dt, "mmm")

    'Total
    Range("C3").Formula = "=SUM(K:K)"

    'PO Number
    Range("C4").Formula = "=VLOOKUP(A8,Master!A:C,3,FALSE)"

    'Route Code
    Range("C5").Formula = "=IF(VLOOKUP(A8,Master!A:E,5,FALSE)=0,"""",IFERROR(VLOOKUP(A8,Master!A:E,5,FALSE),""""))"

    'Invoice Number
    Range("C6").Formula = "=VLOOKUP(A8,Master!A:D,4,FALSE)"
    Range("C6").Value = Range("C6").Text & Format(dt, "mmyy")

    'Total, PO Number, Route Code, Invoice Number Formatting
    With Range("B3:C6")
        .Value = .Value
        .Font.Name = "Arial"
        .Font.Size = 12
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .Interior.Color = 65535
    End With

    'Add Column Headers
    Range(Cells(7, 1), Cells(7, UBound(aHeaders) + 1)) = aHeaders
    Range(Cells(7, 1), Cells(7, UBound(aHeaders) + 1)).HorizontalAlignment = xlCenter

    'Add Column Header Borders
    With Range(Cells(7, 1), Cells(7, UBound(aHeaders) + 1))
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Color = -6908266
        .Borders(xlEdgeTop).TintAndShade = 0
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = -6908266
        .Borders(xlEdgeBottom).TintAndShade = 0
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Font.Color = -8388608
        .Font.TintAndShade = 0
        .Font.Name = "Arial"
        .Font.Size = 10
        .Font.Bold = True
    End With
    With Cells(7, UBound(aHeaders) + 1).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -6908266
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Range("B3:C3").Interior.Color = 16777164

    iCols = ActiveSheet.UsedRange.Columns.Count + 1
    iRows = ActiveSheet.UsedRange.Rows.Count

    'Add eStock Data Vlookup
    If Range("C5").Text <> "" Then
        Cells(7, iCols).Value = "VLOOKUP"
        'Vlookup pulls in cost from the eStock sheet so it
        'can becompared to the cost on the current sheet
        'Column G contains Stock Codes (Item numbers)
        Cells(8, iCols).Formula = "=IFERROR(VLOOKUP(G8,'VMI eStock'!A:K,11,FALSE),"""")"
        On Error Resume Next
        Cells(8, iCols).AutoFill Destination:=Range(Cells(8, iCols), Cells(iRows, iCols))
        On Error GoTo 0
        For i = 8 To ActiveSheet.UsedRange.Rows.Count
            If Cells(i, iCols).Value <> Cells(i, 10).Value Then
                Cells(i, iCols).Interior.Color = 5263615
            End If
        Next
    End If


    ActiveSheet.UsedRange.Columns.EntireColumn.AutoFit

    FillInfo FunctionName:="Template", _
             FileDate:="", _
             Parameters:="SheetName: " & SheetName, _
             ExecutionTime:=Timer - StartTime, _
             Result:="Complete"
End Sub

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

            'ThisWorkbook.Activate
        End If
    Next
End Sub

Sub SaveBooks()
    Dim s As Variant
    Dim dt As Date
    Dim SaveDialog As FileDialog

    Set SaveDialog = Application.FileDialog(msoFileDialogSaveAs)
    dt = DateAdd("m", -1, Date)

    For Each s In ThisWorkbook.Sheets
        If s.Name <> "Drop In" And _
           s.Name <> "PivotTable" And _
           s.Name <> "Info" And _
           s.Name <> "Macro" And _
           s.Name <> "VMI eStock" And _
           s.Name <> "Master" Then
            s.Copy
            If Range("C5").Text <> "" Then
                Columns(ActiveSheet.UsedRange.Columns.Count).Delete
            End If
            SaveDialog.InitialFileName = s.Name & "_" & Format(dt, "mmm_yyyy")
            SaveDialog.Show
            If SaveDialog.SelectedItems.Count > 0 Then
                ActiveWorkbook.SaveAs SaveDialog.SelectedItems.Item(1), xlOpenXMLWorkbook
            End If
            Application.DisplayAlerts = False
            ActiveWorkbook.Close
            Application.DisplayAlerts = True
        End If
    Next
End Sub

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
