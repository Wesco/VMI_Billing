Attribute VB_Name = "CreateReport"
Option Explicit

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

