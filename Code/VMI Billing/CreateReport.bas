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
            SourceData:=Sheets("Drop In").Range(Cells(1, 1), Cells(TotalRows, 16)), _
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

    aHeaders = Array("Cust #", _
                     "Plant", _
                     "2nd Tier Supplier", _
                     "Contract #", _
                     "Invoice Date", _
                     "VMI Order #", _
                     "Order Line", _
                     "Stock Code", _
                     "Description", _
                     "Qty", _
                     "Price", _
                     "Extended Price", _
                     "Invoice Number", _
                     "2nd Tier Supplier Invoice #", _
                     "2nd Tier Supplier Inv date", _
                     "Packing List No.")


    aFields = Array( _
              "Period Covered", _
              "Total", _
              "PO Number", _
              "Release", _
              "Route Code", _
              "Invoice Number")

    Sheets(SheetName).Select
    ActiveSheet.ListObjects(1).Unlist
    Rows(1).Delete

    Rows("1:11").EntireRow.Insert

    'Sheet Title
    Range("A1").Value = "WESCO - VMI - Monthly Summary Invoice"
    Range("A1").Font.Size = 14
    Range("A1").Font.Bold = True

    'Add Header Fields
    Range(Cells(4, 2), Cells(UBound(aFields) + 4, 2)) = WorksheetFunction.Transpose(aFields)
    For Each Rng In Range("B4:C9")
        Rng.BorderAround xlContinuous
    Next

    'Vendor ID
    With Range("H9:H10")
        .Font.Name = "Arial"
        .Font.Size = 14
        .Font.Bold = True
        .NumberFormat = "@"
        .Interior.Color = 65535
        .BorderAround xlContinuous, xlThin
    End With
    Range("H9").Value = "Vendor ID"
    Range("H10").Value = "000132199002"

    'Remit Address
    With Range("H3")
        .Value = "Remit Address"
        .HorizontalAlignment = xlCenter
        .Font.Size = 14
        .Font.Bold = True
    End With
    Range("H4").Value = "WESCO Distribution"
    Range("H5").Value = "10101 Claude Freeman Dr"
    Range("H6").Value = "Suite 220 N."
    Range("H7").Value = "Charlotte NC, 28262 "
    With Range("H4:H7")
        .BorderAround xlContinuous, xlThin
        .Font.Size = 12
        .Font.Bold = True
    End With
    Range("H3:H7").Font.Name = "Arial"

    'Plant Name
    With Range("B3")
        .Formula = "=IFERROR(VLOOKUP(B12,Master!A:D,2,FALSE),"""")"
        .Font.Name = "Arial"
        .Font.Size = 14
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
    End With
    Range("B3").Value = Range("B3").Value
    Range("B3:C3").Merge

    'Period Covered
    With Range("B4:C4")
        .Font.Name = "Arial"
        .Font.Size = 9
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
    End With
    Range("C4").Value = Format(dt, "mmm-yy")

    'Total
    Range("C5").Formula = "=SUM(K:K)"

    'PO Number
    Range("C6").Formula = "=VLOOKUP(B12,Master!A:C,3,FALSE)"

    'Release
    Range("C7").Formula = "=IF(VLOOKUP(B12,Master!A:E,5,FALSE)=0,"""",IFERROR(VLOOKUP(B12,Master!A:E,5,FALSE),""""))"

    'Route Code
    Range("C8").Formula = "=IF(VLOOKUP(B12,Master!A:F,6,FALSE)=0,"""",IFERROR(VLOOKUP(B12,Master!A:F,6,FALSE),""""))"

    'Invoice Number
    Range("C9").Formula = "=VLOOKUP(B12,Master!A:D,4,FALSE)"
    Range("C9").Value = Range("C9").Text & Format(dt, "mmyy")

    'Total-Invoice Formatting
    With Range("B4:C9")
        .Value = .Value
        .Font.Name = "Arial"
        .Font.Size = 12
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .Interior.Color = 65535
    End With

    'Add Column Headers
    Range(Cells(11, 1), Cells(11, UBound(aHeaders) + 1)) = aHeaders
    Range(Cells(11, 1), Cells(11, UBound(aHeaders) + 1)).HorizontalAlignment = xlCenter

    'Add Column Header Borders
    With Range(Cells(11, 1), Cells(11, UBound(aHeaders) + 1))
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
    With Cells(11, UBound(aHeaders) + 1).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -6908266
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Range("B5:C5").Interior.Color = 16777164

    iCols = ActiveSheet.UsedRange.Columns.Count + 1
    iRows = ActiveSheet.UsedRange.Rows.Count

    'Add eStock Data Vlookup if "Route Code" is not blank
    If Range("C8").Text <> "" Then
        Cells(11, iCols).Value = "VLOOKUP"
        'Vlookup pulls in cost from the eStock sheet so it
        'can becompared to the cost on the current sheet
        'Column G contains Stock Codes (Item numbers)
        Range(Cells(11, 8), Cells(iRows, 8)).Cells.Insert
        Range(Cells(11, 8), Cells(iRows, 8)).Formula = "=""="""""" & I11 & """""""""
        Range(Cells(11, 8), Cells(iRows, 8)).Value = Range(Cells(11, 8), Cells(iRows, 8)).Value
        Columns(9).Delete

        Range(Cells(12, iCols), Cells(iRows, iCols)).Formula = "=IFERROR(VLOOKUP(H12,'VMI eStock'!A:K,11,FALSE),"""")"
        Range(Cells(12, iCols), Cells(iRows, iCols)).Value = Range(Cells(12, iCols), Cells(iRows, iCols)).Value

        For i = 12 To ActiveSheet.UsedRange.Rows.Count
            If Cells(i, iCols).Value <> Cells(i, 11).Value Then
                Cells(i, iCols).Interior.Color = 5263615
            End If
        Next
    End If

    ActiveSheet.UsedRange.Columns.AutoFit
    Columns(1).ColumnWidth = 7.86

    FillInfo FunctionName:="Template", _
             FileDate:="", _
             Parameters:="SheetName: " & SheetName, _
             ExecutionTime:=Timer - StartTime, _
             Result:="Complete"
End Sub

