Attribute VB_Name = "bScrapNew"
Sub acFillNewScrap()

    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim maxDate As Date
    
    ' Open the workbook
    'Set wb = Workbooks.Open("H:\KPI SLC\2026\06-Jun\KPI - Scrap Trend Report - June 2026.xlsx")
    Set wb = Workbooks.Open("C:\Users\jh00082083\Documents\aaWeekly Reports\KPI - Scrap Trend Report - June 2026.xlsx")
    Set ws = wb.Sheets("Master Data - Scrap")
    
     ws.Activate
    ' Clear any existing filters
    ws.AutoFilterMode = False
      
      
      
    Worksheets("Master Data - Scrap").Activate
 
' Find the last used row in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Find the maximum date in column A
    maxDate = Application.WorksheetFunction.Max(ws.Range("A2:A" & lastRow))
    

ws.Range("$A$1:$X$" & lastRow).AutoFilter Field:=1, Criteria1:=maxDate
  

' Select visible cells in columns B to J (excluding header)
    On Error Resume Next
    ws.Range("A2:X" & lastRow).SpecialCells(xlCellTypeVisible).Select
    On Error GoTo 0
    
    Selection.ClearContents
    
    
        Rows("1:1").Select
     ActiveSheet.ShowAllData
    Range("A2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1).Select
    ActiveCell.Offset(0, 3).Select
   
Call UpdateTemplate
Call FillInScrap
Call FillScrapProd
    

End Sub

 Sub UpdateTemplate()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim maxDate As Date
    Dim rng As Range
    
    Set rng = ActiveSheet.UsedRange
    
    


Workbooks.Open "C:\Users\jh00082083\Documents\aaTemplates AK\New Scrap Template AK.xlsx"
Worksheets("Combined MAD Reports").Activate

    
    ' Refresh the connection in the opened workbook
    ActiveWorkbook.Connections("Query - MAD Combined Reports").Refresh
    DoEvents





  Windows("New Scrap Template AK").Activate
Worksheets("Combined MAD Reports").Activate

    ActiveWorkbook.Worksheets("Combined MAD Reports").ListObjects( _
        "MAD_Combined_Reports").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Combined MAD Reports").ListObjects( _
        "MAD_Combined_Reports").Sort.SortFields.Add2 Key:=Range( _
        "MAD_Combined_Reports[[#All],[Item]]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Combined MAD Reports").ListObjects( _
        "MAD_Combined_Reports").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("A3").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Select
    Selection.End(xlUp).Select
   ActiveCell.Offset(1).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Select
    Selection.EntireRow.Delete
    
    ActiveWorkbook.Connections("Query - Transaction Register Reports Combined"). _
        Refresh
    DoEvents
    ActiveWorkbook.Connections("Query - Production Sheet Dump").Refresh
    DoEvents
    ActiveWorkbook.Connections("Query - TR WIP ASS SCRAP").Refresh
    DoEvents
    ActiveWorkbook.Connections("Query - MAD 62090").Refresh
    DoEvents
    ActiveWorkbook.Connections("Query - MAD 59170 INV Scrap").Refresh
    DoEvents
    ActiveWorkbook.Connections("Query - Scrap Sheet Dump").Refresh
    DoEvents

        
    Sheets("Sheet1").Select
    ActiveSheet.PivotTables("PivotTable4").PivotCache.Refresh
    ActiveSheet.PivotTables("PivotTable5").PivotCache.Refresh
    
    ActiveWorkbook.Save
    
    ActiveWorkbook.Connections("Query - The Best").Refresh
    DoEvents
    
    
    'ActiveSheet.PivotTables("PivotTable5").PivotFields("Week of Year").PivotFilters _
        .Add2 Type:=xlValueIsGreaterThan, DataField:=ActiveSheet.PivotTables( _
        "PivotTable5").PivotFields("Sum of Txn Value"), Value1:=0

    End Sub
    
 Sub FillInScrap()
 
     Windows("New Scrap Template AK.xlsx").Activate
   
    Sheets("Sheet1").Select
    Range("F5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Windows("KPI - Scrap Trend Report - June 2026.xlsx").Activate
    ActiveSheet.Paste
    Windows("New Scrap Template AK.xlsx").Activate
    Range("G5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("KPI - Scrap Trend Report - June 2026.xlsx").Activate
    ActiveCell.Offset(0, 1).Select
    ActiveSheet.Paste
    Windows("New Scrap Template AK.xlsx").Activate
    Range("H5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("KPI - Scrap Trend Report - June 2026.xlsx").Activate
    ActiveCell.Offset(0, 9).Select
    ActiveSheet.Paste
    
    

    
    Windows("New Scrap Template AK.xlsx").Activate
    Range("I5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("KPI - Scrap Trend Report - June 2026.xlsx").Activate
    ActiveCell.Offset(0, -1).Select
    ActiveSheet.Paste
    
    Windows("New Scrap Template AK.xlsx").Activate
    Range("k5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select    'HERE HERe
    Application.CutCopyMode = False
    Selection.Copy
    Windows("KPI - Scrap Trend Report - June 2026.xlsx").Activate
    ActiveCell.Offset(0, 3).Select
    ActiveSheet.Paste
    
    
    Windows("New Scrap Template AK.xlsx").Activate
    Range("M5:Q5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("KPI - Scrap Trend Report - June 2026.xlsx").Activate
    ActiveCell.Offset(0, 1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    
    
    
    'Filling in the formulas
    
    

    ActiveCell.Offset(0, 5).Select
    ActiveCell.Offset(-1).Select
    Range(Selection, Selection.Offset(0, 2)).Select
    
    Application.CutCopyMode = False
    Selection.Copy
    ActiveCell.Offset(0, -6).Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 6).Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    
    ActiveCell.Offset(0, -16).Select
   
    Range(Selection, Selection.Offset(0, 6)).Select
    Application.CutCopyMode = False
    Selection.Copy
    ActiveCell.Offset(0, -1).Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    
    
    ActiveCell.Offset(0, -5).Select
   ActiveCell.Offset(1).Select
   
   Windows("New Scrap Template AK.xlsx").Activate
    Range("J5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("KPI - Scrap Trend Report - June 2026.xlsx").Activate
    
    ActiveSheet.Paste
   
    ActiveCell.Offset(0, 1).Select
   ActiveCell.Offset(-1).Select
   
    Range(Selection, Selection.Offset(0, 1)).Select
    Application.CutCopyMode = False
    Selection.Copy
    ActiveCell.Offset(0, 2).Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, -2).Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    
    
    
End Sub


Sub FillScrapProd()

 Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim maxDate As Date
    
    ' Open the workbook
    'Set wb = Workbooks.Open("H:\KPI SLC\2026\06-Jun\KPI - Scrap Trend Report - June 2026.xlsx")
    Set wb = Workbooks.Open("C:\Users\jh00082083\Documents\aaWeekly Reports\KPI - Scrap Trend Report - June 2026.xlsx")
    Set ws = wb.Sheets("Prod")
    
    ' Find the last used row in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    ws.Activate
    ' Find the maximum date in column A
    
    maxDate = Application.WorksheetFunction.Max(ws.Range("A2:A" & lastRow))
    
    ' Clear any existing filters
    ws.AutoFilterMode = False
    
    ' Select visible cells in columns B to J (excluding header)
    On Error Resume Next
    ws.Range("B2:B" & lastRow).SpecialCells(xlCellTypeVisible).Select
    On Error GoTo 0

    ' Apply the filter to show only rows with the maximum date
    ws.Range("$A$1:$J$" & lastRow).AutoFilter Field:=1, Criteria1:=maxDate
 ' Select visible cells in columns B to J (excluding header)
    On Error Resume Next
    ws.Range("B2:M" & lastRow).SpecialCells(xlCellTypeVisible).Select
    On Error GoTo 0
    
    Selection.ClearContents
    
    Rows("1:1").Select
 ActiveSheet.ShowAllData
    Range("B2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1).Select
      
Windows("MACRO.xlsm").Activate
Worksheets("Prod").Activate
Range("A2:L2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Windows("KPI - Scrap Trend Report - June 2026.xlsx").Activate
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False




Selection.End(xlDown).Select
ActiveCell.Offset(0, -1).Select
Selection.End(xlUp).Select



Selection.Copy
ActiveCell.Offset(0, 1).Select
Selection.End(xlDown).Select
ActiveCell.Offset(0, -1).Select
Range(Selection, Selection.End(xlUp)).Select
ActiveSheet.Paste

    ' Apply the filter to show only rows with the maximum date
 ws.Range("$A$1:$J$" & lastRow).AutoFilter Field:=1, Criteria1:=maxDate

    Sheets("ByAssembly").Select
    ActiveSheet.PivotTables("PivotTable2").PivotCache.Refresh
    ActiveSheet.PivotTables("PivotTable1").PivotCache.Refresh
    ActiveSheet.PivotTables("PivotTable5").PivotCache.Refresh
    ActiveSheet.PivotTables("PivotTable6").PivotCache.Refresh
    
    Sheets("By Machine & NCMR").Select
    ActiveSheet.PivotTables("PivotTable3").PivotCache.Refresh
    ActiveSheet.PivotTables("PivotTable4").PivotCache.Refresh
    Sheets("PLANT SCRAP DEPARTMENT-NEW DEPT").Select
    
    
    
    
    
    
    
    
Workbooks.Open "R:\SLC Finance\Month End\KPI Report\Templates AK\New Production Template AK.xlsx"

Sheets("MUV & Scrap Feeder Pivot").Select
ActiveSheet.Range("J4:J18").Select
 Selection.Copy

'Workbooks.Open ("H:\KPI SLC\2026\06-Jun\KPI - Scrap Trend Report - June 2026.xlsx")
Workbooks.Open ("C:\Users\jh00082083\Documents\aaWeekly Reports\KPI - Scrap Trend Report - June 2026.xlsx")
   
   
Worksheets("PLANT SCRAP DEPARTMENT-NEW DEPT").Activate
Range("Y25").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
 Worksheets("USE THIS").Activate
    ActiveSheet.PivotTables("PivotTable2").PivotCache.Refresh
  
    ActiveSheet.PivotTables("PivotTable1").PivotCache.Refresh
Sheets("PLANT SCRAP DEPARTMENT-NEW DEPT").Select
End Sub
