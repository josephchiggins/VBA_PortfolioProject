Attribute VB_Name = "aNewMUV"
Sub NewMUVTest()
Attribute NewMUVTest.VB_ProcData.VB_Invoke_Func = " \n14"
'
' NewMUVTest MACRO
'
    Range("B1").Select
    Selection.EntireColumn.Insert
    Selection.EntireColumn.Insert
   ' Range("A2").Select
    'Range(Selection, Selection.End(xlDown)).Select
    'Selection.TextToColumns Destination:=Range("A2"), DataType:=xlFixedWidth, _
     '   FieldInfo:=Array(Array(0, 1), Array(9, 1), Array(19, 1)), TrailingMinusNumbers:= _
      '  True
        
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.TextToColumns Destination:=Range("A2"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(10, 1), Array(20, 1)), TrailingMinusNumbers _
        :=True
        
        
    Columns("B:D").Select
    Selection.Delete Shift:=xlToLeft
    Columns("C:S").Select
    Selection.Delete Shift:=xlToLeft
    Columns("D:D").Select
    Selection.Delete Shift:=xlToLeft
    Columns("E:H").Select
    Selection.Delete Shift:=xlToLeft
    
Call PutIntoMUVReport
Call baPoductionTemplate

'''''

 Dim userResponse As VbMsgBoxResult
    
    userResponse = MsgBox("Do you want to continue to MUV by Component?", vbYesNo + vbQuestion, "Confirm Action")
    
    If userResponse = vbNo Then
        MsgBox "Operation canceled.", vbInformation
        Exit Sub
    End If

    ' Continue running code here
    MsgBox "Continuing with the code...", vbInformation
    
''''''


Call UpdateMUV_By_ComponentTemplate




Workbooks.Open ("C:\Users\jh00082083\Documents\aaWeekly Reports\KPI - MUV - simplified - June 2026.xlsx")

Worksheets("MUVCurrMonth").Activate



End Sub

Sub PutIntoMUVReport()
'
' PutIntoMUVReport MACRO
'
Dim ws As Worksheet
Dim lastRow As Long
Dim rng As Range
Dim newestDate As Date
Dim rngDate As Range
Dim rngSort As Range
Dim maxDate As Date


    ' Open the workbook
    'Set wb = Workbooks.Open("H:\KPI SLC\2026\06-Jun\KPI - MUV - simplified - June 2026.xlsx")
    Set wb = Workbooks.Open("C:\Users\jh00082083\Documents\aaWeekly Reports\KPI - MUV - simplified - June 2026.xlsx")
    
    
     Set ws = wb.Sheets("MUVCurrMonth")
    
     ws.Activate
    ' Clear any existing filters
    ws.AutoFilterMode = False
      
      
      
    Worksheets("MUVCurrMonth").Activate
 
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
    ActiveCell.Offset(0, 1).Select
   
    Windows("MACRO.xlsm").Activate
    Range("C2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Windows("KPI - MUV - simplified - June 2026.xlsx").Activate
    ActiveSheet.Paste
    Windows("MACRO.xlsm").Activate
    Range("D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("KPI - MUV - simplified - June 2026.xlsx").Activate
    ActiveCell.Offset(0, 1).Select
    ActiveSheet.Paste
    ActiveCell.Offset(0, 3).Select
    Windows("MACRO.xlsm").Activate
    Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("KPI - MUV - simplified - June 2026.xlsx").Activate
    ActiveSheet.Paste
    Windows("MACRO.xlsm").Activate
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("KPI - MUV - simplified - June 2026.xlsx").Activate
    ActiveCell.Offset(0, 4).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveCell.Offset(0, -3).Select
    ActiveCell.Offset(-1).Select
    Range(Selection, Selection.Offset(0, 2)).Select
    
    Application.CutCopyMode = False
    Selection.Copy
    ActiveCell.Offset(0, -1).Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    
    ActiveCell.Offset(0, -3).Select
    ActiveCell.Offset(-1).Select
    Range(Selection, Selection.Offset(0, 1)).Select
    Application.CutCopyMode = False
    Selection.Copy
    ActiveCell.Offset(0, -1).Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    
    
    Range("B2").Select
    Selection.End(xlDown).Select
ActiveCell.Offset(0, -1).Select
Selection.End(xlUp).Select



Selection.Copy
ActiveCell.Offset(0, 1).Select
Selection.End(xlDown).Select
ActiveCell.Offset(0, -1).Select
Range(Selection, Selection.End(xlUp)).Select
ActiveSheet.Paste


' Find the last used row in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Find the maximum date in column A
    maxDate = Application.WorksheetFunction.Max(ws.Range("A2:A" & lastRow))
    

ws.Range("$A$1:$X$" & lastRow).AutoFilter Field:=1, Criteria1:=maxDate
  
--------------

Sub acFillMUVrepMonthRecurr()
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim maxDate As Date
    Dim maxDateJ As Date
    
    ' Open the workbook
    'Set wb = Workbooks.Open("H:\KPI SLC\2026\06-Jun\KPI - MUV - simplified - June 2026.xlsx")
    Set wb = Workbooks.Open("C:\Users\jh00082083\Documents\aaWeekly Reports\KPI - MUV - simplified - June 2026.xlsx")
    Set ws = wb.Sheets("MUVCurrMonth")
    
    ' Find the last used row in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Find the maximum date in column A
    maxDate = Application.WorksheetFunction.Max(ws.Range("A2:A" & lastRow))
    
    ' Find the maximum date in column J
    maxDateJ = Application.WorksheetFunction.Max(ws.Range("J2:J" & lastRow))
    
    ws.Activate
    ' Clear any existing filters
    ws.AutoFilterMode = False
    
    ' Apply the filter to show only rows with the maximum date in  column A
    ws.Range("$A$1:$J$" & lastRow).AutoFilter Field:=1, Criteria1:=maxDate
  
 ' Select visible cells in columns B to J (excluding header)
    On Error Resume Next
    ws.Range("B2:L" & lastRow).SpecialCells(xlCellTypeVisible).Select
    On Error GoTo 0
    
    Selection.ClearContents
    
    Rows("1:1").Select
 ActiveSheet.ShowAllData
    Range("B2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1).Select
      
Windows("MACRO.xlsm").Activate
Worksheets("MonthRecurr").Activate
Range("A2:K2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Windows("KPI - MUV - simplified - June 2026.xlsx").Activate
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




ws.Range("$A$1:$J$" & lastRow).AutoFilter Field:=1, Criteria1:=maxDate
ws.Range("$A$1:$J$" & lastRow).AutoFilter Field:=10, Criteria1:=maxDateJ
ActiveWorkbook.Worksheets("MUVCurrMonth").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("MUVCurrMonth").AutoFilter.Sort.SortFields.Add2 Key _
        :=Range("F1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("MUVCurrMonth").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
        
    End With


 
    

End Sub



Sub bcFillMUVProd()

 Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim maxDate As Date
    
    ' Open the workbook
    'Set wb = Workbooks.Open("H:\KPI SLC\2026\06-Jun\KPI - MUV - simplified - June 2026.xlsx")
    Set wb = Workbooks.Open("C:\Users\jh00082083\Documents\aaWeekly Reports\KPI - MUV - simplified - June 2026.xlsx")
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
    ws.Range("B2:K" & lastRow).SpecialCells(xlCellTypeVisible).Select
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
    Windows("KPI - MUV - simplified - June 2026.xlsx").Activate
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
    ActiveSheet.PivotTables("PivotTable6").PivotCache.Refresh
    ActiveSheet.PivotTables("PivotTable4").PivotCache.Refresh
    Sheets("PLANT MUV DEPARTMENT-NEW DEPT").Select
    ActiveSheet.PivotTables("PivotTable2").PivotCache.Refresh
    ActiveSheet.PivotTables("PivotTable3").PivotCache.Refresh
    
    Sheets("MUVCurrMonth").Select

End Sub

 Sub CaMUVEmail()

Dim OutApp As Object
Dim OutMail As Object
Dim table As Range
Dim pic As Picture
Dim wb As Workbook
Dim ws As Worksheet
Dim wordDoc


Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)

'grab table, convert to image, and cut
Set wb = Workbooks.Open("H:\KPI SLC\2026\06-Jun\KPI - MUV - simplified - June 2026.xlsx")
Set ws = wb.Sheets("UseThis")
ws.Activate
Set table = ws.Range("C4:N34")
ws.Activate
table.Copy
Set pic = ws.Pictures.Paste
pic.Cut

'create email message
On Error Resume Next
    With OutMail
        .To = "team@abc.com"
        .CC = ""
        .BCC = ""
        .Subject = "Country Population Data " & Format(Date, "mm-dd-yy")
        .Display
        
        Set wordDoc = OutMail.GetInspector.WordEditor
            With wordDoc.Range
                .PasteandFormat wdChartPicture
                .insertParagraphAfter
                .insertParagraphAfter
                .InsertAfter "Thank you,"
                .insertParagraphAfter
                .InsertAfter "Joe"
            End With
            
        .HTMLBody = "<BODY style = font-size:11pt; font-family:Calibri >" & _
            "Hi Team, <p> Please see table below: <p>" & .HTMLBody
    End With
    On Error GoTo 0
    
Set OutApp = Nothing
Set OutMail = Nothing

End Sub
    
