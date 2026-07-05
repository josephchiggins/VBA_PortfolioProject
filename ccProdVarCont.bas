Attribute VB_Name = "ccProdVarCont"
Sub zzProdVarContr()
'


    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim maxDate As Date

    

'Open and Update template

Workbooks.Open "R:\SLC Finance\Month End\KPI Report\Templates AK\New Production Variance Template AK.xlsx"
Worksheets("Production Variance Dump").Activate

    ActiveWorkbook.Connections("Query - ICU MUV Scrap").Refresh
    DoEvents
    ActiveWorkbook.Connections("Query - Current Weekly Reports").Refresh
    DoEvents
    
    
    
    
    
    
    ' Open the workbook
Set wb = Workbooks.Open("H:\KPI SLC\2026\06-Jun\KPI - Production Variances & Contribution Margin - June 2026.xlsx")
   
Set ws = wb.Sheets("Full Data")
    

Worksheets("Full Data").Activate
        
    
    
 
' Find the last used row in column A
    lastRow = ws.Cells(ws.Rows.Count, "X").End(xlUp).Row
    
    ' Find the maximum date in column A
    maxDate = Application.WorksheetFunction.Max(ws.Range("X2:X" & lastRow))
    

ws.Range("$A$1:$AD$" & lastRow).AutoFilter Field:=24, Criteria1:=2026

' Apply the AutoFilter to column W
ws.Range("$A$1:$AD$" & lastRow).AutoFilter Field:=23, Criteria1:="Jun"

    
    
    
' Select visible cells in columns B to J (excluding header)
    On Error Resume Next
    ws.Range("A2:AD" & lastRow).SpecialCells(xlCellTypeVisible).Select
    On Error GoTo 0
    
 
    
    
    
    Selection.ClearContents
    
     Rows("1:1").Select
     ActiveSheet.ShowAllData
    Range("A2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1).Select


Windows("New Production Variance Template AK").Activate
Worksheets("Production Variance Dump").Activate



Worksheets("Production Variance Dump").Range("A3:AD3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    
    
Windows("KPI - Production Variances & Contribution Margin - June 2026").Activate
Worksheets("Full Data").Activate
  
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False


ActiveWorkbook.Save
    
    Call ProdVarianceEmail
    
    Workbooks.Open ("H:\KPI SLC\2026\06-Jun\KPI - Production Variances & Contribution Margin - June 2026.xlsx")
    Workbooks("KPI - Production Variances & Contribution Margin - June 2026.xlsx").Close

End Sub


Sub ProdVarianceEmail()

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim filePath As String

    ' === Dynamic dates ===
    Dim newDate As String
    newDate = Format(DateAdd("d", -1, Date), "mm/dd/yyyy")

    Dim newMonth As String
    newMonth = Format(DateAdd("d", -1, Date), "MMMM")

    ' === File path ===
    filePath = "H:\KPI SLC\2026\06-Jun\KPI - Production Variances & Contribution Margin - June 2026.xlsx"

    ' === Open workbook (leave open) ===
    Set wb = Workbooks.Open(filePath)
    Set ws = wb.Sheets("Production Variance - Summary")
    ws.Activate

    ' === Open Outlook ===
    Dim oOutlook As Object
    Dim oEmail As Object
    Set oOutlook = CreateObject("Outlook.Application")
    Set oEmail = oOutlook.CreateItem(0)

    With oEmail
        ' === Recipients ===
        .To = _
            "timothy.hansen@icumed.com; madan.pitchamuthu@icumed.com; " & _
            "edward.paulsen@icumed.com; " & _
            "breanna.padilla@icumed.com; " & _
            "adam.hancock@icumed.com; kyle.brady@icumed.com; " & _
            "cj.sudbury@icumed.com; joseph.gutierrez1@icumed.com; " & _
            "ryan.fallon@icumed.com; " & _
            "andreas.outlar@icumed.com; " & _
            "travis.hathaway@icumed.com"

        ' === Subject ===
        .Subject = "SLC Production Variance - " & newMonth

        ' === Body text FIRST ===
        .HTMLBody = "<body style='font-family:Calibri;font-size:11pt;'>" & _
                    "All,<br><br>" & _
                    "Attached is the weekly Production Variance through " & newDate & ".<br><br>" & _
                    "Thanks,<br>" & _
                    "Joe<br><br>" & _
                    "<b><u>Production Variance - General:<\u><\b><br>" & _
                    "<\body>"

        ' === Attach file ===
        .Attachments.Add filePath

        .Display
        DoEvents
        Application.Wait Now + TimeValue("00:00:02")  ' <-- Added wait for Outlook to load
        .GetInspector.WindowState = 1

        ' === Copy screenshot as Bitmap ===
        ws.Range("B3:H46").CopyPicture Appearance:=xlScreen, Format:=xlBitmap  ' <-- xlBitmap

        ' === Paste AFTER text ===
        Dim oInspector As Object      ' <-- Change from Outlook.Inspector
        Dim oWordDoc As Object        ' <-- Change from Word.Document
        Dim oWordRng As Object

        Set oInspector = .GetInspector
        Set oWordDoc = oInspector.WordEditor
        Set oWordRng = oWordDoc.Application.ActiveDocument.Content

        oWordRng.Collapse Direction:=wdCollapseEnd
        oWordRng.InsertAfter " " & vbNewLine
        oWordRng.Collapse Direction:=wdCollapseEnd
        oWordRng.InsertAfter String(1, vbNewLine)      ' <-- Space before image
        oWordRng.Collapse Direction:=wdCollapseEnd
        oWordRng.PasteSpecial DataType:=4

        ' === Resize image ===
        If oWordDoc.InlineShapes.Count > 0 Then
            With oWordDoc.InlineShapes(oWordDoc.InlineShapes.Count)
                .LockAspectRatio = msoTrue  ' <-- Lock aspect ratio
                .Width = 500                 ' <-- Width only, height auto-adjusts
            End With
        End If

    End With

    ' === Cleanup (do NOT close workbook) ===
    Set oEmail = Nothing
    Set oOutlook = Nothing

End Sub

