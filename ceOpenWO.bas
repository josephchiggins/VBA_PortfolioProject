Attribute VB_Name = "ceOpenWO"
Sub ZZWO()
'
' ZZWO Macro
'
Workbooks.Open ("H:\KPI SLC\2026\06-Jun\KPI - Open Work Order Report - June 2026.xlsx")


    ActiveWorkbook.Connections("Query - Discrete Job Listing Reports Combined"). _
        Refresh

    Sheets("Automation Work Orders").Select
    ActiveSheet.PivotTables("PivotTable1").PivotCache.Refresh
    ActiveSheet.PivotTables("PivotTable2").PivotCache.Refresh
    
    Sheets("Molding Work Orders").Select
    ActiveSheet.PivotTables("PivotTable1").PivotCache.Refresh
    ActiveSheet.PivotTables("PivotTable2").PivotCache.Refresh
    
    Sheets("Catheter Work Orders").Select
    ActiveSheet.PivotTables("PivotTable1").PivotCache.Refresh
    ActiveSheet.PivotTables("PivotTable2").PivotCache.Refresh
    
    ActiveWorkbook.Save
    
    Call ZWoEmail
    
    Workbooks.Open ("H:\KPI SLC\2026\06-Jun\KPI - Open Work Order Report - June 2026.xlsx")
    Workbooks("KPI - Open Work Order Report - June 2026.xlsx").Close
End Sub


Sub ZWoEmail()

    ' Define the previous month and date format
    Dim newDate As String
    newDate = Format(DateAdd("d", -1, Now), "mm\dd\yyyy")
    
    Dim newMonth As String
    newMonth = Format(DateAdd("M", -1, Now), "MMMM")

    ' Open the workbook and activate the worksheet
    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = Workbooks.Open("H:\KPI SLC\2026\06-Jun\KPI - Open Work Order Report - June 2026.xlsx")
   ' Set ws = wb.Sheets("Summary")
    'ws.Activate
    
    ' Open Outlook
    Dim oOutlook As Object
    Set oOutlook = CreateObject("Outlook.Application")
    
    ' Open Email
    Dim oEmail As Object
    Set oEmail = oOutlook.CreateItem(0) ' 0 = olMailItem
    
    With oEmail
        .To = "madan.pitchamuthu@icumed.com;Judith.McDonald@icumed.com;Keith.Martin@icumed.com;Robert.Hansen@icumed.com;Micki.Ellefsen@icumed.com;Noureddine.Atif@icumed.com;ramon.camac@icumed.com;Tien.Dieu@icumed.com;Tracy.Parry@icumed.com;timothy.hansen@icumed.com;Aaron.Hernandez2@icumed.com;newton.santana@icumed.com;breanna.padilla@icumed.com;adam.hancock@icumed.com;Karina.Crittenden@icumed.com;Joni.Giles@icumed.com;andreas.outlar@icumed.com;adam.kesling@icumed.com;chris.luebke@icumed.com;shaun.mason@icumed.com;michael.stuart@icumed.com"
        .CC = ""
        .BCC = ""
        .Subject = "SLC Open Work Order Report - June 2026"

        ' Set up the body text with HTML formatting
        .HTMLBody = "<BODY style='font-size:11pt; font-family:Calibri;'>" & _
                    "All,<br><br>" & _
                    "Attached is the weekly SLC Open Work Order Report, updated through " & newDate & "<br><br>" & _
                    "Thank you,<br><br>" & _
                    "" & _
                    "Joe" & _
                    .HTMLBody ' Add this before any other content

                
         .Attachments.Add "H:\KPI SLC\2026\06-Jun\KPI - Open Work Order Report - June 2026.xlsx"


        .Display

 End With
    
End Sub

