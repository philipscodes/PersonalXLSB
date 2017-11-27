Attribute VB_Name = "Sales_Report_POC"

Public varPrepareUCStock As Boolean
Public varPrepareUCPurchase As Boolean
Public varPrepareHiyaza As Boolean
Public varPrepareUCSales As Boolean
Public varPrepareUCSalesYTD As Boolean


Sub sendUCReportEmail()

'*******************************************************************************************************
'This procedure prepares the sales management report for POC LandRover cars sold for the current month.*
'*******************************************************************************************************

'Check to make sure that the code runs only when the specified file is open.
If fileCheck("UC JLR") = False Then End

'On Error Resume Next

Dim tempWorkBook As Workbook
Dim mainWorkBook As Workbook
Dim vTo As String
Dim vSub As String
Dim vMsg As String

'Temporarily Stop Events, Calculations, Screen & Statusbar updations.
Call excelVitalsFalse

Set mainWorkBook = ActiveWorkbook   'Referencing current open original report file without reference to workbook name.
Set tempWorkBook = Workbooks.Add    'Creating and referencing a new workbook to hold temp sheets from original file.

'Move/Copy sheets 'ATM Dashboard', 'Sales Adv', 'Target' to new workbook.
mainWorkBook.Sheets(Array("POC DASHBOARD", "Sales Advisor")).Copy Before:=tempWorkBook.Sheets(1)

'Call Paste value routine against each sheet.
tempWorkBook.Sheets("POC DASHBOARD").Activate
tempWorkBook.Sheets("POC DASHBOARD").Cells.Select
Call pasteValues

tempWorkBook.Sheets("Sales Advisor").Activate
tempWorkBook.Sheets("Sales Advisor").Cells.Select
Call pasteValues

'Delete sheet1
tempWorkBook.Sheets("Sheet1").Delete

'Delete unwanted data

tempWorkBook.Sheets("Sales Advisor").Activate
tempWorkBook.Sheets("Sales Advisor").Range("G:Z").Clear
tempWorkBook.Sheets("Sales Advisor").Range("A1").Select

tempWorkBook.Sheets("POC DASHBOARD").Activate
tempWorkBook.Sheets("POC DASHBOARD").Range("Y:AB").Clear
tempWorkBook.Sheets("POC DASHBOARD").Range("B1").Select

'Save temporary workbook to root directory
tempWorkBook.SaveAs Filename:="C:\Temp\tempNCReport", FileFormat:=xlWorkbookNormal

'Preparing email items
vTo = "tmerheb@altayer.com;hhanei@altayer.com;jelias@altayer.com;hhosny@altayer.com"
vSub = "JLR POC Sales & Stock Report - " & MonthName(Month(Now())) & " " & Year(Now())
vMsg = "Good Morning," & vbCrLf & vbCrLf & _
"JLR POC Sales & Stock Report as of now attached." & vbCrLf & vbCrLf & _
"Best Regards," & vbCrLf & _
"Philip"


'Send new workbook as attachment.
Call sendEmail(vTo, vSub, vMsg, tempWorkBook)


'Close & Delete temporary file
tempWorkBook.Close
Kill "C:\temp\tempNCReport.xls"

'Close original file with save.
mainWorkBook.Close SaveChanges:=True

Call excelVitalsTrue

'Close Excel Application if there are no open workbooks.
Call closeWorkbook

End Sub

Public Sub prepareUCReports()

    Load usrFrmUCReports
    usrFrmUCReports.Show
    
'Check to make sure that the code runs only when the specified file is open.
If fileCheck("UC JLR") = False Then End

'Temporarily Stop Events, Calculations, Screen & Statusbar updations.
Call excelVitals(False)
    
    If varPrepareUCStock = True Then Call prepareUCStock
    If varPrepareUCPurchase = True Then Call prepareUCPurchase
    If varPrepareHiyaza = True Then Call prepareHyazaRelease
    'If varPrepareUCSalesYTD = True Then Call prepareUCSalesYTD
    If varPrepareUCSales = True Then
        Call prepareUCSales
        Call refreshPOCSalesPivots
    End If
    
    
Call excelVitals(True)
    
End Sub

Private Sub prepareUCStock()
Attribute prepareUCStock.VB_ProcData.VB_Invoke_Func = " \n14"
'
' UC_Stock_Clear Macro
'

'

On Error Resume Next

    Sheets("Stock Spool").Activate

    ActiveSheet.ShowAllData
    Call applySort("I", "U")
    Call findAndDeleteRange("Stock Spool", "Franchise:", 100, 18)

End Sub
Private Sub prepareUCPurchase()
Attribute prepareUCPurchase.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Purchase Macro
'

'

Sheets("Purchase Spool").Activate


Dim LR As Long
Dim cnt As Integer
Dim tempCnt As Integer
Dim arrPurchaseSpoolFiltered() As Variant

    On Error Resume Next
    ActiveSheet.ShowAllData
    On Error GoTo 0
    
    Call applySort("G", "O")
    
    
    'Add formulas to sheet 'Purchase Spool'
    Call addFormulas("O2", "=VLOOKUP(RC[-8],'Purchase Details'!C[-8],1,FALSE)", True, , , , , , "RC")
    
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    cnt = Selection.count
    cnt = cnt + 1
   
    LR = Range("G" & Rows.count).End(xlUp).Row
    
On Error GoTo formulas

    With Sheets("Purchase Spool")
        .AutoFilterMode = False
        With .Range("A1:O1000")
            .AutoFilter Field:=15, Criteria1:="#N/A"
            Range("A2:N" & LR).SpecialCells(xlCellTypeVisible).Copy Destination:=Sheets("Purchase Details").Range("A1").End(xlDown).Offset(1, 0)
        End With
    End With
    
formulas:
    Call addFormulas("O2", "=IF(LEFT(RC[-8],3)=""SAL"",""LR"",IF(LEFT(RC[-8],3)=""SAJ"",""JAG"",IF(LEFT(RC[-8],3)=""SAD"",""Jag"",""ZNF"")))", True, "Purchase Details", , , , , "RC")
    Call addFormulas("P2", "=VLOOKUP(RC[-9],'Purchase Spool'!C[-9],1,FALSE)", True, "Purchase Details", , , , , "RC")
    
    Range("A2").Select
    
    
End Sub
Private Sub prepareHyazaRelease()
Attribute prepareHyazaRelease.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Hyaza_Release Macro
'

'

On Error Resume Next

    Sheets("Hiyaza Spool").Activate

    ActiveSheet.ShowAllData
    Call applySort("J", "L")    'Sort by 'released' date - J.
    Call findAndDeleteRange("Hiyaza Spool", "Franchise:", 100, 18)
    Call addFormulas("C2", "=IFERROR(LEFT(RC[1],FIND("" "",RC[1])-1),RC[1])", True, "Hiyaza Spool", , , , , "RC")
    Call addFormulas("H2", "=VLOOKUP(RC[3],Lookup_Tables!C[-1]:C,2,FALSE)", False, "Hiyaza Spool", , , , , "RC")
    
    
    'Change Header
    Sheets("Hiyaza Spool").Range("C1") = "Short Model"
    
End Sub

Public Sub prepareUCSalesYTD()


On Error Resume Next

'Check to make sure that the code runs only when the specified file is open.
If fileCheck("Master Data - POC JLR") = False Then End

'Temporarily Stop Events, Calculations, Screen & Statusbar updations.
Call excelVitals(False)

Sheets("Sales Spool (YTD)").Activate
   
Call pocSalesSpoolYTDFormulas

Call excelVitals(True)

End Sub

Private Sub prepareUCSales()
Attribute prepareUCSales.VB_ProcData.VB_Invoke_Func = " \n14"
'
' POC_Sales Macro
'

'


On Error Resume Next

Dim LR As Long
Dim cnt As Integer
Dim tempCnt As Integer

    Sheets("Sales Spool").Activate
    
    Call pocSalesSpoolFormulas

    
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    cnt = Selection.count


    ActiveSheet.Range("$A$1:$T$65525").AutoFilter Field:=20, Criteria1:="#N/A"
    Range("A1").Select
    
    LR = Range("A" & Rows.count).End(xlUp).Row
    Range("A2:T" & LR).SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    
    Sheets("Sales Details").Activate
    
    Range("A1").End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
    
    Range("A2").Select
    Selection.End(xlDown).Select
    
'To remove 'Grand Total Line'

    Do While UCase(Left(Trim(ActiveCell.Value), 3)) = "GRA"
        ActiveCell.EntireRow.Select
        Selection.Delete
        Range("A2").Select
        Selection.End(xlDown).Select
    Loop

    
'To add ATM & PM company info
    tempCnt = 2
    Do While cnt >= tempCnt
        
        If Range("A" & tempCnt) = "r31" Then
            Range("T" & tempCnt) = "PM"
        ElseIf Range("A" & tempCnt) = "r32" Then
            Range("T" & tempCnt) = "PM"
        ElseIf Range("A" & tempCnt) = "r34" Then
            Range("T" & tempCnt) = "PM"
        Else
            Range("T" & tempCnt) = "ATM"
        End If
            
        tempCnt = tempCnt + 1
    Loop

'To add Franchisee info
    tempCnt = 2
    Do While cnt >= tempCnt
        
        If Left(Range("H" & tempCnt), 3) = "SAL" Then
            Range("U" & tempCnt) = "LR"
        ElseIf Left(Range("H" & tempCnt), 3) = "SAJ" Then
            Range("U" & tempCnt) = "JAG"
        ElseIf Left(Range("H" & tempCnt), 3) = "SAD" Then
            Range("U" & tempCnt) = "JAG"
        Else
            Range("U" & tempCnt) = "ZNF"
        End If
            
        tempCnt = tempCnt + 1
    Loop

'To remove 'Inter-co' info
    tempCnt = 2
    Do While cnt >= tempCnt
        
        If Left(Range("J" & tempCnt), 5) = "Inter" Then
            Range("L" & tempCnt).Clear
            Range("M" & tempCnt).Clear
            Range("T" & tempCnt).Clear
            Range("U" & tempCnt).Clear
        End If
            
        tempCnt = tempCnt + 1
    Loop
   
    Call pocSalesDetailedFormulas

End Sub
Private Sub pocSalesSpoolYTDFormulas()

    Call addFormulas("P2", "=IF(OR(LEFT(RC[-8],3)=""SAJ"",LEFT(RC[-8],3)=""SAD""),""JAG"",IF(LEFT(RC[-8],3)=""SAL"",""LR"",""ZNF""))", False, "Sales Spool (YTD)", , , , , "RC")
    Call addFormulas("Q2", "=TEXT(RC[-14],""MMM-YYYY"")", False, "Sales Spool (YTD)")
    Call addFormulas("R2", "=IF(OR(RC[-17]=""r31"",RC[-17]=""r32"",RC[-17]=""r34""),""PM"",""ATM"")", False, "Sales Spool (YTD)", , , , , "RC")

'Change Header

    Sheets("Sales Spool (YTD)").Range("P1") = "Franchise"
    Sheets("Sales Spool (YTD)").Range("P1") = "Invoice Month-Year"
    Sheets("Sales Spool (YTD)").Range("P1") = "Company"
    
End Sub
Private Sub pocSalesSpoolFormulas()

    Call addFormulas("T2", "=VLOOKUP(RC[-12],'Sales Details'!C[-12],1,FALSE)", False, , , , , , "RC")

'Change Header

    Sheets("Sales Spool").Range("T1") = "Sales Details"
    
End Sub

Private Sub pocSalesDetailedFormulas()

    Call addFormulas("M2", "=VLOOKUP(RC[-5],'Sales Spool'!C[-5]:C,6,FALSE)", False, , , , , , "RC")
    Call addFormulas("V2", "=VLOOKUP(RC[-14],'Sales Spool'!C[-14],1,FALSE)", False, , , , , , "RC")
    Call addFormulas("W2", "=IF(RIGHT(RC[-16],3)=""(A)"",""A"",""N"")", False, , , , , , "RC")

'Change Header

    Sheets("Sales Spool").Range("M1") = "Sales Type"
    Sheets("Sales Spool").Range("V1") = "Sales Spool"
    Sheets("Sales Spool").Range("W1") = "Approved Vehicles"
    
End Sub

Private Sub refreshPOCSalesPivots()


    Sheets("Sales Advisor").PivotTables("pvtPOCSales").RefreshTable
    Sheets("Sales Advisor").PivotTables("pvtHiyaza").RefreshTable


End Sub


Sub clearUCSalesSpool()

    Call clearSheetData("UC JLR", "Sales Spool", "A", "A2", "S")

End Sub

Sub clearHiyaza()

    Call clearSheetData("UC JLR", "Hiyaza Spool", "A", "A2", "K")

End Sub

Sub clearUCStockSpool()

    Call clearSheetData("UC JLR", "Stock Spool", "A", "A2", "U")

End Sub

Sub clearPurchase()

    Call clearSheetData("UC JLR", "Purchase Spool", "A", "A2", "N")

End Sub

