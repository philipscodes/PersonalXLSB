Attribute VB_Name = "Reports"
Public Sub InvoicesVsDelivered()

'*********************************************************************************************************************
'This report is prepared on a weekly basis every Saturday.
'It is saved in the shared folder network location:- \\ho-webstore\ATM\LR and UC Sales Report\LR\General Reports\
'Report Requested By Mr. Rajendran Ammal - JLR Administration Manager on 13th February, 2016
'Report Excel filename - 'Invoice vs Reported Vehicles.xlsb'
'This report serves the following requirements of higher management.
'---> Obtain Count of Invoiced Not Reported
'---> Obtain Count of Reported Not Invoiced
'---> Obtain Count of Actual Stock Available to Report
'---> Obtain Count of Delivered Not Invoiced
'---> Obtain a variance report, if any, and thus tally the ATM vs gDN Stock
'*********************************************************************************************************************

On Error GoTo ErrorHandler

Dim xlWBReport As Workbook, xlWBMaster, xlWBOverall As Workbook
Dim xlWSReport As Worksheet
Dim xlWBWSALStock As Worksheet, xlWBMasterWSALStock As Worksheet
Dim xlWSALSales As Worksheet, xlWSgDNSales As Worksheet
Dim xlWBWSgDNStock As Worksheet, xlWBMasterWSgDNStock As Worksheet


'Check to make sure that the code runs only when the specified file is open.
If fileCheck("LR SALES") = False Then End

'Temporarily Stop Events, Calculations, Screen & Statusbar updations.
Call excelVitals(False)

'Preset LR NC Sales File
Dim wb As Workbook

checkWB = 0

For Each wb In Application.Workbooks

    If wb.CodeName = "LRNCSales" Then
        Set xlWBMaster = wb
        Exit For
    End If
    
Next wb

'Open Excel WB - 'InvoicesVsDelivered' and 'Overall Data' in hidden mode
Set xlWBReport = openExcelWB("Invoice vs Reported Vehicles.xlsb", "\\ho-webstore\ATM\LR and UC Sales Report\LR\General Reports\")
Set xlWBOverall = openExcelWB("Overall Data.xlsb", "\\ho-webstore\ATM\LR and UC Sales Report\LR\General Reports\")

'Preset certain repeated worksheets as variables for efficiency purpose
Set xlWBWSALStock = xlWBReport.Sheets("AL Stock")
Set xlWBWSgDNStock = xlWBReport.Sheets("gDN Stock")
Set xlWBMasterWSALStock = xlWBMaster.Sheets("Stock Spool")
Set xlWBMasterWSgDNStock = xlWBMaster.Sheets("gDN Stock Spool")
Set xlWSALSales = xlWBOverall.Sheets("AL Sales Master")
Set xlWSgDNSales = xlWBOverall.Sheets("gDN Sales Master")

xlWBReport.Activate

'Clear Content from 'AL Stock' Sheet
xlWBWSALStock.UsedRange.Clear


'Clear Content from 'gDN Stock' Sheet
xlWBWSgDNStock.UsedRange.Clear

'Clear Content from 'Invoiced To Date' Sheet
xlWBReport.Sheets("Invoiced To Date").UsedRange.Clear


'Clear Content from 'Reported To Date' Sheet
xlWBReport.Sheets("Reported To Date").UsedRange.Clear

'Clear Content from 'Reported To Date' Sheet
xlWBReport.Sheets("Delivery Date").UsedRange.Clear

'***************************MASTER DATA COPY PROCEDURE****************************
Dim tempArray() As Variant
Dim tempRowCnt As Long
Dim ALStockRowCnt As Long
Dim gDNStockRowCnt As Long

'Copy 'AL Stock' data from Master File to relevant sheet
ALStockRowCnt = xlWBMasterWSALStock.Range("G" & Rows.count).End(xlUp).Row
tempArray = xlWBMasterWSALStock.Range("G1:G" & ALStockRowCnt).Value
xlWBWSALStock.Range("A1:A" & ALStockRowCnt) = tempArray   'Model

tempArray = xlWBMasterWSALStock.Range("J1:J" & ALStockRowCnt).Value
xlWBWSALStock.Range("B1:B" & ALStockRowCnt) = tempArray  'VIN

'Copy 'gDN Stock' data from Master File to relevant sheet
gDNStockRowCnt = xlWBMasterWSgDNStock.Range("H" & Rows.count).End(xlUp).Row
tempArray = xlWBMasterWSgDNStock.Range("H1:H" & gDNStockRowCnt).Value
xlWBWSgDNStock.Range("A1:A" & gDNStockRowCnt) = tempArray 'Model+VIN

tempArray = xlWBMasterWSgDNStock.Range("J1:J" & gDNStockRowCnt).Value
xlWBWSgDNStock.Range("B1:B" & gDNStockRowCnt) = tempArray 'Model+VIN

'Copy 'Invoiced To Date' data from Master File to relevant sheet
tempRowCnt = xlWSALSales.UsedRange.Rows.count
tempArray = xlWSALSales.UsedRange.Value
xlWBReport.Sheets("Invoiced To Date").Range("A1:M" & tempRowCnt) = tempArray


'Copy 'Reported To Date' data from Master File to relevant sheet
tempRowCnt = xlWSgDNSales.UsedRange.Rows.count
tempArray = xlWSgDNSales.UsedRange.Value
xlWBReport.Sheets("Reported To Date").Range("A1:T" & tempRowCnt) = tempArray


'Copy 'NP3 - Delivery Date' data from Master File to relevant sheet
tempRowCnt = xlWBMaster.Sheets("Delivery Master").Range("A" & Rows.count).End(xlUp).Row
tempArray = xlWBMaster.Sheets("Delivery Master").Range("A1:A" & tempRowCnt).Value
xlWBReport.Sheets("Delivery Date").Range("B1:B" & tempRowCnt) = tempArray 'Delivery Date

tempArray = xlWBMaster.Sheets("Delivery Master").Range("B1:B" & tempRowCnt).Value
xlWBReport.Sheets("Delivery Date").Range("A1:A" & tempRowCnt) = tempArray 'VIN
'***************************MASTER DATA COPY PROCEDURE****************************

'Create New Sheet in Excel WorkBook - 'InvoicesVsDelivered' and name it per current date
xlWBReport.Activate
Set xlWSReport = createSheet(xlWBReport)

With xlWSReport
    .Move Before:=Worksheets(Worksheets.count) 'Move new sheet to begining of file
    .name = Format(Date, "dd-mm-yy")
End With

'Load 'AL Stock' & 'gDN Stock' data into a 2-dimensional array & save resultant data in new sheet at cell "A2"

Dim rng1() As Variant, rng2() As Variant

rng1 = xlWBWSALStock.Range("A2:B" & ALStockRowCnt).Value
rng2 = xlWBWSgDNStock.Range("A2:B" & gDNStockRowCnt).Value

With xlWSReport
    .Cells(2, 1).Resize(UBound(rng1), UBound(rng1, 2)) = rng1
    .Cells(2, 1)(UBound(rng1) + 1, 1).Resize(UBound(rng2), UBound(rng2, 2)) = rng2
End With



'Obtain the unique set of data(VIN & Model) from 'AL Stock' & 'gDN Stock' Sheets
xlWSReport.Range("A1:B" & Range("A" & Rows.count).End(xlUp).Row).RemoveDuplicates Columns:=2, header:=xlYes


'Add 3 vlookup formulas to new sheet for obtaining 'Hoegh Osaka', 'Invoiced Date', 'Reported Date' & 'Delivery Date'
'Plus convert formulas to values
Call addFormulas("C2", "=VLOOKUP(RC[-1],HO!C[1],1,FALSE)", True, xlWSReport.name, , , , , "RC")
Call addFormulas("D2", "=INDEX('[Overall Data.xlsb]AL Sales Master'!C[-1],MATCH('" & xlWSReport.name & "'!RC[-2],'[Overall Data.xlsb]AL Sales Master'!C[4],0))", True, xlWSReport.name, , , , , "RC")
Call addFormulas("E2", "=VLOOKUP(RC[-3],'[Overall Data.xlsb]gDN Sales Master'!C[2]:C[4],3,FALSE)", True, xlWSReport.name, , , , , "RC")
Call addFormulas("F2", "=VLOOKUP(RC[-5],'Delivery Date'!C[-5]:C[-4],2,FALSE)", True, xlWSReport.name, , , , , "RC")


'Add header titles to new report columns and bold them
With xlWSReport
    .Range("A1") = "Model"
    .Range("B1") = "VIN"
    .Range("C1") = "H Osaka"
    .Range("D1") = "Invoiced Date (AL - AN6)"
    .Range("E1") = "Reported Date (gDN)"
    .Range("F1") = "Delivery Date"
    .Rows("1:1").Font.Bold = True
End With



'Apply Autofilter, date format to columns D through to F, replace '#N/A' error display, autofit columns.
With xlWSReport
    .UsedRange.AutoFilter
    .Range("D:F").NumberFormat = "m/d/yyyy"
    .Cells.Replace What:="#N/A", Replacement:="", LookAt:=xlWhole, _
                SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, _
                ReplaceFormat:=False
    .Cells.EntireColumn.AutoFit
    .Range("A1").Select
End With

'Recalculate 'Variance Report' to point to new sheet.
With xlWBReport.Sheets("Variance Report")
    .Range("D3").FormulaR1C1 = "=RC[-1]-RC[-2]"
    .Range("G3").FormulaR1C1 = "=RC[-2]-RC[-1]"
    .Range("H3").FormulaR1C1 = "=RC[-4]-RC[-1]"
End With

'Save Excel WB - 'InvoicesVsDelivered' & close
xlWBReport.Save
xlWBOverall.Close SaveChanges = False

'Close workbook & release memory+handle
'Set xlWBReport = Nothing
Set xlWBOverall = Nothing

ErrorHandler:
Call excelVitals(True)


End Sub
