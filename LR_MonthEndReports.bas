Attribute VB_Name = "LR_MonthEndReports"

Sub LR_Summary_Report()

'*****************************************************************************
'This procedure prepares the 'ATM LR Stock & Sales Summary' monthly report.*
'*****************************************************************************

'Check to make sure that the code runs only when the specified file is open.
If fileCheck("LR SALES") = False Then End


'Temporarily Stop Events, Calculations, Screen & Statusbar updations.
Call excelVitals(False)

Dim thisSheet As Worksheet

Set thisSheet = Sheets("ATM Sum. Rep.")

'On Error Resume Next
'ActiveSheet.ShowAllData
thisSheet.ShowAllData
'On Error GoTo 0



End Sub
