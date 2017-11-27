Attribute VB_Name = "HH1_Reservation_Report"
Sub prepareHH1BookingsReport()

'Check to make sure that the code runs only when the specified file is open.
If fileCheck("JLR Bookings Report") = False Then End

'Turn off all Excel events
Call excelVitals(False)



Dim recordCnt As Double
Dim prevDate As String


'Obtain previous day's date.

If Weekday(Now() - 1) = vbFriday Then
    prevDate = Format((Now() - 2), "dd/mm/yyyy")
Else
    prevDate = Format((Now() - 1), "dd/mm/yyyy")
End If


'Filter previous day's date from column Z of sheet 'HH1 Spool'
Sheets("HH1 Spool").Range("$A$1:$AH$10000").AutoFilter Field:=26, Criteria1:=prevDate


'Find total records
recordCnt = totalRecords("A", "2")


'Copy filtered records from columns A to AG and append to 'HH1 Master' Sheet
Sheets("HH1 Spool").Range("A2:AG" & recordCnt).SpecialCells(xlCellTypeVisible).Copy _
    Destination:=Sheets("HH1 Master").Range("A1").End(xlDown).Offset(1, 0)


'Populate 'DATEVALUE' formula to new records in column AI
Call addFormulas("AI2", "=DATEVALUE(RC[-9])", False, "HH1 Master", , , , , "RC")

'Turn on all Excel events
Call excelVitals(True)

End Sub
