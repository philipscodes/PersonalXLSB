Attribute VB_Name = "POC_Invoice_List"

Sub prepareInvoiceList()
    
'Check to make sure that the code runs only when the specified file is open.
If fileCheck("POC Delivery Master") = False Then End
 
'Temporarily Stop Events, Calculations, Screen & Statusbar updations.
Call excelVitals(False)
    
Sheets("Invoice").Activate

Call deleteColumnsPOC
Call deliveryDatePOC
Call secondSetDatePOC
Call applySort("G", "H")
Call applyBorderFormat

Cells.AutoFilter    'Apply Data Filter

Call excelVitals(True)
    
End Sub

Sub deleteColumnsPOC()

    Call deleteColumns("S", "A1:L1")
    Call deleteColumns("Loc", "A1:K1")
    Call deleteColumns("Order No", "A1:J1")
    Call deleteColumns("Released By", "A1:I1")

End Sub

Sub deliveryDatePOC()

'*****************************************************************************************************************
'This procedure places a 'vlookup' formula into a field to lookup the date of delivery from the main Master sheet*
'*****************************************************************************************************************

    Call addFormulas("H2", "=VLOOKUP(RC[-6],Main!C[-6]:C,7,FALSE)", True, , , , , , "RC")
    
    Call formatRangeAsDate("Invoice", "H")
    
    Range("H1") = "Delivery Date"
    
    Call applyYellowRedhighlight("H1")
    
End Sub

Sub secondSetDatePOC()

'*****************************************************************************************************************
'This procedure places a 'vlookup' formula into a field to lookup the date of delivery from the main Master sheet*
'*****************************************************************************************************************

    'Call addFormulas("I2", "=VLOOKUP(RC[-7],'[JLR ATM.xlsx]ATMC'!C5:C14,10,FALSE)", True)
    
    Call formatRangeAsDate("Invoice", "I")
    
    Range("I1") = "Second Set Date"
    Call applyYellowRedhighlight("I1")

End Sub
