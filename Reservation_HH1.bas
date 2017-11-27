Attribute VB_Name = "Reservation_HH1"

Sub checkReservations()

Dim visRecordCnt As Integer
Dim bulkEnquiryFlag As Boolean: bulkEnquiryFlag = False
Dim defLE As Boolean: defLE = False
Dim leasedVehicleFlag As Boolean: leasedVehicleFlag = False
Dim vipFlag As Boolean: vipFlag = False
Dim depositFlag As Boolean: depositFlag = False
Dim rereserve As Boolean: rereserve = False
Dim unreserve As Boolean: unreserve = False
Dim model As String
Dim modelVariant As String
Dim enquiryNumber As String
Dim stockNumber As String
Dim vehicleStatus As String
Dim salesAdvisor As String
Dim MK As String
Dim ETA As Variant
Dim reservedFromDate As Variant
Dim reservedTillDate As Variant
Dim reservationPeriod As Integer
Dim stockAge As Integer
Dim modelYear As Integer
Dim deposit1 As Variant
Dim deposit2 As Variant
Dim deposit3 As Variant
Dim customerName As String
Dim companyName As String
Dim arrSplitCustName() As String
Dim arrVIPNamesDB() As Variant
Dim errorCode As Integer

On Error GoTo Whoa

'Check to make sure that the code runs only when the specified file is open.
If fileCheck("LR SALES") = False Then End

'Temporarily Stop Events, Calculations, Screen & Statusbar updations.
Call excelVitals(False)

Sheets("HH1 Spool").Activate   'Activate sheet 'HH1 Spool'

On Error GoTo -1

'If any filter is active, show all data
On Error Resume Next
    ActiveSheet.ShowAllData
On Error GoTo -1

On Error GoTo Whoa

'1st BUSINESS LOGIC: Apply filter on column 'AA'(Reserved Untill) for current dates - if current day is Thursday, also filter Friday date
If Weekday(Now()) = 5 Then 'If Thursday then include Friday Date

    ActiveSheet.Range("$A$1:$AI$4298").AutoFilter Field:=27, _
    Criteria1:=">=" & Format(Date, "mm/dd/yyyy"), _
    Operator:=xlAnd, _
    Criteria2:="<=" & Format(Date + 1, "mm/dd/yyyy")

Else

    ActiveSheet.Range("$A$1:$AI$4298").AutoFilter Field:=27, _
    Criteria1:=">=" & Format(Date, "mm/dd/yyyy"), _
    Operator:=xlAnd, _
    Criteria2:="<=" & Format(Date, "mm/dd/yyyy")   '***********Remove +3 after testing**************

End If


'Obtaining total filtered(visible) record count
visRecordCnt = ActiveSheet.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.count - 1
    
'Check if visible area count is more than 1 - that is if any records exist besides header
If visRecordCnt <= 0 Then

    MsgBox "There are no reservations expiring today.", vbOKOnly, "No Reservation Expiry"
    GoTo Whoa
    
End If


Range("A1").Select 'make cell 'A1' active

For Each cell In ActiveSheet.AutoFilter.Range.Columns(4).SpecialCells(xlCellTypeVisible).Cells  'Column 'D' - Model Code _
                                                                                                (First Reference Point for navigation)
    
    model = cell.Value
    
    
    '2ND BUSINESS LOGIC: Skip if vehicle model is 'L405' / 'L494' / Header row
    If model = "RANGEROVERNEW" Or model = "RRSPORTNEW" Or model = "Model Code" Then
    
        GoTo continueLoop
        
    End If
    
    'Reading all required data to variables
    stockNumber = cell.Offset(0, -3).Value
    vehicleStatus = cell.Offset(0, -2).Value
    modelVariant = cell.Offset(0, 2).Value
    ETA = CDate(cell.Offset(0, 9).Value)
    modelYear = CInt(cell.Offset(0, 10).Value)
    enquiryNumber = cell.Offset(0, 12).Value
    deposit1 = cell.Offset(0, 14).Value
    deposit2 = cell.Offset(0, 15).Value
    deposit3 = cell.Offset(0, 16).Value
    customerName = cell.Offset(0, 19).Value
    companyName = cell.Offset(0, 20).Value
    reservedFromDate = CDate(cell.Offset(0, 22).Value)
    reservedTillDate = CDate(cell.Offset(0, 23).Value)
    salesAdvisor = cell.Offset(0, 24).Value
    MK = cell.Offset(0, 25).Value
    
    
    '3RD BUSINESS LOGIC: Check if column 'F'(Variant) has the word 'Adventure' or 'Heritage' as part of the Defender model names.
    If InStr(UCase(modelVariant), "ADV") > 0 Or InStr(UCase(modelVariant), "HERI") > 0 Then
    
        defLE = True 'Set Defender LE flag to 'TRUE' for future use in function
    
    End If
    
    '4TH BUSINESS LOGIC: Check if enquiry# is part of 'Bulk Enquiry Spool' sheet list
    If Not Sheets("Bulk Enquiry Spool").Range("A:A").Find(What:=enquiryNumber, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False) Is Nothing Then
    
        bulkEnquiryFlag = True 'Set bulk enquiry flag to 'True' for future use in function.
        
    End If
    
    '5TH BUSINESS LOGIC: Check if column 'W'(Customer Name) has VIP names/titles('Ruling Family', 'Other Sheikhs', and 'Government Organization', etc)
    customerName = UCase(Trim(customerName))      'Setting variable to its uppercase & removing all leading and trailing spaces
    
        'arrSplitCustName = Split(customerName, " ")         'Splitting customer name at space and holding values in array.
        arrVIPNamesDB = Sheets("Data").Range("AB2:AB100").Value   'Populating array with data from column 'AB'(VIP NAMES) of 'Data' sheet
        
'        For cnt = LBound(arrSplitCustName) To UBound(arrSplitCustName)  'Executing For loop on each element of customer name array
'            For cnt2 = LBound(arrVIPNamesDB) To UBound(arrVIPNamesDB)   'Executing nested For loop on each element of VIP NAMES array
'                If arrSplitCustName(cnt) = arrVIPNamesDB(cnt2, 1) Then
'                    vipFlag = True 'Set VIP flag to 'True' for future use in function.
'                    Exit For
'
'                End If
'            Next cnt2
'            If vipFlag = True Then Exit For
'        Next cnt
        
        For cnt2 = LBound(arrVIPNamesDB) To UBound(arrVIPNamesDB)   'Executing nested For loop on each element of VIP NAMES array
        
            If InStr(customerName, Trim(arrVIPNamesDB(cnt2, 1))) > 0 Then
            
                vipFlag = True 'Set VIP flag to 'True' for future use in function.
                Exit For
            
            End If
            
        Next cnt2
    
    
    '6TH BUSINESS LOGIC: Check if column 'X' 'Company Name'(20) has 'ATM JLR Leased Vehicle'
    If UCase(companyName) = "ATM JLR LEASED VEHICLE" Then
    
        leasedVehicleFlag = True 'Set leased vehicle flag to 'True' for future use in function.
        
    End If
    
    'Set arguements for Quick Macro on column 'AM' of 'HH1 Spool ' sheet
    Range("AM3") = enquiryNumber 'Enquiry#
    Range("AM4") = stockNumber 'Stock#
    
    'Check if deposit is available.
    If Not Trim(deposit1) = Null Or Not Trim(deposit2) = Null Or Not Trim(deposit3) = Null Then
        depositFlag = True 'Setting deposit flag to true if deposit on any of the 3 linked is found
    End If
    
    '7TH BUSINESS LOGIC: 'Check Dates (LR4, EVOQUE, DISCOVERY, DEFENDERS - NON HERITAGE/ADVENTURE)
    reservationPeriod = (reservedTillDate - reservedFromDate) + 1 'Obtaining reservation period
    stockAge = (Date - ETA) + 1 'Obtaining stock age
    
    
    If modelYear = 2015 Then 'And depositFlag = True Then
    
        If reservationPeriod < 30 Then
            unreserve = True
        ElseIf reservationPeriod >= 30 And stockAge > 30 Then
            unreserve = True
        ElseIf reservationPeriod >= 30 And stockAge < 30 Then
            rereserve = True
        End If
        
    End If
    
    
    If modelYear = 2016 And depositFlag = True Then
        If reservationPeriod < 30 Then
            rereserve = True
        ElseIf reservationPeriod >= 30 And stockAge < 45 Then
            rereserve = True
        ElseIf reservationPeriod >= 30 And stockAge > 45 Then
            unreserve = True
        End If
    ElseIf modelYear = 2016 And depositFlag = False Then
        If reservationPeriod < 30 Then
            unreserve = True
        ElseIf reservationPeriod >= 30 And stockAge < 45 Then
            rereserve = True
        ElseIf reservationPeriod >= 30 And stockAge > 45 Then
            unreserve = True
        End If
    End If
    
    
    '8TH BUSINESS LOGIC: Call function to check column 'B'(Status) for Stock(S) or on Order(O) vehicles and thereby call respective Autoline unreserve/re-reserve routine.
    'If the function returns 'True' then call email procedure
    If unreserveReserve(vehicleStatus, bulkEnquiryFlag, defLE, leasedVehicleFlag, vipFlag, ETA, rereserve, unreserve) = True Then
    
        If vehicleStatus <> "O" Then
            If sendReservationEmailNotification(vehicleStatus, salesAdvisor, _
            enquiryNumber, stockNumber, MK, _
            customerName, modelVariant) = False Then
                'ActiveWorkbook.Activate
                MsgBox "The email notification was not sent", vbOKOnly, "Email Not Sent"
            Else
                'ActiveWorkbook.Activate
                MsgBox "Email Sent Successfully", vbOKOnly, "Email Sent"
            End If
        End If
        
    Else
    
        'Highlight entire row as red in Excel
        
    End If
    
    'Reset all flags to 'False' in preparation for next record(loop)
    bulkEnquiryFlag = False
    defLE = False
    leasedVehicleFlag = False
    vipFlag = False
    depositFlag = False

continueLoop:
Next cell

MsgBox "HH1 Report analyses, completed successfully", vbOKOnly, "Analyses Completed"

Whoa:
Call excelVitals(False)


End Sub

Function unreserveReserve(Status As String, _
bulkEnquiryFlag As Boolean, _
defLE As Boolean, _
leasedVehicleFlag As Boolean, _
vipFlag As Boolean, _
ETA As Variant, _
rereserve As Boolean, _
unreserve As Boolean) As Boolean

Dim wsh As Object               'Holds a WSrcipt shell executed program
Dim retVal As Variant           'Holds the return value
Dim waitOnReturn As Boolean: waitOnReturn = True
Dim windowStyle As Integer: windowStyle = 1


Set wsh = VBA.CreateObject("WScript.Shell") 'Initiate variable object as 'Windows Shell Script'

If Status = "S" Or Status = "T" Then        'If Stock(S) or Transfer(T) status then

    
    If bulkEnquiryFlag = True Or leasedVehicleFlag = True Or vipFlag = True Or defLE = True Then
    
        'If column 'B'(Status) is a Stock(S) vehicle and is against a bulk enquiry - run Re-Reserve EXE via autoline
        
        Range("AM2") = 2    'Set QM 'Macro Type' argument for unreserve & then rereserve for stock vehicles
        retVal = wsh.Run("C:\Temp\Reservation.exe", windowStyle, waitOnReturn)
        GoTo Result
    
'    ElseIf defLE = True Then
'
'        'If column 'B'(Status) is a Stock(S) vehicle and is against a bulk enquiry - run unreserve + BulkReserve EXE via autoline
'
'        Range("AM2") = 1    'Set QM 'Macro Type' argument for unreserve
'        'MsgBox "I am in Stock - defLE", vbOKOnly
'        retVal = wsh.Run("C:\Temp\Reservation.exe", windowStyle, waitOnReturn)
'        'retVal = wsh.Run("C:\Temp\Bulkreserve.exe", windowStyle, waitOnReturn)
'        GoTo result
    ElseIf rereserve = True Then
    
        'Range("AM2") = 2    'Set QM 'Macro Type' argument for unreserve & then rereserve for stock vehicles
        retVal = wsh.Run("C:\Temp\Reservation.exe", windowStyle, waitOnReturn)
        GoTo Result
        
    ElseIf unreserve = True Then
    
        Range("AM2") = 2    'Set QM 'Macro Type' argument for unreserve & then rereserve for stock vehicles
        retVal = wsh.Run("C:\Temp\Reservation.exe", windowStyle, waitOnReturn)
        GoTo Result
        
    Else
    
        'If column 'B'(Status) is a Stock(S) vehicle - run unreserve EXE via autoline
        Range("AM2") = 1    'Set QM 'Macro Type' argument for unreserve
        retVal = wsh.Run("C:\Temp\Reservation.exe", windowStyle, waitOnReturn)
        GoTo Result
    
    End If

ElseIf Status = "O" Then                'If Order(O) status then

    'If column 'B'(Status) is an 'On Order(O)' vehicle - run re-reserve EXE via autoline = ETA + 30 days
    Range("AM2") = 3    'Set QM 'Macro Type' argument for unreserve & then rereserve for order vehicles
    reservationDays = (ETA - Date) + 30 'Calculate difference of ETA date against current date and add 30 to it.
    If reservationDays > 90 Then
        Range("AM4") = 90
    Else
        Range("AM4") = reservationDays
   End If
    retVal = wsh.Run("C:\Temp\Reservation.exe", windowStyle, waitOnReturn)
    GoTo Result

End If


Result:
    If retVal = 0 Then
        MsgBox "Autoline procedure executed successfully.", vbOKOnly, "AL Procedure Success"
        unreserveReserve = True
    Else
        MsgBox "Program exited with error code " & retVal & ".", vbOKOnly, "AL Procedure Unsuccessfull"
        unreserveReserve = False
    End If

End Function

Function sendReservationEmailNotification( _
Status As String, _
salesAdvisor As String, _
enquiry As String, _
stock As String, _
MK As String, _
custName As String, _
modelVariant As String) As Boolean

Dim vTo As String
Dim vcc As String
Dim vSub As String
Dim vMsg As String
Dim tempRange As Range

'Prepare email.
If Status = "S" Or Status = "T" Then

'Find sales advisor email address and add to vTo variable
    'Check if enquiry# is part of 'Bulk Enquiry Spool' sheet list
    Set tempRange = Sheets("Email Address").Range("A:A").Find(What:=salesAdvisor, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
    
    If Not tempRange Is Nothing Then
        vTo = tempRange.Offset(0, 1).Value
    End If
    
'Find SM, SRM & ASRM email address of corresponding sales advisor and add these and central administrators to vCC variable
    Set tempRange = Sheets("Reporting Managers").Columns("A:A").Find(What:=salesAdvisor, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
    
    If Not tempRange Is Nothing Then
        vcc = tempRange.Offset(0, 1).Value
        vcc = vcc & ";" & tempRange.Offset(0, 2).Value
        vcc = vcc & ";" & tempRange.Offset(0, 3).Value
        vcc = vcc & ";" & "aeapen@altayer.com"
    End If

'Prepare subject

    vSub = "Notification Email - Vehicle Unreserved"
    
'Prepare message

    vMsg = "Dear " & Left(salesAdvisor, InStr(salesAdvisor, " ") - 1) & "<br><br>" & _
            "Kindly note that the reservation period for the below mentioned vehicle has expired as of today." & "<br>" & _
            "The same has been unreserved from enquiry and put back to stock." & "<br><br>" & _
            "<b>Enquiry#: </b>" & enquiry & "<br>" & _
            "<b>Stock#: </b>" & stock & "<br>" & _
            "<b>MK#: </b>" & MK & "<br>" & _
            "<b>Customer Name: </b>" & custName & "<br>" & _
            "<b>Model Variant: </b>" & modelVariant & "<br><br>" & _
            "Best Regards" & "<br>" & _
            "Philip"
            
    
End If

If sendEmail(vTo, vSub, vMsg, , vcc) = True Then
    sendReservationEmailNotification = True
Else
    sendReservationEmailNotification = False
End If

End Function
