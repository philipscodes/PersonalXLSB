Attribute VB_Name = "Damaged_Removal"
Public Sub damagedOrdersOld()

'Check to make sure that the code runs only when the specified file is open.
If fileCheck("LR SALES") = False Then End

'Temporarily Stop Events, Calculations, Screen & Statusbar updations.
Call excelVitals(False)

Dim strVIN As String
Dim found As Range


'Pick & store New & Old Order# from sheet 'RO SHEET'

ActiveWorkbook.Sheets("Damaged").Activate
'ActiveWorkbook.Sheets("Damaged").Range("D1").Select 'Making sure that the code starts reading from first record always

For i = 1 To 576
    
    'ActiveWorkbook.Sheets("Damaged").Activate
    
    strVIN = ActiveWorkbook.Sheets("Damaged").Range("D1").Offset(i, 0).Value


'Find VIN number in sheet 'Orders Spool'

    Set found = ActiveWorkbook.Sheets("Orders Spool").Cells.Find(What:=strVIN, After:= _
        ActiveCell, LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)

'Delete 1st and 8th cell value of current cell offsets
    'ActiveWorkbook.Sheets("Orders Spool").Activate
    
    If Not found Is Nothing Then
        found.Offset(0, 1).Clear
        found.Offset(0, 8).Clear
        'found.Select
        'ActiveWorkbook.Sheets("Orders Spool").ActiveCell.Offset(0, 1) = ""
        'ActiveWorkbook.Sheets("Orders Spool").ActiveCell.Offset(0, 8) = ""
    Else
        'Shade cell as red & exit loop
        'ActiveWorkbook.Sheets("Damaged").ActiveCell.Cells.colour = 38
        GoTo Repeat_Process
    End If


Repeat_Process: 'Repeat Process
 
'        If I = 50 Or I = 100 Or I = 150 Or I = 200 Or I = 300 Or I = 350 Then
'            MsgBox "500", vbOKOnly
'        End If
Next i

'ActiveWorkbook.Sheets("ATM DASHBOARD").Activate

Call excelVitals(True)

'MsgBox ("Done")

End Sub

Public Sub damagedOrders()

'Check to make sure that the code runs only when the specified file is open.
If fileCheck("LR SALES") = False Then End

'Temporarily Stop Events, Calculations, Screen & Statusbar updations.
Call excelVitals(False)

Dim rowCnt1 As Integer
Dim rowCnt2 As Integer
Dim arrDamageList() As Variant
Dim arrOrdersList() As Variant
Dim Destination As Range

arrDamageList = ActiveWorkbook.Sheets("Damaged").Range("D2:D577").Value         'Store VINs list for damaged vehicles in array
arrOrdersList = ActiveWorkbook.Sheets("Orders Spool").Range("A2:U5000").Value   'Store sheet 'Orders Spool' range in array

For rowCnt1 = LBound(arrDamageList) To UBound(arrDamageList)
    
    If IsEmpty(arrDamageList(rowCnt1, 1)) Then
        Exit For
    End If
'Find VIN number in sheet 'Orders Spool'
    For rowCnt2 = LBound(arrOrdersList) To UBound(arrOrdersList)
    
        If IsEmpty(arrOrdersList(rowCnt1, 1)) Then
            GoTo repeatNext
        End If

        If arrOrdersList(rowCnt2, 7) = arrDamageList(rowCnt1, 1) Then
        
            arrOrdersList(rowCnt2, 8) = ""
            arrOrdersList(rowCnt2, 15) = ""
            Exit For
            
        End If
repeatNext:
    Next rowCnt2

Next rowCnt1

Set Destination = ActiveWorkbook.Sheets("Orders Spool").Range("A2")
Destination.Resize(UBound(arrOrdersList, 1), UBound(arrOrdersList, 2)).Value = arrOrdersList

Call excelVitals(True)

End Sub









