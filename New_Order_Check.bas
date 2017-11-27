Attribute VB_Name = "New_Order_Check"
Public Sub checkOrders()

Call excelVitals(False)
'Pick & store New & Old Order# from sheet 'RO SHEET'

Dim oldOrder As String
Dim newOrder As String
Dim sht As Worksheet
Dim found As Range
Dim oldOrderFound As Range
Dim newOrderFound As Range
Dim cnt As Long
Dim i As Integer
Dim j As Integer


cnt = InputBox("Number of orders to check?")

ActiveWorkbook.Sheets("RO Sheet").Activate
ActiveWorkbook.Sheets("RO Sheet").Range("J1").Select 'Making sure that the code starts reading from first record always

For i = 1 To cnt
    
    ActiveWorkbook.Sheets("RO Sheet").Activate
    
    oldOrder = ActiveCell.Offset(i, 1).Value
    newOrder = ActiveCell.Offset(i, 0).Value


'Find new order number in sheet 'New Orders VISTA'
    Set found = ActiveWorkbook.Sheets("New Orders VISTA").Cells.Find(What:=oldOrder, After:= _
        ActiveCell, LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)

'Goto 1st column of row
    ActiveWorkbook.Sheets("New Orders VISTA").Activate
    
    If Not found Is Nothing Then
        found.Select
        Application.Range(ActiveCell.EntireRow.Address)(1, 1).Select
        Set oldOrderFound = ActiveCell 'Goto begining of found row
    Else
        'Shade cell as red & exit loop
        ActiveWorkbook.Sheets("RO Sheet").ActiveCell.Offset(i, 0).Cells.colour = 38
        Exit For
    End If

'Find old order number in sheet 'Old Order VISTA'
    Set found = ActiveWorkbook.Sheets("Old Order VISTA").Cells.Find(What:=newOrder, After:= _
        ActiveCell, LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)

    
'Goto 1st column of row.
    ActiveWorkbook.Sheets("Old Order VISTA").Activate
    
    If Not found Is Nothing Then
        found.Select
        Application.Range(ActiveCell.EntireRow.Address)(1, 1).Select
        Set newOrderFound = ActiveCell 'Goto begining of found row
    Else
        'Shade cell as red & exit loop
        ActiveWorkbook.Sheets("RO Sheet").ActiveCell.Offset(i, 1).Interior.ColorIndex = 38
        Exit For
    End If

'Loop to check each cell of each sheet
    For j = 2 To 59
    
        ActiveWorkbook.Sheets("New Orders VISTA").Activate
        newOrd = ActiveCell.Offset(0, j).Value
        
        ActiveWorkbook.Sheets("Old Order VISTA").Activate
        oldOrd = ActiveCell.Offset(0, j).Value
        
        If newOrd <> oldOrd Then
            ActiveWorkbook.Sheets("New Orders VISTA").Activate
            ActiveCell.Offset(0, j).Interior.ColorIndex = 37
        End If
'        If j = 40 Then
'            Stop
'        End If
    Next j
    
'Repeat Process

Next i

Call excelVitals(True)

MsgBox ("Done")

End Sub



