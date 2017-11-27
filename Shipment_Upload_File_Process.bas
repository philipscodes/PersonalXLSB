Attribute VB_Name = "Shipment_Upload_File_Process"
Public Sub shipmentFileProcessor()

Call screenHalt
Call NSCRefFilter
Call shipmentFilter
Call changeAutobioColours
Call changeModelCodes
Call changeDates
Call clearFile
Call screenRelease

MsgBox "File successfully Processed", vbOKOnly, "File Process - Success"

End Sub

Private Sub screenHalt()

    Application.EnableEvents = False
    Application.ScreenUpdating = False

End Sub

Private Sub screenRelease()

    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

Private Sub clearFile()

    Cells.Select
    Range("A1").Activate
    Selection.EntireColumn.Hidden = False
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    
' Delete header and save as Text tab delimited file

    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select
    ActiveWorkbook.SaveAs Filename:="C:\IMPORT\Lrimport.txt", FileFormat:= _
        xlText, CreateBackup:=False


End Sub

Private Sub NSCRefFilter()
Attribute NSCRefFilter.VB_ProcData.VB_Invoke_Func = " \n14"

' Delete Blanks from column I - NSC Ref

Dim myList As ListObject

'mylist = wo
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$FI$10000").AutoFilter Field:=9, Criteria1:="="
    If Range("I:I").SpecialCells(xlCellTypeVisible).Areas.count > 1 Then
    
            Range("I2:I6000").SpecialCells(xlCellTypeVisible, xlErrors).EntireRow.Delete
    End If
    
    ActiveSheet.ShowAllData
    
End Sub

Private Sub shipmentFilter()
Attribute shipmentFilter.VB_ProcData.VB_Invoke_Func = " \n14"

' Delete Shipment that's already uploaded
' inputbox requesting shipment name; if not skip this function

Dim arrShipment() As String
Dim varYN As Boolean
Dim subscriptCnt As Integer



ReDim arrShipment(1) As String
subscriptCnt = 1

shipment:

temparrShipment = Trim(UCase(InputBox("Enter Shipment Name:", "Shipment Name")))

'Checking on empty entry and querying for function exit.
If Trim(temparrShipment) = "" Then
    varYN = MsgBox("Empty Text." & vbCrLf & vbCrLf & "Do you want to skip this function?", vbYesNo)
    If varYN = True Then
        GoTo shipment:
    Else
        End
    End If
End If

'If the above check is passed through then enter temp varibale into array variable.
arrShipment(UBound(arrShipment)) = temparrShipment

'Query if the user would like to enter another shipment name
If MsgBox("Would you like to enter another shipment name?", vbYesNo, "Query") = vbYes Then
    ReDim Preserve arrShipment(UBound(arrShipment) + 1) As String
    GoTo shipment:
End If

For subscriptCnt = LBound(arrShipment) + 1 To UBound(arrShipment)
    'Filter on column BJ and delete unwanted shipment records.
    On Error GoTo errHandlerFilter:
    ActiveSheet.Range("$A$1:$FI$10000").AutoFilter Field:=62, Criteria1:= _
        "=" & arrShipment(subscriptCnt)
    If Range("BJ:BJ").SpecialCells(xlCellTypeVisible).Areas.count > 1 Then
    
            Range("BJ2:BJ10000").SpecialCells(xlCellTypeVisible, xlErrors).EntireRow.Delete
            
    End If
    
        ActiveSheet.ShowAllData
        
Next subscriptCnt

Range("A1").Select

Exit Sub

errHandlerFilter:
    MsgBox Err & ": " & Error(Err)
    varYN = MsgBox("Invalid Shipment Name or Shipment name not found." & vbCrLf & vbCrLf & "Do you want to skip this function?", vbYesNo)
    If varYN = True Then
        GoTo shipment:
    Else
        End
    End If
    

End Sub

Private Sub changeAutobioColours()
Attribute changeAutobioColours.VB_ProcData.VB_Invoke_Func = " \n14"

Dim arrAutobioColours() As String
Dim subscriptCnt As Integer
Dim visualRange As Range
Dim cl As Range

arrAutobioColours = findAutobioColours

If arrAutobioColours(0) = "No Auto" Then
    End
End If

For subscriptCnt = LBound(arrAutobioColours) + 1 To UBound(arrAutobioColours)
    
    'MsgBox arrAutobioColours(subscriptCnt)
    'Autobio colours change name
    If arrAutobioColours(subscriptCnt) = "" Then GoTo skipFor
    
    Range("CK1").Select
    
    
    ActiveSheet.Range("$A:$FI").AutoFilter Field:=89, Criteria1:= _
        "=" & arrAutobioColours(subscriptCnt)
            
    If Range("CK2:CK10000").SpecialCells(xlCellTypeVisible).Rows.count > 0 Then
        
        rowCnt = Range("CK2:CK10000").SpecialCells(xlCellTypeVisible).Rows.count
        Range("CK1").Select
        Set visualRange = Range("CK2", Range("CK2").End(xlDown)).Cells.SpecialCells(xlCellTypeVisible)

        For Each cl In visualRange
            
            cl.Value = Mid(cl.Offset(0, 25).Value, 5)
            cl.Offset(0, -1).Value = Left(cl.Offset(0, 25).Value, 3)
            
        Next
        
        ActiveSheet.ShowAllData
        
    End If
    
skipFor:
Next subscriptCnt

End Sub

Private Function findAutobioColours() As Variant

Dim arrAutobioColours() As String
Dim tempStr As Variant
Dim flagFound As Boolean

ReDim arrAutobioColours(1) As String
flagFound = False

Range("CK1").Select

ActiveCell.Offset(1, 0).Select
Do While Trim(ActiveCell.Value) <> ""
    If Left(UCase(Trim(ActiveCell.Value)), 4) = "AUTO" Then
    
        ' Check to see if current found autobio colour already stored in array variable
        For Each tempStr In arrAutobioColours
            If tempStr = ActiveCell.Value Then
                flagFound = True
                Exit For
            End If
            flagFound = False
        Next tempStr
        
        ' If not previoulsy stored then expand array by 1 and store newly found autobio colour
        If flagFound = False Then
            arrAutobioColours(UBound(arrAutobioColours)) = ActiveCell.Value
            ReDim Preserve arrAutobioColours(UBound(arrAutobioColours) + 1) As String
        End If
    End If
    ActiveCell.Offset(1, 0).Select
Loop

'For subscriptCnt = LBound(arrAutobioColours) + 1 To UBound(arrAutobioColours)
'
'    MsgBox arrAutobioColours(subscriptCnt)
'
'Next subscriptCnt

If arrAutobioColours(1) <> "" Then
    findAutobioColours = arrAutobioColours()
Else
    findAutobioColours = "No Auto"
End If

End Function

Private Sub changeModelCodes()
Attribute changeModelCodes.VB_ProcData.VB_Invoke_Func = " \n14"

Dim recCnt As Integer
Dim cnt As Integer

' Hide J to CR for changing model code
    Columns("CK:DI").Select
    Selection.EntireColumn.Hidden = False
    Range("CK1").Select
    Columns("J:CR").Select
    Selection.EntireColumn.Hidden = True
    Range("CS1").Select
    
' Find record count

    Range("I2").Select
    Range(Selection, Selection.End(xlDown)).Select
    recCnt = Selection.count
    Range("I1").Select

' Logic
    
    If recCnt < 1 Then Exit Sub
    
    For cnt = 1 To recCnt
    
        ActiveCell.Offset(1, 0).Select
        If UCase(Right(ActiveCell.Value, 2)) = "LE" Then
            If UCase(Left(ActiveCell.Value, 5)) = "LR4V8" Then
                Range("CS" & ActiveCell.Row).Value = Range("CS" & ActiveCell.Row).Value & "T"
            Else
                Range("CS" & ActiveCell.Row).Value = Range("CS" & CStr(ActiveCell.Row)).Value & "L"
            End If
        ElseIf UCase(ActiveCell.Value) = "L538DP5" Then
                Range("CS" & ActiveCell.Row).Value = Range("CS" & ActiveCell.Row).Value & "P"
        ElseIf UCase(ActiveCell.Value) = "L538DPC" Then
                Range("CS" & ActiveCell.Row).Value = Range("CS" & ActiveCell.Row).Value & "P"
        ElseIf UCase(Right(ActiveCell.Value, 4)) = "HSES" Then
                Range("CS" & ActiveCell.Row).Value = Range("CS" & ActiveCell.Row).Value & "S"
        End If
        
                
    Next cnt
    
' Unhide all columns
    Cells.Select
    Selection.EntireColumn.Hidden = False
    Range("I1").Select
    
End Sub

Private Sub changeDates()

Dim recCnt As Integer
Dim cl As Range
Dim visualRange As Range
Dim newDate As Date

' Filter out the blanks on column Z

    ActiveSheet.Range("$A$1:$FI$10000").AutoFilter Field:=26, Criteria1:="<>"

' Save Visual Range & record count

    recCnt = ActiveSheet.AutoFilter.Range.Columns(8).SpecialCells(xlCellTypeVisible).Cells.count - 1

    Set visualRange = Range("Z2", Range("Z2").End(xlDown)).Cells.SpecialCells(xlCellTypeVisible)

   
' Exit procedure if record count less than or equal 1

    If recCnt < 1 Then Exit Sub
    
' Last shipment date input

    lastDate = InputBox("Provide the last shipment date.")
        
' Logic 1

    newETA = DateAdd("d", 15, lastDate)
    
    For Each cl In visualRange
        
        If cl.Value < newETA Then
            cl.Value = newETA
        End If
        
    Next
        
    ActiveSheet.ShowAllData

' Filter the blanks on column Z

    ActiveSheet.Range("$A$1:$FI$10000").AutoFilter Field:=26, Criteria1:="="

' Save Visual Range

    recCnt = ActiveSheet.AutoFilter.Range.Columns(25).SpecialCells(xlCellTypeVisible).Cells.count - 1

    Set visualRange = Range("Y2", Range("Y2").End(xlDown)).Cells.SpecialCells(xlCellTypeVisible)
    
' Logic 2 - add 15 days to date in column Y and place in the blanks of column Z
    For Each cl In visualRange
    
            cl.Offset(0, 1).Value = DateAdd("d", 15, cl.Value)
            
    Next
        
    ActiveSheet.ShowAllData

End Sub
