Attribute VB_Name = "Find_Mark"
Public NextBlink As Double
Public BlinkCell As String 'The cell that you want to blink
Public blinkCounter As Integer

Public Sub Find_Mark_Main()

'Load the form to select workbook and worksheet on which stock reconciliation operation is to be carries out.

Load usrFrmStock
usrFrmStock.Show

End Sub

Public Sub mainExit()

ThisWorkbook.Close SaveChanges:=False

End Sub

Function findvalue(fndvalue As String) As Range

    Dim foundRange As Range

    'Search for the VIN in Stock Excel Sheet and input result to range variable 'foundRange'
    Set foundRange = Cells.Find(fndvalue, After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False)
        
    If Not foundRange Is Nothing Then
        Set findvalue = foundRange
        BlinkCell = foundRange.Address
    Else
        Set findvalue = Nothing
    End If

End Function

Sub revertHighlightedEntry()

On Error Resume Next
    'Reverting the back & fore colour of previous cell to Automatic(Black)
    ActiveCell.Previous.Activate
    ActiveCell.Font.ColorIndex = xlAutomatic
    ActiveCell.Offset(0, 1).Font.ColorIndex = xlAutomatic

End Sub

Sub highlightEnry()

'Highlight using Cell Colour Fill

    ActiveCell.Cells.Interior.Color = 65535
    'ActiveCell.Offset(0, 1).Cells.Interior.Color = 65535


'Highlight using Font

'    ActiveCell.Font.Color = -16776961
'    ActiveCell.Offset(0, 1).Font.Color = -16776961
    
End Sub



Sub closingProc()

On Error Resume Next

        ActiveCell.Font.ColorIndex = xlAutomatic
        ActiveCell.Offset(0, 1).Font.ColorIndex = xlAutomatic
        
    If ActiveCell.Previous.Font.Color = -16776961 Then
        ActiveCell.Previous.Font.ColorIndex = xlAutomatic
        ActiveCell.Previous.Offset(0, 1).Font.ColorIndex = xlAutomatic
    End If

End Sub


Sub txtBxFocus()

    usrFrmStock.txtVIN.SetFocus
    usrFrmStock.txtVIN.SelStart = 0
    usrFrmStock.txtVIN.SelLength = Len(usrFrmStock.txtVIN.Text)

End Sub


'Start blinking
Public Sub StartBlinking()


    'Application.Goto Range("A1"), 1
    'If the color is red, change the color and text to white
    If Range(BlinkCell).Interior.ColorIndex = 3 Then
        Range(BlinkCell).Interior.ColorIndex = 0
    'If the color is white, change the color and text to red
    Else
        Range(BlinkCell).Interior.ColorIndex = 3
    End If
    'Wait one second before changing the color again
    If blinkCounter = 3 Then
        Call StopBlinking
    Else
        blinkCounter = blinkCounter + 1
        NextBlink = Now + TimeSerial(0, 0, 1)
        Application.OnTime NextBlink, "StartBlinking", , True
    End If
    
End Sub

'Stop blkinking
Private Sub StopBlinking()

    blinkCounter = 0
    'Set color to white
    Range(BlinkCell).Interior.ColorIndex = 0
    On Error Resume Next
    Application.OnTime NextBlink, "StartBlinking", , False
    Err.Clear
    
    Call highlightEnry
    ActiveWorkbook.Save
    Call txtBxFocus

    
End Sub
