Attribute VB_Name = "Common_Functions"

Public highLightSwitch As Boolean

Public Sub deleteShapes()

    Dim Shp As Shape
    
    For Each Shp In ActiveSheet.Shapes
    
        Shp.Delete
        
    Next Shp
    
End Sub

Public Sub markDuplicates(colomnLetter As String)

    With Columns(colomnLetter & ":" & colomnLetter)
    
        .FormatConditions.Delete    'Delete any previous conditional formats, if any.
        .FormatConditions.AddUniqueValues
        .FormatConditions(Selection.FormatConditions.count).SetFirstPriority
        .FormatConditions(1).DupeUnique = xlDuplicate
    'Set Exterior and Interior Colours
        .FormatConditions(1).Font.Color = -16383844
        .FormatConditions(1).Interior.PatternColorIndex = xlAutomatic
        .FormatConditions(1).Interior.Color = 13551615
        .FormatConditions(1).StopIfTrue = False 'Continue formating even if condition is met
        
    End With
    
    
End Sub

Public Function totalRecords(startCol As String, startRow As String, Optional dataSheet As String) As Double
    
    Dim startCell As String
    
    startCell = startCol & startRow
    
    If dataSheet <> "" Then
        totalRecords = Sheets(dataSheet).Range(Sheets(dataSheet).Range(startCell), Sheets(dataSheet).Range(startCol & Rows.count).End(xlUp)).count + 1
    Else
        totalRecords = Range(Range(startCell), Range(startCol & Rows.count).End(xlUp)).count + 1
    End If

End Function
Public Sub applyYellowRedhighlight(rangeValue As String)

    Range(rangeValue).Cells.Interior.Color = 65535      'Apply cell fill as yellow
    Range(rangeValue).Font.Color = -16776961            'Apply Red Font


End Sub

Public Sub deleteColumns(searchValue As Variant, searchRowRange As String)

    'Dim colIndex As Long
    
    colIndex = Application.Match(searchValue, Range(searchRowRange), 0)
    If Not IsError(colIndex) Then
        Columns(colIndex).EntireColumn.Delete
    End If

End Sub
  

Public Sub applyBorderFormat()

    With ActiveSheet.UsedRange
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
    End With
    
    With ActiveSheet.UsedRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With ActiveSheet.UsedRange.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With ActiveSheet.UsedRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With ActiveSheet.UsedRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With ActiveSheet.UsedRange.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With ActiveSheet.UsedRange.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With


End Sub

Public Sub applySort(rangeSortKeyColumnLetter As String, rangeLastColumnLetter As String)

Dim cnt As Double

    'Obtaining count of total records from first column
    cnt = Range(Range("A2"), Range("A" & Rows.count).End(xlUp)).count + 1
    
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add Key:=Range( _
        rangeSortKeyColumnLetter & "2:" & rangeSortKeyColumnLetter & cnt), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption:= _
        xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range("A1:" & rangeLastColumnLetter & cnt)
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub

Public Function cellAddressColSplitter(cellAddress As String) As String

    Dim colLetter As String
    Dim strLength As Integer
    Dim cnt As Integer
    Dim dollarCount As Integer
    
    cnt = 1
    strLength = Len(cellAddress)
    
    
    Do While cnt <= strLength
        
        If IsNumeric(Mid(cellAddress, cnt)) = True Then     'Test for numeric instance
                                           
            colLetter = Left(cellAddress, cnt - 1)          'Obtain non-numeric part of string
            Exit Do
            
        End If
            
        cnt = cnt + 1
    
    Loop
    
    'Obtain total occurance of dollar mark within string.
    dollarCount = UBound(Split(colLetter, "$"))
    
    'Remove all instances of dollar mark, if exists.
    colLetter = Replace(colLetter, "$", "", , dollarCount)
    
    cellAddressColSplitter = colLetter
    

End Function


Public Sub addFormulas(formulaStartCell As String, _
formula As String, _
copyPasteAsValues As Boolean, _
Optional sheetName As String, _
Optional totalRowCnt As Integer, _
Optional totalCntRangeStart As String, _
Optional singleCellFormula As Boolean, _
Optional arrayFormula As Boolean, _
Optional formula_RC_A1_Style As String, _
Optional replaceString1 As String, _
Optional replaceString2 As String, _
Optional replaceString3 As String, _
Optional replaceString4 As String, _
Optional replaceString5 As String, _
Optional unselectHeader As Boolean)


'****FOR FUTURE CHANGE CODE to utilize Application.ConvertFormula() - from A1 to R1C1 style  *****
    
  Dim cnt As Double
  Dim tempSheet As String
  Dim formulaPart1 As String
  Dim formulaPart2 As String

    
    'Save the current active sheet to be retrieved later and activate the sheet to input formula
    If Not Trim(sheetName) = vbNullString Then
        tempSheet = ActiveSheet.name
        Sheets(sheetName).Activate
    End If
    
    'Obtaining count of total records
    If totalRowCnt > 0 Then
        cnt = totalRowCnt
    Else
        If totalCntRangeStart = "" Then
            totalCntRangeStart = "A2"
        End If
        colLetter = cellAddressColSplitter(totalCntRangeStart)
        cnt = Range(totalCntRangeStart, Range(colLetter & Rows.count).End(xlUp)).count + 1
    End If

    'Split formula if greater then 255 characters
'    If Len(formula) > 255 Then
'
'        formulaPart1 = Left(formula, 196)
'        formulaPart2 = Right(formula, Len(formula) - Len(formulaPart1))
'        formula = formulaPart1 & "X_X_X"
'
'    End If
    
    'Write formula to first data cell
    If formula_RC_A1_Style = "RC" Then
    
        Range(formulaStartCell).FormulaR1C1 = formula

    ElseIf formula_RC_A1_Style = "A1" And arrayFormula = False Then
    
        Range(formulaStartCell).formula = formula
        
    ElseIf formula_RC_A1_Style = "A1" _
    And arrayFormula = True _
    And Len(replaceString1) > 0 _
    And Len(replaceString2) > 0 _
    And Len(replaceString3) > 0 _
    And Len(replaceString4) > 0 _
    And Len(replaceString5) > 0 Then
    
        Range(formulaStartCell).FormulaArray = formula
        Range(formulaStartCell).Replace "11", replaceString2
        Range(formulaStartCell).Replace "22", replaceString3
        Range(formulaStartCell).Replace "33", replaceString4
        Range(formulaStartCell).Replace "44", replaceString5
        Range(formulaStartCell).Replace "X_C", replaceString1
        
    ElseIf formula_RC_A1_Style = "A1" And arrayFormula = True And Len(replaceString1) Then
    
        Range(formulaStartCell).FormulaArray = formula
        Range(formulaStartCell).Replace "X_C", replaceString1
        
    ElseIf formula_RC_A1_Style = "A1" And arrayFormula = True Then
    
        Range(formulaStartCell).FormulaArray = formula
        
        
    End If
    
    
    'Autofill to all records if required
    If singleCellFormula = False Then
        'Range(formulaStartCell).FillDown
        Range(formulaStartCell).AutoFill Destination:=Range(formulaStartCell & ":" & Left(formulaStartCell, Len(formulaStartCell) - 1) & cnt)
    End If

    'Execute manual calculate
    Range(formulaStartCell & ":" & Left(formulaStartCell, Len(formulaStartCell) - 1) & cnt).Calculate
    
    If copyPasteAsValues = True Then
        
        If singleCellFormula = False Then
        
            'Copy & Paste values to range
            Range(formulaStartCell, Range(formulaStartCell).End(xlDown)).Copy
            Range(formulaStartCell, Range(formulaStartCell).End(xlDown)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            
        Else
        
            'Copy & Paste value to a single cell
            Range(formulaStartCell, Range(formulaStartCell)).Copy
            Range(formulaStartCell, Range(formulaStartCell)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
        
        End If
        
    End If
    
    'Deactivate copy mode
    Application.CutCopyMode = False

    'Select header row
    If unselectHeader = False Then
        Range(formulaStartCell).Offset(-1, 0).Select
    End If
    
    If Not Trim(sheetName) = vbNullString Then
        Sheets(tempSheet).Activate
    End If
    
End Sub

Public Sub formatRangeAsDate(sheetName As String, colum As String)

    Dim cnt As Double
    
    cnt = Range("A2", Range("A" & Rows.count).End(xlUp)).count + 1
    
    Sheets(sheetName).Range(colum & "2:" & colum & cnt).NumberFormat = "dd-mm-yyyy;@"

End Sub


Public Sub findAndDeleteRange(sheetName As String, findvalue As String, rowVal As Integer, colVal As Integer)

    Dim found As Range
    
    'Find 'Franchise' and delete 3 rows hence
    Set found = ActiveWorkbook.Sheets(sheetName).Cells.Find(What:=findvalue, After:= _
        ActiveCell, LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
    
    If Not found Is Nothing Then
        found.Select
        ActiveCell.Resize(rowVal, colVal).ClearContents
    End If

End Sub

Public Sub findAndDeleteCells(sheetName As String, findvalue As String, startCell As String, processRepeat As Integer, ParamArray rowColOffsets() As Variant)

    Dim found As Range
    Dim outerCounter, innerCounter, cnt, upperBound As Integer
    
    cnt = 0
    upperBound = UBound(rowColOffsets())
    
    Range(startCell).Select
    
    Do While outerCounter <= processRepeat
    
        Do While innerCounter < upperBound   'final value of cnt will always be less than upperBound as upperBound is always odd and cnt is incremented by 2
        
            'Find value and delete cells given by ParamArray rowColOffSet()
                Set found = ActiveWorkbook.Sheets(sheetName).Cells.Find(What:=findvalue, After:= _
                    ActiveCell, LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
                    SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
             
                If Not found Is Nothing Then
                    'found.Select
                    'p = found.Offset(rowColOffsets(cnt), rowColOffsets(cnt + 1)).Value
                    found.Offset(rowColOffsets(innerCounter), rowColOffsets(innerCounter + 1)).ClearContents
                End If
                
            innerCounter = innerCounter + 2   'Increasing counter by 2, as each ParamArray subset is a set of (cell) row value and a column value.
        
        Loop
    
    Loop


End Sub


Public Sub addCompanyHonorific()

ActiveCell = "M/s. " & ActiveCell.Value

End Sub

Public Sub clearAllFormats()

    Application.ActiveSheet.Cells.ClearFormats

End Sub

Public Sub clearAllContents()

    Application.ActiveSheet.Cells.Clear

End Sub

Public Sub highLightSwtch()

If highLightSwitch = True Then
    highLightSwitch = False
Else
    highLightSwitch = True
End If

End Sub

Public Sub changeToDate(recordCountRow As String)

Dim tCnt As Long
Dim prevCol As Range

'Save selected cell address
prevCol = ActiveCell.Address

'Obtain Record Count
tCnt = ws.Range(recordCountRow & Rows.count).End(xlUp).Row

'Insert column
prevCol.EntireColumn.Insert

'Transfer Header


'Apply 'DATEVALUE' Formula

'Change new column format to date

'Delete old Date Column

End Sub
Public Function dateOrder() As String

Dim dtFormat As String

Select Case Application.International(xlDateOrder)
        Case Is = 0
            dtFormat = "mm/dd/yyyy"
        Case Is = 1
            dtFormat = "dd/mm/yyyy"
        Case Is = 2
            dtFormat = "yyyy/mm/dd"
        Case Else
            dtFormat = "Error"
End Select


dateOrder = dtFormat

End Function


Public Sub excelVitals(alive As Boolean)

If alive = False Then
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayStatusBar = False
        .Calculation = xlCalculationManual
    End With

ElseIf alive = True Then

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayStatusBar = True
        .Calculation = xlCalculationAutomatic
    End With

End If

End Sub

Public Sub deleteSelectionToLeft()

    Selection.Delete Shift:=xlToLeft

End Sub
Public Sub RefreshAllPivots()

    Dim wks As Worksheet
    Dim pt As PivotTable

    For Each wks In Worksheets
        For Each pt In wks.PivotTables
            pt.RefreshTable
        Next pt
    Next wks
    
End Sub

Public Sub excelVitalsFalse()

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
        .DisplayStatusBar = False
        .Calculation = xlCalculationManual
    End With

End Sub

Public Sub excelVitalsTrue()

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
        .DisplayStatusBar = True
        .Calculation = xlCalculationAutomatic
    End With


End Sub

Public Sub pasteValues()
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

End Sub

Function sendEmail(vTo As String, vSub As String, vMsg As String, Optional wkbk As Workbook, Optional vcc As Variant, Optional vBCC As Variant) As Boolean
' Works in Excel 2000, Excel 2002, Excel 2003, Excel 2007, Excel 2010, Outlook 2000, Outlook 2002, Outlook 2003, Outlook 2007, Outlook 2010.
' This example sends the last saved version of the Activeworkbook object .
    Dim OutApp As Object
    Dim OutMail As Object

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    On Error Resume Next
   ' Change the mail address and subject in the macro before you run it.
    With OutMail
        .To = vTo
        .CC = vcc
        .BCC = vBCC
        .Subject = vSub
        .HTMLBody = vMsg
        If Not wkbk Is Nothing Then
            .Attachments.Add wkbk.FullName
        End If
        ' You can add other files by uncommenting the following line.
        '.Attachments.Add ("C:\test.txt")
        ' In place of the following statement, you can use ".Display" to
        ' display the mail.
        .Display
    End With
    On Error GoTo 0

        On Error Resume Next
        Err.Clear
        ' Send the message
        'OutMail.Send
        If Err.Number = 0 Then
            sendEmail = True
        Else
            sendEmail = False
        End If
           
    Set OutMail = Nothing
    Set OutApp = Nothing
End Function

Sub closeWorkbook()


    If Application.Workbooks.count = 2 Then
        Application.Quit
    End If

End Sub

Function fileCheck(Filename As String) As Boolean

    'Check to make sure that the code runs only when the specified file is open.
    If UCase(Left(ActiveWorkbook.name, Len(Filename))) <> UCase(Filename) Then
        MsgBox "This procedure runs only with the workbook '" & Filename & vbCrLf & vbCrLf & _
        "Program cannot run & will terminate now.", vbOKOnly
        fileCheck = False
        Exit Function
    Else
        fileCheck = True
    End If

End Function

Public Sub clearSheetData(workbookName As String, sheetName As String, recordCountRow As String, strartRange As String, Optional endRangeColumn As String, Optional fullSheet As Boolean)

    Dim tCnt As Double
    Dim ws As Worksheet
    
    'Check to make sure that the code runs only when the specified file is open.
    If fileCheck(workbookName) = False Then End
    
    'Set worksheet
    Set ws = ActiveWorkbook.Sheets(sheetName)
    
    'Obtain total record count
    If fullSheet = True Then
        tCnt = 1048575
        endRangeColumn = "XFD"
    Else
        tCnt = ws.Range(recordCountRow & Rows.count).End(xlUp).Row
    End If
    
    'Clear Range values to Null
    ws.Range(strartRange & ":" & endRangeColumn & tCnt).Clear

End Sub

Public Function openExcelWB(workbookName As String, workbookLocation As String) As Workbook

    Dim xlw As New Excel.Workbook
    
    Set xlw = Application.Workbooks.Open(workbookLocation & workbookName)
    Set openExcelWB = xlw
    
End Function

Public Sub deleteSheetContent(wb As Workbook, sheetName As String, Optional rangeToDelete As String)
    
    If LenB(rangeToDelete) = 0 Then
    
        wb.Sheets(sheetName).Range(wb.Sheets(sheetName).UsedRange).Clear
    
    Else
    
        wb.Sheets(sheetName).Range(rangeToDelete).Clear
        
    End If

End Sub

Public Sub copySheetContent(WBFrom As Workbook, WBTo As Workbook, WSFrom As String, WSTo As String, copyRange As String, pasteRange As String)

    WBTo.Sheets(WSTo).Range(pasteRange) = WBFrom.Sheets(WSFrom).Range(copyRange)
    

End Sub

Function createSheet(wbk As Workbook, Optional newWSName As String) As Worksheet

    If LenB(newWSName) = 0 Then
        Set createSheet = wbk.Sheets.Add
    Else
        Set createSheet = wbk.Sheets.Add.name = newWSName
    End If
    

End Function

Function combine2Arrays(range2 As Range, range1 As Range, Optional WS1 As Worksheet, Optional WS2 As Worksheet) As Variant

    Dim R1 As Range, R2 As Range
    Dim Arr1() As Variant, Arr2() As Variant
    Dim LR1 As Long, LR2 As Long
    Dim rng1 As Variant, rng2 As Variant
    
    'rng1 = Split(range1.Address, "$")(1)
    'rng2 = Split(range2.Address, "$")(1)
    
    'LR1 = WS1.Cells(Rows.Count, rng1).End(xlUp).Row
    'LR2 = WS2.Cells(Rows.Count, rng2).End(xlUp).Row
    
    'Set R1 = Range(Cells(1, rng1), Cells(LR1, rng1))
    'Set R2 = Range(Cells(1, rng2), Cells(LR2, rng2))
    
    Arr1 = range1.Value
    Arr2 = range2.Value
    combine2Arrays = Array(Arr1, Arr2)

End Function

'Main Function
Function SpellNumber(ByVal MyNumber)
    Dim Dollars, Cents, temp
    Dim DecimalPlace, count
    ReDim Place(9) As String
    Place(2) = " Thousand "
    Place(3) = " Million "
    Place(4) = " Billion "
    Place(5) = " Trillion "
    ' String representation of amount.
    MyNumber = Trim(Str(MyNumber))
    ' Position of decimal place 0 if none.
    DecimalPlace = InStr(MyNumber, ".")
    ' Convert cents and set MyNumber to dollar amount.
    If DecimalPlace > 0 Then
        Cents = GetTens(Left(Mid(MyNumber, DecimalPlace + 1) & _
                  "00", 2))
        MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
    End If
    count = 1
    Do While MyNumber <> ""
        temp = GetHundreds(Right(MyNumber, 3))
        If temp <> "" Then Dollars = temp & Place(count) & Dollars
        If Len(MyNumber) > 3 Then
            MyNumber = Left(MyNumber, Len(MyNumber) - 3)
        Else
            MyNumber = ""
        End If
        count = count + 1
    Loop
    Select Case Dollars
        Case ""
            Dollars = "No Dollars"
        Case "One"
            Dollars = "One Dollar"
         Case Else
            Dollars = Dollars & " Dollars"
    End Select
    Select Case Cents
        Case ""
            Cents = " and No Cents"
        Case "One"
            Cents = " and One Cent"
              Case Else
            Cents = " and " & Cents & " Cents"
    End Select
    SpellNumber = Dollars & Cents
End Function
      
' Converts a number from 100-999 into text
Function GetHundreds(ByVal MyNumber)
    Dim Result As String
    If Val(MyNumber) = 0 Then Exit Function
    MyNumber = Right("000" & MyNumber, 3)
    ' Convert the hundreds place.
    If Mid(MyNumber, 1, 1) <> "0" Then
        Result = GetDigit(Mid(MyNumber, 1, 1)) & " Hundred "
    End If
    ' Convert the tens and ones place.
    If Mid(MyNumber, 2, 1) <> "0" Then
        Result = Result & GetTens(Mid(MyNumber, 2))
    Else
        Result = Result & GetDigit(Mid(MyNumber, 3))
    End If
    GetHundreds = Result
End Function
      
' Converts a number from 10 to 99 into text.
Function GetTens(TensText)
    Dim Result As String
    Result = ""           ' Null out the temporary function value.
    If Val(Left(TensText, 1)) = 1 Then   ' If value between 10-19...
        Select Case Val(TensText)
            Case 10: Result = "Ten"
            Case 11: Result = "Eleven"
            Case 12: Result = "Twelve"
            Case 13: Result = "Thirteen"
            Case 14: Result = "Fourteen"
            Case 15: Result = "Fifteen"
            Case 16: Result = "Sixteen"
            Case 17: Result = "Seventeen"
            Case 18: Result = "Eighteen"
            Case 19: Result = "Nineteen"
            Case Else
        End Select
    Else                                 ' If value between 20-99...
        Select Case Val(Left(TensText, 1))
            Case 2: Result = "Twenty "
            Case 3: Result = "Thirty "
            Case 4: Result = "Forty "
            Case 5: Result = "Fifty "
            Case 6: Result = "Sixty "
            Case 7: Result = "Seventy "
            Case 8: Result = "Eighty "
            Case 9: Result = "Ninety "
            Case Else
        End Select
        Result = Result & GetDigit _
            (Right(TensText, 1))  ' Retrieve ones place.
    End If
    GetTens = Result
End Function
     
' Converts a number from 1 to 9 into text.
Function GetDigit(Digit)
    Select Case Val(Digit)
        Case 1: GetDigit = "One"
        Case 2: GetDigit = "Two"
        Case 3: GetDigit = "Three"
        Case 4: GetDigit = "Four"
        Case 5: GetDigit = "Five"
        Case 6: GetDigit = "Six"
        Case 7: GetDigit = "Seven"
        Case 8: GetDigit = "Eight"
        Case 9: GetDigit = "Nine"
        Case Else: GetDigit = ""
    End Select
End Function

