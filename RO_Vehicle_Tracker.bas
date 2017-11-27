Attribute VB_Name = "RO_Vehicle_Tracker"
Sub sortROTracker()

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    ActiveWorkbook.Worksheets("Tracker").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Tracker").Sort.SortFields.Add Key:=Range("B2:B1048576") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Tracker").Sort
        .SetRange Range("A1:I1048576")
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("A1").Select
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    
End Sub


'Below code is deprecated since 21/01/2014 in favour of prepareGDNSheet procedure in gDN_Data_Functions module

'Sub prepareROVehiclesSheet()
'
'Dim LR As Long
'Dim lRow As Integer
'Dim rang As Range
'
'On Error Resume Next
'
''Temporarily Stop Events, Calculations, Screen & Statusbar updations.
'Call excelVitals(False)
'
''Copy new records from 'HH1 Spool' to 'HH1 Cumulated'
'
'    Sheets("HH1 Cumulated").Select
'
'    ActiveSheet.ShowAllData
'
'    Sheets("HH1 Spool").Select
'
'    'Filter '#N/A" records from fresh spool (Vlookup against HH1 Cumulated)
'    ActiveSheet.Range("$A$1:$AI$64900").AutoFilter Field:=35, Criteria1:="#N/A"
'
'    'Copying "#N/A" filtered cells
'    LR = Range("A" & Rows.Count).End(xlUp).Row
'    Range("A2:AH" & LR).SpecialCells(xlCellTypeVisible).Copy
'
'
'    Sheets("HH1 Cumulated").Select
'
'
'    If Trim(Range("A2")) = "" Then
'        Range("A2").Select
'    Else
'        Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Select
'    End If
'
'    Set rang = ActiveCell 'Saving activecell for later
'    ActiveSheet.Paste 'Paste copied cells
'    rang.Select
'    rang.Interior.ColorIndex = 24
'
''On 'HH1 Cumulative' Sheet, filter for MK# 346931 on column#29 - 'AC'(MK Magic)
'
'    ActiveSheet.Range("$A$1:$AI$4298").AutoFilter Field:=29, Criteria1:="=346931"
'
'
''On 'HH1 Cumulative' Sheet, filter for 'S' on column#2 - 'B'(Stock)
'
'    ActiveSheet.Range("$A$1:$AI$4298").AutoFilter Field:=2, Criteria1:="=S"
'
''Hide unwanted columns in 'HH1 Cumulative'
'
'    Columns("N:AG").EntireColumn.Hidden = True
'    Columns("J:L").EntireColumn.Hidden = True
'    Columns("H").EntireColumn.Hidden = True
'    Columns("B:D").EntireColumn.Hidden = True
'
'
''Copy to clipboard filtered values separately for Stock#, XCO No, Model, Exterior Colour, VIN & ETA
'
'    lRow = Range("A" & Rows.Count).End(xlUp).Row
'    Range("A2:M" & lRow).SpecialCells(xlCellTypeVisible).Copy
'
''On 'RO Vehicles' sheet, determine last used row and offset to next row of coloumn A & paste clipboard details.
'
'
'    Sheets("RO Vehicles").Select
'
'    If Trim(Range("A2")) = "" Then
'        Range("A2").Select
'    Else
'        Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Select 'Select the next empty row
'    End If
'
'    Set rang = ActiveCell 'Saving activecell for later
'    ActiveSheet.Paste 'Paste copied cells
'    rang.Select
'    rang.Interior.ColorIndex = 24 'Mark pasted values in purple
'
'
''Unhide unwanted columns in 'HH1 Cumulative'
'
'    Sheets("HH1 Cumulated").Columns("N:AG").EntireColumn.Hidden = False
'    Sheets("HH1 Cumulated").Columns("J:L").EntireColumn.Hidden = False
'    Sheets("HH1 Cumulated").Columns("AH").EntireColumn.Hidden = False
'    Sheets("HH1 Cumulated").Columns("B:D").EntireColumn.Hidden = False
'    ActiveSheet.ShowAllData
'
'Call excelVitals(True)
'
'End Sub
