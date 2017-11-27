Attribute VB_Name = "Regional_Office"
Public Sub clearHeaderSpace_gDN_VISTA_SpooledReports()

    Rows("1:1").Replace What:=" ", Replacement:="_", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
End Sub

Public Sub preparegDNSales()
   
    'Check to make sure that the code runs only when the specified file is open.
    If fileCheck("LR SALES") = False Then End
    
    'Temporarily Stop Events, Calculations, Screen & Statusbar updations.
    Call excelVitals(False)
        
    Call gDNSalesSpoolFormula
    'Call gDNSalesMasterTransfer
        
    Call excelVitals(True)
    
End Sub


Private Sub gDNSalesMasterTransfer()

    Dim rowCnt As Integer
    Dim xlWBLRSALES As Workbook
    Dim xlWBOverall As Workbook
    Dim xlWSgDNSales As Worksheet
    Dim xlWSgDNSalesMaster As Worksheet
    
    rowCnt = totalRecords("A", "2", "gDN Sales Spool")
    
    'Open Excel WB - 'Overall Data' and set 'gDN Sales Master' sheet
    Set xlWBLRSALES = ActiveWorkbook
    Set xlWBOverall = openExcelWB("Overall Data.xlsb", "\\ho-webstore\ATM\LR and UC Sales Report\LR\General Reports\")
    Set xlWSgDNSalesMaster = xlWBOverall.Sheets("gDN Sales Master")
    Set xlWSgDNSales = xlWBLRSALES.Sheets("gDN Sales Spool")
    
    'Turn off Autofilter for sheet and re-set the filter per given criteria
    'Copy new records to 'gDN Sales Master' sheet
    With xlWSgDNSales
        .AutoFilterMode = False
        With .Range("A1:X" & rowCnt)
            .AutoFilter Field:=21, Criteria1:="#N/A"
        End With
        .Range("A2:T" & rowCnt).SpecialCells(xlCellTypeVisible).Copy Destination:=xlWSgDNSalesMaster.Range("A1").End(xlDown).Offset(1, 0)
    End With
    
    'Save Excel WB - 'Overall Data' & close
    xlWBOverall.Close SaveChanges = True
    
    'Close workbook & release memory+handle
    Set xlWBOverall = Nothing

End Sub


Private Sub gDNSalesSpoolFormula()
    
    'Temporarily Stop Events, Calculations, Screen & Statusbar updations.
    Call excelVitals(False)
    
    Sheets("gDN Sales Spool").Activate
    
        Call addFormulas("F2", "=INDEX('[Overall Data.xlsb]HH1 Master'!C6,MATCH(RC[1],'[Overall Data.xlsb]HH1 Master'!C9,0))", True, , , , , , "RC")
        Call addFormulas("U2", "=VLOOKUP(RC[-14],'[Overall Data.xlsb]gDN Sales Master'!C7,1,FALSE)", True, , , , , , "RC")
        Call addFormulas("V2", "=VLOOKUP(RC[-15],'Vista Sales Spool'!C[-21]:C[-18],4,FALSE)", False, , , , , , "RC") 'CHANGED'
        Call addFormulas("W2", "=VLOOKUP(RC[-16],'[Overall Data.xlsb]HH1 Master'!C9:C14,6,FALSE)", True, , , , , , "RC")
    
    'Change Header
        Sheets("gDN Sales Spool").Range("F1") = "AutoLine Based Model (HH1)"
        Sheets("gDN Sales Spool").Range("U1") = "Overall gDN Reported"
        Sheets("gDN Sales Spool").Range("V1") = "VISTA Reported Sale Type"
        Sheets("gDN Sales Spool").Range("W1") = "Model Year (HH1)"
    
    Call clearROModelCode
    
    Sheets("gDN Sales Spool").Activate
    
    Call excelVitals(True)

End Sub

Public Sub gDNStockSpoolFormula()

'Temporarily Stop Events, Calculations, Screen & Statusbar updations.
Call excelVitals(False)

Sheets("gDN Stock Spool").Activate

    Call addFormulas("H2", "=INDEX('[Overall Data.xlsb]HH1 Master'!C6,MATCH(RC[2],'[Overall Data.xlsb]HH1 Master'!C9,0))", True, , , , , , "RC")
    'Call addFormulas("O2", "=VLOOKUP(RC[-5],Damaged!C[-11],1,FALSE)", False, , , , , , "RC")
    Call addFormulas("P2", "=VLOOKUP(RC[-6],'[Overall Data.xlsb]AL Sales Master'!C8,1,FALSE)", True, , , , , , "RC")
    Call addFormulas("Q2", "=VLOOKUP(RC[-7],'[Overall Data.xlsb]gDN Sales Master'!C7,1,FALSE)", True, , , , , , "RC")
    'Call addFormulas("R2", "=VLOOKUP(RC[-8],'[Overall Data.xlsb]AL Arrivals Master'!C7,1,FALSE)", True)
    Call addFormulas("S2", "=VLOOKUP(RC[-9],'Stock Spool'!C[-9],1,FALSE)", False, , , , , , "RC")
    Call addFormulas("T2", "=VLOOKUP(RC[-10],'Vista Stock Spool'!C[-19],1,FALSE)", False, , , , , , "RC") 'CHANGED'
    Call addFormulas("U2", "=VLOOKUP(RC[-11],RNI!C[-18],1,FALSE)", False, , , , , , "RC")
    
    Call clearROModelCode

'Change Header

    With Sheets("gDN Stock Spool")
        .Range("H1") = "AutoLine Based Model (HH1)"
        '.Range("O1") = "HO Units - Damaged"
        .Range("P1") = "Invoiced - Master"
        .Range("Q1") = "gDN Reported - Master"
        .Range("R1") = "Arrivals Master"
        .Range("S1") = "ATM Stock Spool"
        .Range("T1") = "VISTA Stock Spool"
        .Range("U1") = "RNI"
    End With
    
'Add Header Comments

     With Sheets("gDN Stock Spool")
        With .Range("H1")
            .Value = "AL Model"
            '.ClearComments
            '.AddComment
            '.Comment.Visible = True
            '.Comment.Text Text:="Philip Jacob:" & " Test"
        End With

        '.Range("O1") = "HO Units - Damaged"
        .Range("P1") = "Invoiced - Master"
        .Range("Q1") = "gDN Reported - Master"
        .Range("R1") = "Arrivals Master"
        .Range("S1") = "ATM Stock Spool"
        .Range("T1") = "VISTA Stock Spool"
        With .Range("U1")
            .Value = "RNI"
            .ClearComments
            .AddComment
            .Comment.Visible = False
            .Comment.Text Text:="Philip Jacob:" & vbCrLf & _
                                "Coloumn Usages:-" & vbCrLf & vbCrLf & _
                                " > RNI ideally should show only '#N/A'" & vbCrLf & _
                                "If not check relevent sheets & then delete that row from 'gDN stock Sheet'."
        End With
    End With
   
Call excelVitals(True)

End Sub

Public Sub addFormulasVISTASales()

    'CHANGED'
'Temporarily Stop Events, Calculations, Screen & Statusbar updations.
Call excelVitals(False)
    
     Call addFormulas("L2", "=INDEX('[Overall Data.xlsb]HH1 Master'!C6,MATCH(RC[-11],'[Overall Data.xlsb]HH1 Master'!C9,0))", True, , , , , , "RC")
     Call addFormulas("M2", "=VLOOKUP(RC[-12],'gDN Sales Spool'!C[-6],1,FALSE)", False, , , , , , "RC")
     'Call addFormulas("N2", "=VLOOKUP(RC[-13],'[Overall Data.xlsb]VISTA Sales Master'!C13,1,FALSE)", True)

     Call clearROModelCode

'Change Header
    
    Sheets("Vista Sales Spool").Range("L1") = "AL Based Model (HH1)"
    Sheets("Vista Sales Spool").Range("M1") = "gDN Sales Spool"
    'Sheets("Vista Sales Spool").Range("N1") = "VISTA Sales Master"
    
    Call refreshVISTASaleReportInMasterFile
    
Call excelVitals(True)

End Sub

Public Sub addFormulasVISTAStock()

'Temporarily Stop Events, Calculations, Screen & Statusbar updations.
Call excelVitals(False)
    
     Call addFormulas("L2", "=INDEX('[Overall Data.xlsb]HH1 Master'!C6,MATCH(RC[-11],'[Overall Data.xlsb]HH1 Master'!C9,0))", True, , , , , , "RC")
    
     Call clearROModelCode

'Change Header
    
    Sheets("Vista Stock Spool").Range("L1") = "AL Based Model (HH1)"
    
Call excelVitals(True)


End Sub

Public Sub refreshVISTASaleReportInMasterFile()
   
    Dim wb As Excel.Workbook
    
    Set wb = Application.Workbooks.Open("P:\LR\General Reports\" & "Overall Data.xlsb")
    
    wb.Sheets("VISTA Sale Master").Range("A1").ListObject.QueryTable.Refresh BackgroundQuery:=False
    
    wb.Save
    wb.Close
    
End Sub


Private Sub clearROModelCode()

'****************************************************************************************
'This procedure checks the model code field for Range Rover SVR vehicles.               *
'It then updates specific RO(client) Model Code field with the codes:                   *
''SCBV-SVR                                                                              *
''SDBV-SE                                                                             *
''RDBV-340PS                                                                            *
''RDBV-380PS                                                                            *
'The client report does not distinguish this variant and so is not available.           *
'However, client has provided targets for these sub-variants in their business plan.    *
'The procedure is repeated for all the 4 sheets of client data.                         *
'                                                                                       *
'Procedure Date: 5th May, 2015                                                          *
'****************************************************************************************



'Check to make sure that the code runs only when the specified file is open.
If fileCheck("LR SALES") = False Then End


'Loop Through All Worksheets in a Workbook
Dim WS_Count As Integer
Dim i As Integer
Dim checkRange As String
Dim changeRange As String
Dim iCnt As Long
Dim tCnt As Long


' Set WS_Count equal to the number of worksheets in the active
' workbook.
WS_Count = ActiveWorkbook.Worksheets.count


' Begin the loop.
'For i = 1 To WS_Count

'   If ActiveWorkbook.Worksheets(i).Name = "Vista Sales Spool" Or _
'        ActiveWorkbook.Worksheets(i).Name = "gDN Sales Spool" Or _
'        ActiveWorkbook.Worksheets(i).Name = "Vista Stock Spool" Or _
'        ActiveWorkbook.Worksheets(i).Name = "gDN Stock Spool" Then
   If ActiveWorkbook.ActiveSheet.name = "Vista Sales Spool" Or _
        ActiveWorkbook.ActiveSheet.name = "gDN Sales Spool" Or _
        ActiveWorkbook.ActiveSheet.name = "Vista Stock Spool" Or _
        ActiveWorkbook.ActiveSheet.name = "gDN Stock Spool" Then
            
        'Select Case ActiveWorkbook.Worksheets(i).Name
        Select Case ActiveWorkbook.ActiveSheet.name
            
            Case "Vista Sales Spool"    'CHANGED'
            
                checkRange = "L"
                changeRange = "B"
                
            Case "gDN Sales Spool"
            
                 checkRange = "F"
                 changeRange = "C"
                    
            Case "Vista Stock Spool"    'CHANGED'
            
                 checkRange = "L"
                 changeRange = "B"
                
            Case "gDN Stock Spool"
            
            'MsgBox ActiveWorkbook.Worksheets(I).Name
            
                 checkRange = "H"
                 changeRange = "C"
        
        End Select
        
        'Activate sheet and clear filters
        'ActiveWorkbook.Worksheets(i).Activate
        
        On Error Resume Next
        ActiveSheet.ShowAllData
        On Error GoTo 0
        
        Range(checkRange & "1").Select
        
        'Obtain total count of records
        
        tCnt = Range(checkRange & Rows.count).End(xlUp).Row
        
        Range(checkRange & "1").Select
        
        For iCnt = 2 To tCnt
            
            If Len(Range(changeRange & iCnt).Value) = 4 Then
            
                If InStr(Range(checkRange & iCnt).Value, "SVR") > 0 Then
                    
                       Range(changeRange & iCnt).Value = "SCBV-SVR"
                End If
                       
                If InStr(Range(checkRange & iCnt).Value, "Range Rover Sport 3.0 SC SE") > 0 Then
                    
                       Range(changeRange & iCnt).Value = "SDBV-SE"
                End If
                
                If InStr(Range(checkRange & iCnt).Value, "340") > 0 Then
                    
                       Range(changeRange & iCnt).Value = Range(changeRange & iCnt).Value & "-340PS"
                End If
                    
                If InStr(Range(checkRange & iCnt).Value, "380") > 0 Then
                    
                       Range(changeRange & iCnt).Value = Range(changeRange & iCnt).Value & "-380PS"
                    
                End If
                
            End If
            
        Next iCnt
        
    End If

   
   

'Next i


End Sub


Public Sub prepareROReportList(listSource As String)

    Dim fullColRange As Range, filteredCells As Range
    Dim roReportListRecordCount As Integer
    Dim repeatModelCounter As Integer

    Set fullColRange = Range("B15", Range("B15").End(xlDown)).Cells.SpecialCells(xlCellTypeVisible)
    roReportListRecordCount = totalRecords("A", "2", "RO Report List")

    If roReportListRecordCount > 0 Then

        With Sheets("RO Report List")
            .Activate
            .Range(.Range("A2"), .Range("Q" & roReportListRecordCount)).Clear   'Delete data from range A2:Qx
            .Range("D1:Q1").Clear
            .Range("A2").Select
        End With



    End If


    'Fill short model in column A
    For Each filteredCells In fullColRange

        If UCase(listSource) = "CARRY FORWARD" Then

            repeatModelCounter = filteredCells.Offset(0, 14).Value

        Else

            repeatModelCounter = filteredCells.Offset(0, 10).Value

        End If

        For Counter = 1 To repeatModelCounter

            ActiveCell = filteredCells.Value
            ActiveCell.Offset(1, 0).Activate

        Next Counter

    Next

    If UCase(listSource) = "CARRY FORWARD" Then

        salesCount = totalRecords("A", "2", "Carry Forward Units")

        Application.DisplayAlerts = False   'Suppress file open dialog
        Call addFormulas("B2", "=IFERROR(INDEX(X_C!G:G,SMALL(IF((X_C!$D$2:$D$" & salesCount & "=0)+(X_C!$F$2:$F$" & salesCount & "=$A2)+(X_C!$AH$2:$AH$" & salesCount & "<>""Export"")+(X_C!$AH$2:$AH$" & salesCount & "<>""Demo"")+('Carry Forward Units'!$AH$2:$AH$" & salesCount & "<>""Courtesy"")+(X_C!$L$2:$L$" & salesCount & "<>""In House Deals J"")=6,ROW(X_C!$G$2:$G$" & salesCount & ")),COUNTIF(A$2:A2,A2))),""No Model Found In Sales List"")", True, "RO Report List", , , , True, "A1", "'Carry Forward Units'")
        Call addFormulas("C2", "=IFERROR(INDEX(X_C!H:H,SMALL(IF((X_C!$D$2:$D$" & salesCount & "=0)+(X_C!$F$2:$F$" & salesCount & "=$A2)+(X_C!$AH$2:$AH$" & salesCount & "<>""Export"")+(X_C!$AH$2:$AH$" & salesCount & "<>""Demo"")+('Carry Forward Units'!$AH$2:$AH$" & salesCount & "<>""Courtesy"")+(X_C!$L$2:$L$" & salesCount & "<>""In House Deals J"")=6,ROW(X_C!$G$2:$G$" & salesCount & ")),COUNTIF(A$2:A2,A2))),""No Model Found In Sales List"")", True, "RO Report List", , , , True, "A1", "'Carry Forward Units'")
        Application.DisplayAlerts = True

    Else

        salesCount = totalRecords("A", "2", "Sales Detailed")

        Application.DisplayAlerts = False   'Suppress file open dialog
        Call addFormulas("B2", "=IFERROR(INDEX(X_C!G:G,SMALL(IF((11" & salesCount & "=0)+(22" & salesCount & "=$A2)+(33" & salesCount & "<>""Export"")+(44" & salesCount & "<>""Demo"")+(X_C!$AH$2:$AH$" & salesCount & "<>""Courtesy"")=5,ROW(X_C!$G$2:$G$" & salesCount & ")),COUNTIF(A$2:A2,A2))),""No Model Found In Sales List"")", True, "RO Report List", , , , True, "A1", "'Sales Detailed'", "X_C!$D$2:$D$", "X_C!$F$2:$F$", "X_C!$AH$2:$AH$", "X_C!$AH$2:$AH$")
        Call addFormulas("C2", "=IFERROR(INDEX(X_C!H:H,SMALL(IF((11" & salesCount & "=0)+(22" & salesCount & "=$A2)+(33" & salesCount & "<>""Export"")+(44" & salesCount & "<>""Demo"")+(X_C!$AH$2:$AH$" & salesCount & "<>""Courtesy"")=5,ROW(X_C!$G$2:$G$" & salesCount & ")),COUNTIF(A$2:A2,A2))),""No Model Found In Sales List"")", True, "RO Report List", , , , True, "A1", "'Sales Detailed'", "X_C!$D$2:$D$", "X_C!$F$2:$F$", "X_C!$AH$2:$AH$", "X_C!$AH$2:$AH$")
        Application.DisplayAlerts = True

    End If

    Call markDuplicates("C")

    Sheets("RO Report List").Range("A1").Select

End Sub
