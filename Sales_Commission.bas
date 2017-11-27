Attribute VB_Name = "Sales_Commission"



Public Sub prepareJLRCommission()


    Dim franchise As String
    

    For i = 1 To 2
    
        Select Case i
        
            Case 1
            
                franchise = "LANDROVER"
                Sheets("LR SALES DETAILS").Activate
                
            Case 2
            
                franchise = "JAGUAR"
                Sheets("JAG SALES DETAILS").Activate
                
        End Select
        
     
        Call executeSalesCommissionQuery(franchise)
        
        Call subtotalRange
        
        Call reformatToTable(franchise)
            
        Call insertSalesAdvisorName(franchise)
        
        Call updateDetailSheetColoumnFormat(franchise)
        
        ActiveWindow.Zoom = True
        'Cells.EntireColumn.AutoFit
        ActiveWindow.Zoom = 75
        
        Call addSummarySheetFormulas(franchise)
    
    Next i
  
End Sub

Private Sub executeSalesCommissionQuery(franchise As String)


    Dim salesDataQuery As String
    

    Select Case franchise
        
        Case "LANDROVER"
        
        salesDataQuery = "" & _
        "SELECT qry_AN6_PrevMonth_LR_SalesCommission.Loc, qry_AN6_PrevMonth_LR_SalesCommission.Main_Company, qry_AN6_PrevMonth_LR_SalesCommission.INV_No, qry_AN6_PrevMonth_LR_SalesCommission.INV_Date, qry_AN6_" & _
        "PrevMonth_LR_SalesCommission.VSB, qry_AN6_PrevMonth_LR_SalesCommission.MY, qry_AN6_PrevMonth_LR_SalesCommission.Description, qry_AN6_PrevMonth_LR_SalesCommission.Chassis, qry_AN6_PrevMonth_LR_SalesCom" & _
        "mission.Customer_Name, qry_AN6_PrevMonth_LR_SalesCommission.Sales_Executive, qry_AN6_PrevMonth_LR_SalesCommission.Sale_Type, qry_AN6_PrevMonth_LR_SalesCommission.Normal, qry_AN6_PrevMonth_LR_SalesComm" & _
        "ission.Promotions, qry_AN6_PrevMonth_LR_SalesCommission.`Internal_&_Others`, qry_AN6_PrevMonth_LR_SalesCommission.Total" & Chr(13) & "" & Chr(10) & "FROM `P:\LR\General Reports\ATMDB.accdb`.qry_AN6_PrevMonth_LR_SalesCommission q" & _
        "ry_AN6_PrevMonth_LR_SalesCommission" & Chr(13) & "" & Chr(10) & "ORDER BY qry_AN6_PrevMonth_LR_SalesCommission.Main_Company, qry_AN6_PrevMonth_LR_SalesCommission.Sales_Executive"

        Case "JAGUAR"
        
        salesDataQuery = "" & _
            "SELECT qry_AN6_PrevMonth_JAG_SalesCommission.Loc, qry_AN6_PrevMonth_JAG_SalesCommission.Main_Company, qry_AN6_PrevMonth_JAG_SalesCommission.INV_No, qry_AN6_PrevMonth_JAG_SalesCommission.INV_Date, qry_" & _
            "AN6_PrevMonth_JAG_SalesCommission.VSB, qry_AN6_PrevMonth_JAG_SalesCommission.MY, qry_AN6_PrevMonth_JAG_SalesCommission.Description, qry_AN6_PrevMonth_JAG_SalesCommission.Chassis, qry_AN6_PrevMonth_JAG" & _
            "_SalesCommission.Customer_Name, qry_AN6_PrevMonth_JAG_SalesCommission.Sales_Executive, qry_AN6_PrevMonth_JAG_SalesCommission.Sale_Type, qry_AN6_PrevMonth_JAG_SalesCommission.Normal, qry_AN6_PrevMonth_" & _
            "JAG_SalesCommission.Promotions, qry_AN6_PrevMonth_JAG_SalesCommission.`Internal_&_Others`, qry_AN6_PrevMonth_JAG_SalesCommission.Total" & Chr(13) & "" & Chr(10) & "FROM `P:\LR\General Reports\ATMDB.accdb`.qry_AN6_PrevMonth_JAG_S" & _
            "alesCommission qry_AN6_PrevMonth_JAG_SalesCommission" & Chr(13) & "" & Chr(10) & "ORDER BY qry_AN6_PrevMonth_JAG_SalesCommission.Main_Company, qry_AN6_PrevMonth_JAG_SalesCommission.Sales_Executive"
            
    End Select
    

    With ActiveSheet.ListObjects.Add(SourceType:=0, source:=Array(Array( _
        "ODBC;DSN=MS Access Database;DBQ=P:\LR\General Reports\ATMDB.accdb;DefaultDir=P:\LR\General Reports;DriverId=25;FIL=MS Access;MaxBuff" _
        ), Array("erSize=2048;PageTimeout=15;")), Destination:=Range("$A$4")).QueryTable _

        .CommandType = xlCmdSql
        .CommandText = salesDataQuery
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "Jaguar_Sales_Commission_Detail_Table"
        .Refresh BackgroundQuery:=False
        
    End With
    
    'ActiveSheet.ListObjects(1).ShowTableStyleColumnStripes = True
    
    Range("A4").Select
    
End Sub



Private Sub subtotalRange()

    ActiveSheet.ListObjects(1).Unlist       'Convert to range
    
    'Subtotal division of ATM & PM deals
    Range("A4").CurrentRegion.Subtotal GroupBy:=2, Function:=xlCount, TotalList:=Array(2), _
        Replace:=True, PageBreaks:=False, SummaryBelowData:=True
            
    'Subtotal division of Sales Advisor with summation of 'Normal', 'Promotions', 'Int & Others', 'Total' fields
    Range("A4").CurrentRegion.Subtotal GroupBy:=10, Function:=xlSum, TotalList:=Array(12, 13 _
        , 14, 15), Replace:=False, PageBreaks:=False, SummaryBelowData:=True

End Sub

Private Sub reformatToTable(franchise As String)

    Dim tblSource As Range
    Dim lo As Variant
    Dim tableName As String
    
    Select Case franchise
    
        Case "JAGUAR"
            tableName = "Jaguar_Sales_Commission_Detail_Table"
            
        Case "LANDROVER"
            tableName = "LR_Sales_Commission_Detail_Table"
            
    End Select

    Set tblSource = Range("A4").CurrentRegion
    
    
    Set lo = ActiveSheet.ListObjects.Add(SourceType:=xlSrcRange, _
                                    source:=tblSource, _
                                    xlListObjectHasHeaders:=xlYes, _
                                    TableStyleName:="TableStyleLight9")
                                    
    ActiveSheet.ListObjects(1).name = tableName
    

End Sub



Private Sub insertSalesAdvisorName(franchise)


    Dim cnt As Integer
    Dim tableName As String
    
    
    
    Select Case franchise
    
        Case "LANDROVER"
            
            tableName = "LR_Sales_Commission_Detail_Table"
    
        Case "JAGUAR"
            
            tableName = "Jaguar_Sales_Commission_Detail_Table"
            
    End Select
    

    Range("H4").Select
    
    
    
    Do While 1
        
        On Error GoTo exitRoutine
        If ActiveCell.ListObject.name = tableName Then
            
    
            If IsEmpty(ActiveCell.Value) = True And IsEmpty(ActiveCell.Offset(0, 2).Value) = False Then
                
                Call addFormulas(ActiveCell.Address, "=LEFT([@[Sales_Executive]],LEN([@[Sales_Executive]])-6)", True, , , , True, , "A1", , , , , , True)
                ActiveCell.Font.Color = ActiveCell.Interior.Color       'Equate interior colour with font to hide the content
                Call addFormulas(ActiveCell.Offset(0, 7).Address, "=SUM(" & tableName & "[@[Normal]:[Internal_&_Others]])", False, , , , True, , "A1", , , , , , True)
                
                For cnt = 4 To 7
                    ActiveCell.Offset(0, cnt).Font.Bold = True
                Next cnt
                
            End If
        
        
            ActiveCell.Offset(1, 0).Select
            
        End If
            
        
    Loop

exitRoutine:

End Sub


Private Sub updateDetailSheetColoumnFormat(franchise As String)

    Dim detailTableName As String
    
    Select Case franchise
    
        Case "LANDROVER"
        
                detailTableName = "LR_Sales_Commission_Detail_Table"
   
        Case "JAGUAR"
        
                detailTableName = "Jaguar_Sales_Commission_Detail_Table"
                
    End Select

    Range(detailTableName & "[[#All],[INV_No]]").NumberFormat = "General"
    
    Range(detailTableName & "[[#All],[INV_Date]]").NumberFormat = "m/d/yyyy"
    
    Range(detailTableName & "[[#All],[VSB]:[MY]]").NumberFormat = "General"
    
    Range(detailTableName & "[[#All],[Normal]:[Total]]").NumberFormat = "#,##0"
    
End Sub

Private Sub addSummarySheetFormulas(franchise)


    Dim summarySheetName As String
    Dim summaryTableName As String
    Dim detailTableName As String
    Dim supportSANameTable As String
    
    Select Case franchise
    
        Case "LANDROVER"
        
                summarySheetName = "LR NC"
                summaryTableName = "LR_Summary_Table"
                detailTableName = "LR_Sales_Commission_Detail_Table"
                supportSANameTable = "LR_SA_Name_Support_Table"
   
        Case "JAGUAR"
        
                summarySheetName = "Jaguar NC"
                summaryTableName = "Jaguar_Summary_Table"
                detailTableName = "Jaguar_Sales_Commission_Detail_Table"
                supportSANameTable = "SA_Name_Support_Table"
                
    End Select
    
    Sheets(summarySheetName).Activate

    Range(summaryTableName & "[Acheived]").FormulaR1C1 = _
        "=COUNTIFS(" & detailTableName & "[[#All],[Sales_Executive]]," & supportSANameTable & "[@SALES])"
        
    Range(summaryTableName & "[Target]").FormulaR1C1 = _
        "=" & summaryTableName & "[Acheived]"
        
    Range(summaryTableName & "[Sales]").FormulaR1C1 = _
        "=SUMIFS(" & detailTableName & "[[#All],[Total]]," & detailTableName & "[[#All],[Chassis]]," & supportSANameTable & "[[#All],[SALES]])"
        
    Range(summaryTableName & "[Total]").FormulaR1C1 = _
        "=SUM(" & summaryTableName & "[@[Sales]:[ASAP]])"
        
    Range(summaryTableName & "[Line Total]").FormulaR1C1 = _
        "=" & summaryTableName & "[@Total]-" & summaryTableName & "[@[Performance 30%]]-" & summaryTableName & "[@[Sales Data 10%]]-" & summaryTableName & "[@[Demo 10%]]-" & summaryTableName & "[@[CI / MS 10%]]+" & summaryTableName & "[@[Excel 20%]]"
        

End Sub

