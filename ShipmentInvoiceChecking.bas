Attribute VB_Name = "ShipmentInvoiceChecking"

Dim sheetNums As Integer
Dim invoiceNum, shipmentName As String


Sub dialogSheets()

sheetNums = InputBox("Enter number of sheets to process", "Sheets to Process")
shipmentName = InputBox("Enter Shipment Name", "Shipment Name")
invoiceNum = InputBox("Enter Invoice Number", "Invoice Number")

End Sub

Sub activateRangeA()
    
    Windows(shipmentName & " " & invoiceNum & ".xlsx").Activate
    For cnt = 1 To sheetNums
    
        Sheets(CStr(cnt)).Activate
        Range("A3").Select
    
    Next

End Sub

Function chartFinder(findThis As String) As Boolean

Dim sht As Worksheet
Dim found As Range

 For Each sht In Workbooks("Price Chart 2013 2014.xls").Worksheets

    Set found = sht.Cells.Find(What:=findThis, After:= _
        ActiveCell, LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
    
    If Not found Is Nothing Then Exit For
    
Next sht

    If Not found Is Nothing Then
        
        sht.Activate
        found.Select
        chartFinder = True
    Else
        chartFinder = False
    End If

End Function

Function chartOptionFinder(findThis As String) As Boolean

    Windows(shipmentName & " " & invoiceNum & ".xlsx").Activate
    
    Set found = Cells.Find(What:=findThis, After:= _
        ActiveCell, LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
        
    If Not found Is Nothing Then
        
        found.Select
        chartOptionFinder = True
    
    Else
        
        chartOptionFinder = False
        
    End If


End Function

Sub lessItem(invoiceValue As Long)

    Windows(shipmentName & " " & invoiceNum & ".xlsx").Activate
    'Inserting and highlighting the difference.
    ActiveCell.Offset(1, 0).Rows.EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveSheet.Cells(ActiveCell.Row, 1).Select
    ActiveCell.Offset(0, 1).Value = "Less"
    ActiveCell.Offset(0, 2).Value = invoiceValue
    Windows("Price Chart 2013 2014.xls").Activate
    ActiveCell.Offset(1, 0).Select
            
            
End Sub

Sub addItemDifference(priceChartValue As Long, invoiceValue As Long)

    Windows(shipmentName & " " & invoiceNum & ".xlsx").Activate
    'Inserting and highlighting the difference.
    ActiveCell.Offset(1, 0).Rows.EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveSheet.Cells(ActiveCell.Row, 1).Select
    ActiveCell.Offset(0, 1).Value = "Add"
    ActiveCell.Offset(0, 2).Value = priceChartValue - invoiceValue
    Windows("Price Chart 2013 2014.xls").Activate
    ActiveCell.Offset(1, 0).Select


End Sub

Sub lessItemDifference(priceChartValue As Long, invoiceValue As Long)

    Windows(shipmentName & " " & invoiceNum & ".xlsx").Activate
    'Inserting and highlighting the difference.
    ActiveCell.Offset(1, 0).Rows.EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveSheet.Cells(ActiveCell.Row, 1).Select
    ActiveCell.Offset(0, 1).Value = "Less"
    ActiveCell.Offset(0, 2).Value = invoiceValue - priceChartValue
    Windows("Price Chart 2013 2014.xls").Activate
    ActiveCell.Offset(1, 0).Select


End Sub

Sub priceChecker()

Dim priceChartValue As Long
Dim invoiceValue As Long
Dim findThis As String

Do While Not IsEmpty(ActiveCell)

    findThis = ActiveCell.Value
    priceChartValue = ActiveCell.Offset(0, 1).Value

    If ActiveCell.Offset(0, 1) = 0 Then
    
        If ActiveCell.Offset(0, 2) <> "AO" Then
            If chartOptionFinder(ActiveCell.Value) = True Then
                Windows(shipmentName & " " & invoiceNum & ".xlsx").Activate
                invoiceValue = ActiveCell.Offset(0, 1).Value
                Call lessItem(invoiceValue)
                GoTo continueLoop
            End If
        End If
            
        
    End If
    
    Windows("Price Chart 2013 2014.xls").Activate
        
    If ActiveCell.Offset(0, 1) > 0 Then
    
            If chartOptionFinder(findThis) = True Then
               
                invoiceValue = ActiveCell.Offset(0, 2).Value
    
                If priceChartValue = invoiceValue Then
                    Windows(shipmentName & " " & invoiceNum & ".xlsx").Activate
                    ActiveCell.Offset(0, 4).FormulaR1C1 = ChrW(&H2713)
                    Windows("Price Chart 2013 2014.xls").Activate
                    ActiveCell.Offset(1, 0).Select
                ElseIf invoiceValue < priceChartValue Then
                    Call addItemDifference(priceChartValue, invoiceValue)
                    GoTo continueLoop
                ElseIf invoiceValue > priceChartValue Then
                    Call lessItemDifference(priceChartValue, invoiceValue)
                    GoTo continueLoop
                End If
            
            Else
            
                Windows("Price Chart 2013 2014.xls").Activate
            
                If ActiveCell.Offset(0, 2).Value <> "AO" Then
                
                    Windows(shipmentName & " " & invoiceNum & ".xlsx").Activate
                    'Adding price option that is not found.
                    Range("A1").Select
                    Selection.End(xlDown).Select
                    ActiveCell.Offset(1, 0).Select
                    ActiveSheet.Cells(ActiveCell.Row, 1).Select
                    ActiveCell.Value = findThis
                    ActiveCell.Offset(0, 1).Value = "Add"
                    ActiveCell.Offset(0, 2).Value = priceChartValue
                    Windows("Price Chart 2013 2014.xls").Activate
                    ActiveCell.Offset(1, 0).Select
                
                Else
                
                    Windows("Price Chart 2013 2014.xls").Activate
                    ActiveCell.Offset(1, 0).Select
                
                End If

            End If
    Else
    
            Windows("Price Chart 2013 2014.xls").Activate
            ActiveCell.Offset(1, 0).Select
    
        
    End If
    
continueLoop:
Loop

End Sub

Static Sub checkInvoice()

Dim findThis As String

Call dialogSheets
Call activateRangeA

For cnt = 1 To sheetNums

    Windows(shipmentName & " " & invoiceNum & ".xlsx").Activate
    Sheets(CStr(cnt)).Activate
    findThis = ActiveCell.Value
    Windows("Price Chart 2013 2014.xls").Activate
    
    If chartFinder(findThis) = True Then
        Call priceChecker
    Else
        MsgBox "Sorry this invoice can't be checked as the corresponding model can't be found in the price chart.", vbOKOnly, "Model Not Found - Price Chart"
    End If
    
Next

Windows(shipmentName & " " & invoiceNum & ".xlsx").Activate
Sheets("1").Activate

End Sub

