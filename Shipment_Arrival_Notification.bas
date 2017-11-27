Attribute VB_Name = "Shipment_Arrival_Notification"
Sub shipmentArrivalNotification()
Attribute shipmentArrivalNotification.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro6 Macro
'

'
    Columns("E:E").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("E7").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(""SAL"",RC[-2],RC[-1])"
    Range("E7").Select
    Selection.AutoFill Destination:=Range("E7:E461")
    Range("E7:E461").Select
    Columns("E:E").EntireColumn.AutoFit
    Range("D6").Select
    Selection.Copy
    Range("E6").Select
    ActiveSheet.Paste
    Columns("C:D").Select
    Application.CutCopyMode = False
    Range("E7").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("C:D").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Columns("E:E").Select
    Selection.Delete Shift:=xlToLeft
    Columns("G:G").Select
    Selection.Delete Shift:=xlToLeft
    
End Sub
