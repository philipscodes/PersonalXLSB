Attribute VB_Name = "VCC_Number_Upload"

Sub VCCNumUpload()

    Range("F2").Select
    ActiveCell.FormulaR1C1 = "Location"
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Columns("A:C").Select
    Selection.Delete Shift:=xlToLeft
    Columns("D:F").Select
    Selection.Delete Shift:=xlToLeft
    Range("C2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Selection.End(xlUp).Select
    Selection.End(xlToLeft).Select
    
    ActiveCell.FormulaR1C1 = "=upper"
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "=UPPER(RC[-1])"
    Range("D2").Select
    Selection.AutoFill Destination:=Range("D2:D168")
    Range("D2:D168").Select
    Selection.Copy
    Range("C2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("D:D").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Selection.End(xlToLeft).Select
    ActiveWorkbook.Save
   
End Sub

Sub lookupVCCLocation()

    Range("C2").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-2],'[LR-ATMC-VCC.xls]LR-ATMC-VCC'!C3:C7,5,FALSE)"
    Range("C2").Select
    Selection.AutoFill Destination:=Range("C2:C168")
    Range("C2:C168").Select
    Selection.End(xlUp).Select
    Range("C2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("C2").Select
    Application.CutCopyMode = False
    Columns("C:C").Select
    Selection.Replace What:="0", Replacement:="ATMC", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="Sales", Replacement:="ATMC", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="Shj-3", Replacement:="SHJ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

End Sub


