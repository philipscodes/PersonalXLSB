Attribute VB_Name = "Special_Codes"
Sub delete_empty_cells_shift_left()
    Range("A1:U255").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.Delete Shift:=xlToLeft
End Sub
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Cells.count > 1 Then Exit Sub
    Application.ScreenUpdating = False
    ' Clear the color of all the cells
    Cells.Interior.ColorIndex = 0
    With Target
        ' Highlight the entire row and column that contain the active cell
        .EntireRow.Interior.ColorIndex = 8
        .EntireColumn.Interior.ColorIndex = 8
    End With
    Application.ScreenUpdating = True
End Sub

