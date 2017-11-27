Attribute VB_Name = "TBO_Functions"
Sub chuma()

    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

Public Function clrLZeros(ByVal Target As Range) As String

Dim varTemp As String

varTemp = Target.Value

Do While Left(varTemp, 1) = 0

    varTemp = Right(varTemp, Len(varTemp) - 1)

Loop

clrLZeros = varTemp

End Function

Sub TBOMY(ByVal Target As Range)

    Dim aCell As Range

    On Error GoTo Whoa

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    If Not Intersect(Target, Columns(8)) Is Nothing Then
        For Each aCell In Target
            If aCell.Value = 1 Then
                aCell.Value = "2014 MY"
            ElseIf aCell.Value = 2 Then
                aCell.Value = "2015 MY"
            End If
        Next
    End If
    
    If Not Intersect(Target, Columns(9)) Is Nothing Then
        For Each aCell In Target
            If UCase(aCell.Value) = "T" Then
                aCell.Value = Date
            ElseIf aCell.Value = 1 Then
                aCell.Value = Date - 1
            ElseIf aCell.Value = 2 Then
                aCell.Value = Date - 2
            ElseIf aCell.Value = 3 Then
                aCell.Value = Date - 3
            ElseIf aCell.Value = 4 Then
                aCell.Value = Date - 4
            End If
        Next
    End If

Letscontinue:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    
    Exit Sub
Whoa:
    MsgBox Err.description
    Resume Letscontinue

End Sub

Sub Test_InsertCode()

    Dim Commands As String
    Commands = Chr(13) & _
                "Private Sub Worksheet_Change(ByVal Target As Range)" & Chr(13) & _
                "Application.Run ""Personal.xlsb!TBOMY"", Target" & Chr(13) & _
                "End Sub"

    Dim VBComps As VBComponents
    Dim VBComp As VBComponent
    Dim VBCodeMod As CodeModule

    Dim oSheet As Worksheet
    
    rsult = InputBox("L405 - 1, L405 - 2", "Choose Model")
    
    If rsult = 1 Then
        Set VBComps = Workbooks("L405 TBO- 2014MY - Philip.xlsm").VBProject.VBComponents
        Set oSheet = Workbooks("L405 TBO- 2014MY - Philip.xlsm").Worksheets("L405")
    ElseIf rsult = 2 Then
        Set VBComps = Workbooks("L494 Bookings Philip.xlsm").VBProject.VBComponents
        Set oSheet = Workbooks("L494 Bookings Philip.xlsm").Worksheets("L494")
    End If

    
    
   
    'For Each oSheet In Workbooks("L405 TBO- 2014MY - Philip.xlsm").Worksheets
        Set VBComp = VBComps(oSheet.CodeName)
        Set VBCodeMod = VBComp.CodeModule
        InsertCode VBCodeMod, Commands
    'Next oSheet

    'Here's a quick example of how to insert code in a new Module
    'Set VBComp = VBComps.Add(vbext_ct_StdModule)
    'InsertCode VBComp.CodeModule, Commands

End Sub

Private Sub InsertCode(VBCodeMod As CodeModule, Commands As String)

    Dim LineNum As Long
    With VBCodeMod
        LineNum = .CountOfLines + 1
        .InsertLines LineNum, Commands
    End With

End Sub

Sub temp()

temp1 = "Semi-Aniline Leather Rear Seat Style 19 (Bench st)"
temp1 = Trim(StrConv(temp1, vbProperCase))

MsgBox temp1, vbOKOnly

End Sub
Public Sub optionEntry()

Dim temp As Variant


temp = InputBox("Enter Option", "Option Placer")
temp = Trim(StrConv(temp, vbProperCase))

Select Case temp

    Case "Ebony Macassar"
        
        Range("Q" & ActiveCell.Row).Value = temp
    
    Case "Shadow Walnut Veneer"
        
        Range("Q" & ActiveCell.Row).Value = "Shadow Walnut"
        
    Case "Shadow Zebrano"
        
        Range("Q" & ActiveCell.Row).Value = "Shadow Zebrano"
    
    Case "Grand Black Wood"
        
            Range("Q" & ActiveCell.Row).Value = temp
        
    Case "Grand Black"
        
            Range("Q" & ActiveCell.Row).Value = "Grand Black"
        
    Case "Micro Mesh Aluminium"
        
            Range("Q" & ActiveCell.Row).Value = "Micro Mesh Aluminium"
        
    Case "Sports Textured Aluminium"
        
            Range("Q" & ActiveCell.Row).Value = "Sports Textured Aluminium"
    
    Case "Dark Grey Oak"
        
            Range("Q" & ActiveCell.Row).Value = "Dark Grey Oak"
    
    Case "21" & Chr(34) & " Alloy Wheels Style 5 (2015my)"
        
        If Left(ActiveWorkbook.name, 4) = "L405" Then
            Range("S" & ActiveCell.Row).Value = "21" & Chr(34) & " Style 5"
        Else
            Range("R" & ActiveCell.Row).Value = "21" & Chr(34) & " Style 5"
        End If
        
    Case "21" & Chr(34) & " Alloy Wheels Style 5 (2014my)"
        
        If Left(ActiveWorkbook.name, 4) = "L405" Then
            Range("S" & ActiveCell.Row).Value = "21" & Chr(34) & " Style 5"
        Else
            Range("R" & ActiveCell.Row).Value = "21" & Chr(34) & " Style 5"
        End If
        
    Case "21" & Chr(34) & " Alloy Wheels Style 5"
        
        If Left(ActiveWorkbook.name, 4) = "L405" Then
            Range("S" & ActiveCell.Row).Value = "21" & Chr(34) & " Style 5"
        Else
            Range("R" & ActiveCell.Row).Value = "21" & Chr(34) & " Style 5"
        End If
        
    Case "22" & Chr(34) & " Alloy Wheels Style 6 (2015my)"
        
        If Left(ActiveWorkbook.name, 4) = "L405" Then
            Range("S" & ActiveCell.Row).Value = "22" & Chr(34) & " Style 6"
        Else
            Range("R" & ActiveCell.Row).Value = "22" & Chr(34) & " Style 6"
        End If
    
    Case "22" & Chr(34) & " Alloy Wheels Style 6 (2014my)"
        
        If Left(ActiveWorkbook.name, 4) = "L405" Then
            Range("S" & ActiveCell.Row).Value = "22" & Chr(34) & " Style 6"
        Else
            Range("R" & ActiveCell.Row).Value = "22" & Chr(34) & " Style 6"
        End If
        
    Case "20" & Chr(34) & " Alloy Wheels Style 2 (2015my)"
        
        If Left(ActiveWorkbook.name, 4) = "L405" Then
            Range("S" & ActiveCell.Row).Value = "20" & Chr(34) & " Style 2"
        Else
            Range("R" & ActiveCell.Row).Value = "20" & Chr(34) & " Style 2"
        End If
        
        
    'Case "20" & Chr(34) & " Alloy Wheels Style 2 (2015MY)"
        'Range("S" & ActiveCell.Row).Value = "21 Style 5"

    Case "21" & Chr(34) & " Alloy Style 15"
        If Left(ActiveWorkbook.name, 4) = "L405" Then
            Range("S" & ActiveCell.Row).Value = "21" & Chr(34) & "  Style 15"
        Else
            Range("R" & ActiveCell.Row).Value = "21" & Chr(34) & " Style 15"
        End If
    
    Case "21" & Chr(34) & " Alloy Style 4"
        If Left(ActiveWorkbook.name, 4) = "L405" Then
            Range("S" & ActiveCell.Row).Value = "21" & Chr(34) & " Style 4"
        Else
            Range("R" & ActiveCell.Row).Value = "21" & Chr(34) & " Style 4"
        End If
        
    Case "21" & Chr(34) & " Alloy Wheels Style 4 (2015my)"
        
        If Left(ActiveWorkbook.name, 4) = "L405" Then
            Range("S" & ActiveCell.Row).Value = "21" & Chr(34) & " Style 4"
        Else
            Range("R" & ActiveCell.Row).Value = "21" & Chr(34) & " Style 4"
        End If
        
    Case "21" & Chr(34) & " Alloy Style 14"
        If Left(ActiveWorkbook.name, 4) = "L405" Then
            Range("S" & ActiveCell.Row).Value = "21" & Chr(34) & " Style 14"
        Else
            Range("R" & ActiveCell.Row).Value = "21" & Chr(34) & " Style 14"
        End If
        
    Case "22" & Chr(34) & " Alloy Style 6"
        If Left(ActiveWorkbook.name, 4) = "L405" Then
            Range("S" & ActiveCell.Row).Value = "22" & Chr(34) & " Style 6"
        Else
            Range("R" & ActiveCell.Row).Value = "22" & Chr(34) & " Style 6"
        End If
        
    Case "Wood/leather Steering Wheel (2014my)"
        
        Range("AA" & ActiveCell.Row).Value = "Yes"
        
    Case "Wood/leather Steering Wheel"
        
        Range("AA" & ActiveCell.Row).Value = "Yes"
        
    Case "Atlas Side Vent Graphics"
        
        Range("U" & ActiveCell.Row).Value = "Yes"
        
    Case "Atlas Side Vent"
        
        Range("U" & ActiveCell.Row).Value = "Yes"
    
    Case "Contrast Roof (santorini Black)"
               
            Range("T" & ActiveCell.Row).Value = "Santorini Black"
        
    Case "Contrast Roof - Santorini Black"
               
            Range("S" & ActiveCell.Row).Value = "Santorini Black"

    Case "Contrast Roof (indus Silver)"
        
        Range("T" & ActiveCell.Row).Value = "Indus Silver"

    Case "Interior Mood Lighting (ambience Lighting 3)"
        
        If Range("AC" & ActiveCell.Row).Value = "" Then
            Range("AC" & ActiveCell.Row).Value = temp
        Else
            Range("AC" & ActiveCell.Row).Value = Range("AC" & ActiveCell.Row).Value & ", " & temp
        End If
       
    Case "Perpendicular Park Assist"
        
        If Range("AC" & ActiveCell.Row).Value = "" Then
            Range("AC" & ActiveCell.Row).Value = temp
        Else
            Range("AC" & ActiveCell.Row).Value = Range("AC" & ActiveCell.Row).Value & ", " & temp
        End If
        
    Case "Park Assist"
    
        If Left(ActiveWorkbook.name, 4) = "L405" Then
        
            If Range("AC" & ActiveCell.Row).Value = "" Then
                Range("AC" & ActiveCell.Row).Value = temp
            Else
                Range("AC" & ActiveCell.Row).Value = Range("AC" & ActiveCell.Row).Value & ", " & temp
            End If
        Else
                Range("U" & ActiveCell.Row).Value = temp
                
        End If
        
    Case "Dual View Touchscreen"
        
        If Left(ActiveWorkbook.name, 4) = "L405" Then
            If Range("AC" & ActiveCell.Row).Value = "" Then
                Range("AC" & ActiveCell.Row).Value = temp
            Else
                Range("AC" & ActiveCell.Row).Value = Range("AC" & ActiveCell.Row).Value & ", " & temp
            End If
        Else
                Range("T" & ActiveCell.Row).Value = temp
            
        End If
        
    Case "Dark Atlas Accents Pack - 2"
        
        If Left(ActiveWorkbook.name, 4) = "L405" Then
            If Range("AC" & ActiveCell.Row).Value = "" Then
                Range("AC" & ActiveCell.Row).Value = temp
            Else
                Range("AC" & ActiveCell.Row).Value = Range("AC" & ActiveCell.Row).Value & ", " & temp
            End If
        End If
        
    Case "Dark Atlas Pack-1(dark Atlas S/vent & Body Graph)"
        
        If Left(ActiveWorkbook.name, 4) = "L405" Then
            If Range("AC" & ActiveCell.Row).Value = "" Then
                Range("AC" & ActiveCell.Row).Value = temp
            Else
                Range("AC" & ActiveCell.Row).Value = Range("AC" & ActiveCell.Row).Value & ", " & "Dark Atlas Pack - 1"
            End If
        End If
    
    Case "Semi-aniline Leather Rear Seat Style 19 (bench St)"
        If Left(ActiveWorkbook.name, 4) = "L405" Then
                Range("X" & ActiveCell.Row).Value = "Style 19"
        End If
    
    Case "Rear Seat Entertaintment"
        If Left(ActiveWorkbook.name, 4) = "L405" Then
        
            Range("W" & ActiveCell.Row).Value = "Yes"
        
        Else
        
            Range("X" & ActiveCell.Row).Value = "Yes"
            
        End If
        
    Case "Surround Camera System"
        If Left(ActiveWorkbook.name, 4) = "L494" Then
                Range("W" & ActiveCell.Row).Value = "Yes"
        End If
        
    Case "Adaptive Cruise Control & Queue Assist"
        If Left(ActiveWorkbook.name, 4) = "L494" Then
                Range("Y" & ActiveCell.Row).Value = "Yes"
        End If
        
    Case "Meridian Signature Reference (1700w)"
        If Left(ActiveWorkbook.name, 4) = "L494" Then
                Range("Z" & ActiveCell.Row).Value = "Yes"
        End If
        
    Case "7 Seats"
        If Left(ActiveWorkbook.name, 4) = "L494" Then
            If Range("AA" & ActiveCell.Row).Value = "" Then
                Range("AA" & ActiveCell.Row).Value = temp
            Else
                Range("AA" & ActiveCell.Row).Value = Range("AC" & ActiveCell.Row).Value & ", " & temp
            End If
        End If
    
    Case Else
        If Left(ActiveWorkbook.name, 4) = "L405" Then
            If Range("AC" & ActiveCell.Row).Value = "" Then
                Range("AC" & ActiveCell.Row).Value = temp
            Else
                Range("AC" & ActiveCell.Row).Value = Range("AC" & ActiveCell.Row).Value & ", " & temp
            End If
        ElseIf Left(ActiveWorkbook.name, 4) = "L494" Then
            If Range("AA" & ActiveCell.Row).Value = "" Then
                Range("AA" & ActiveCell.Row).Value = temp
            Else
                Range("AA" & ActiveCell.Row).Value = Range("AA" & ActiveCell.Row).Value & ", " & temp
            End If
        
        End If
    

End Select

End Sub
