Attribute VB_Name = "VBE_Functions"
Public Sub CountLines()

Dim vbProj As VBProject
Dim VBComp As VBComponent
Dim codeLines As String


For Each VBComp In ThisWorkbook.VBProject.VBComponents

    codeLines = codeLines & VBComp.name & " = " & VBComp.CodeModule.CountOfLines & " code lines." & vbCrLf
Next VBComp

    MsgBox codeLines


End Sub

'---------------------------------------------------------------------------------------
' Procedure : GetVBEDeatils
' Author    : CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Goes throught the VBE and creates a text file which give a brief listing
'             of the procedures within each module and a line count for each
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
' Requirements: reference to the Microsoft Visual Basic for Application Extensibility
'               library.
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2011-June-04            Initial Release
'---------------------------------------------------------------------------------------
Public Function GetVBEDeatils()
    Dim vbProj          As VBProject
    Dim VBComp          As VBComponent
    Dim vbMod           As CodeModule
    Dim sProcName       As String
    Dim pk              As vbext_ProcKind
    Dim FileNumber      As Integer
    Dim strFile         As String
    Const vbNormalFocus = 1
 
    'Where do youwant the text file created
    strFile = "C:\Users\pjacob\Documents\VBEDetails.txt"
    If Len(Dir(strFile)) > 0 Then Kill strFile
    FileNumber = FreeFile                           'Get unused file number.
    Open strFile For Append As #FileNumber          'Create file name.
        
    For Each vbProj In Application.VBE.VBProjects   'Loop through each project
        Print #FileNumber, vbProj.name
        For Each VBComp In vbProj.VBComponents      'Loop through each module
            Set vbMod = VBComp.CodeModule
            Print #FileNumber, "   " & VBComp.name & " :: " & vbMod.CountOfLines & " total lines"
            Print #FileNumber, "   " & String(80, "*")
            iCounter = 1
            Do While iCounter < vbMod.CountOfLines  'Loop through each procedure
                sProcName = vbMod.ProcOfLine(iCounter, pk)
                If sProcName <> "" Then
                    Print #FileNumber, "      " & sProcName & " :: " & vbMod.ProcCountLines(sProcName, pk) & " lines"
                    iCounter = iCounter + vbMod.ProcCountLines(sProcName, pk)
                Else
                    iCounter = iCounter + 1
                End If
            Loop
            Print #FileNumber, ""
        Next VBComp
    Next vbProj
 
    Close #FileNumber                               'Close file.
    Set vbMod = Nothing
 
    'Open the generated text file
    Shell "cmd /c """ & strFile & """", vbNormalFocus
End Function
