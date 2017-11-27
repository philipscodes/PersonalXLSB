Attribute VB_Name = "gDN_Data_Functions"
Option Explicit


Public Function GetNamePart(NameStr As String, NamePart As String, Optional UseMiddle As Boolean = True, _
    Optional FirstMiddleLast As Boolean = True)
    
    ' Function by Patrick Matthews
    
    ' This code may be freely used or distributed so long as you acknowledge authorship and cite the URL
    ' where you found it
    
    ' This function relies on Regular Expressions.  For more information on RegExp, please see:
    ' http://www.experts-exchange.com/articles/Programming/Languages/Visual_Basic/Parsing-Names-in-MS-Office-Visual-Basic-6-and-Visual-Basic-for-Applications.html
    
    ' This function parses a name string, and depending on the arguments passed to it it returns a
    ' title, first name, middle name, surname, or suffix.
    
    ' This function can be used in any VBA or VB6 project.  It can also be used directly in Excel worksheet
    ' formulae and in Access queries
    
    ' The function takes the following arguments:
    '
    ' 1) NameStr is the full name to be parsed.  Its assumed structure is determined by the FirstMiddleLast
    '       argument
    ' 2) NamePart indicates the portion of the name you want returned.  Valid values (NOT case sensitive):
    '       Title: "TITLE", "HONORIFIC", "T", "H", "1"
    '       First: "FIRST NAME", "FIRSTNAME", "FNAME", "F NAME", "FIRST", "F", "FN", "F N", "2"
    '       Middle: "MIDDLE NAME", "MIDDLENAME", "MNAME", "M NAME", "MIDDLE", "M", "MN", "M N", "3"
    '       Last: "LAST NAME", "LASTNAME", "LNAME", "L NAME", "LAST", "L", "LN", "L N", "SURNAME", "4"
    '       Suffix: "SUFFIX", "5"
    ' 3) UseMiddle indicates whether or not a middle name *might* be present in the NameStr.  If True or
    '       omitted, the function assumes that a middle name *might* be there.  If False, it assumes there
    '       is never a middle name
    ' 4) FirstMiddleLast indicates the order of the name parts.  If True or omitted, the function assumes:
    '       <Title (optional)> <First name> <Middle Name (optional)> <Surname> <Suffix (optional)>
    '       If False, the function assumes:
    '       <Surname> <Suffix (optional)>, <First name> <Middle Name (optional)>
    
    ' Notes:
    ' 1) The function has a prebuilt list of Titles (see GenerateLists function below), which you can modify to
    '       fit your needs.  The Titles string will be embedded in a RegExp Pattern string, so be sure to follow
    '       proper RegExp Pattern syntax
    ' 2) The function will recognize compound titles, as long as they are delimited by spaces
    ' 3) The function has a prebuilt list of Suffixes (see GenerateLists function below), which you can modify to
    '       fit your needs.  The Suffixes string will be embedded in a RegExp Pattern string, so be sure to
    '       follow proper RegExp Pattern syntax
    ' 4) The function will recognize compound suffixes, as long as they are delimited by commas and/or spaces
    ' 5) If you are using True (or omitting) for FirstMiddleLast:
    '       a) It is always assumed that the first name has a single "word"
    '       b) It is always assumed that the middle name, if present, has a single "word"
    '       c) After the function has identfied the title, first name, middle name, and suffix, it assumes that
    '           whatever is left must be the surname/last name
    '       d) Thus, this function will process compound first/middle names incorrectly
    ' 6) If you are using False for FirstMiddleLast:
    '       a) It is always assumed that the last comma in NameStr delimits the <Surname><Suffix> block
    '           from the <Title><First name><Middle name> block
    '       b) Whatever is left in the <Surname><Suffix> block after the suffix has been removed is assumed
    '           to be the last name
    '       c) After the Title is removed from the <Title><First name><Middle name> block, if there is only
    '           one "word", it is the first name.  If there are 2+ "words" and UseMiddle = True or omitted,
    '           then the last word is the middle name, and the rest is the first name
    '       d) Thus, this function will process compound middle names incorrectly, and may erroneously think
    '           a compound first name is a first name and a middle name
    
    Dim Title As String
    Dim FName As String
    Dim MName As String
    Dim LName As String
    Dim Suffix As String
    Dim RegXReturn As Object
    Dim NameArr As Variant
    Dim Counter As Long
    Dim StartsAt As Long
    Dim TitleLen As Long
    Dim LastComma As Long
    Dim Part1 As String
    Dim Part2 As String
    
    Static Titles As String
    Static Suffixes As String
    Static RegX As Object 'holding as a Static variable to improve performance
    
    If Trim(NameStr) = vbNullString Or Trim(NamePart) = vbNullString Then
        GetNamePart = ""
        Exit Function
    End If
    
    If Titles = vbNullString Then Titles = GenerateLists("Titles")
    If Suffixes = vbNullString Then Suffixes = GenerateLists("Suffixes")
        
    ' remove leading and trailing spaces
    
    NameStr = Trim(NameStr)
    
    ' instantiate RegExp if needed (static variable, so it will remain in between calls)
    
    If RegX Is Nothing Then
        Set RegX = CreateObject("VBScript.RegExp")
        With RegX
            .IgnoreCase = True  ' case insensitive
            .Global = True      ' finds all matches, not just first match
        End With
    End If
    
    ' Determine structure of NameStr
    
    If FirstMiddleLast Then
        
        ' NameStr is structured <Title (optional)> <First name> <Middle Name (optional)> <Surname> <Suffix (optional)>
        
        ' Set Pattern to look for titles at beginning of NameStr
        RegX.Pattern = "^(" & Titles & ")\.? +"
        
        ' Look for titles.  Use Do...Loop to allow for compound suffixes, as in "Rev. Mr. Arthur Dimmesdale"
        Do
            Set RegXReturn = RegX.Execute(Mid(NameStr, TitleLen + 1))
            If RegXReturn.count > 0 Then
                TitleLen = TitleLen + Len(RegXReturn(0))
            Else
                Exit Do
            End If
        Loop
        
        ' Get Title
        Title = Trim(Left(NameStr, TitleLen))
        
        ' Adjust NameStr
        NameStr = Mid(NameStr, TitleLen + 1)
        
        ' Look for suffixes.  Use Do...Loop to allow for compound suffixes, as in "Alfred E. Neumann, PhD, Esq."
        
        ' StartsAt indicates where the suffix(es) start in the NameStr.  Initialize to -1, indicating no suffix
        StartsAt = -1
        
        ' Set Pattern to look for suffix at end of NameStr
        RegX.Pattern = "(, *| +)(" & Suffixes & ")\.?$"
        
        ' Evaluate the NameStr.  As long as a suffix is found in the portion evaluated, reset the StartsAt variable.
        ' When no more suffixes are found, terminate the loop
        Do
            Set RegXReturn = RegX.Execute(Left(NameStr, IIf(StartsAt > -1, StartsAt, Len(NameStr))))
            If RegXReturn.count > 0 Then
                StartsAt = RegXReturn(0).FirstIndex     ' Recall that FirstIndex starts at position zero, not 1 !!!
            Else
                Exit Do
            End If
        Loop
        
        ' If a suffix is found, then grab the entire suffix
        If StartsAt > -1 Then
            Suffix = Mid(NameStr, StartsAt + 1)
            ' remove comma if applicable
            If Left(Suffix, 1) = "," Then Suffix = Mid(Suffix, 2)
            ' remove possible leading space
            Suffix = Trim(Suffix)
            ' adjust NameStr to remove suffixes
            NameStr = Left(NameStr, StartsAt)
        End If
        
        ' Ensure we have space delimiters for remaining NameStr
        NameStr = Replace(NameStr, ".", ". ")
        
        ' Remove extraneous spaces
        RegX.Pattern = " {2,}"
        NameStr = Trim(RegX.Replace(NameStr, " "))
        
        ' Create zero-based array with remaining "words" in the name
        NameArr = Split(NameStr, " ")
        
        ' First name is always assumed to be in position zero
        FName = NameArr(0)
        
        ' Depending on how many "words" are left and whether middle name is assumed to possibly be there
        ' (UseMiddle argument), grab middle/last names
        
        Select Case UBound(NameArr)
            Case 0
                'no middle or last names
            Case 1
                ' first name and last name
                LName = NameArr(1)
            Case Else
                ' potentially first, middle, and last names are present
                If UseMiddle Then
                    ' position 1 is assumed to be middle name, and positions 2 to N the last name
                    MName = NameArr(1)
                    ' remaining "words" are for last name
                    For Counter = 2 To UBound(NameArr)
                        LName = LName & " " & NameArr(Counter)
                    Next
                    ' drop leading space
                    LName = Trim(LName)
                Else
                    ' assume no middle name, and all remaining words are for the last name
                    For Counter = 1 To UBound(NameArr)
                        LName = LName & " " & NameArr(Counter)
                    Next
                    ' drop leading space
                    LName = Trim(LName)
                End If
        End Select
    
    Else
    
        ' NameStr is structured <Surname> <Suffix (optional)>, <Title (optional)> <First name> <Middle Name (optional)>
        
        ' Find position of last comma
        LastComma = InStrRev(NameStr, ",")
        
        If LastComma > 0 Then
            
            ' Part1 will be <Surname> <Suffix (optional)> block;
            ' Part2 is <Title (optional)> <First name> <Middle Name (optional)>
            Part1 = Trim(Left(NameStr, LastComma - 1))
            Part2 = Trim(Mid(NameStr, LastComma + 1))
        
            ' Look for suffixes.  Use Do...Loop to allow for compound suffixes, as in "Neumann, PhD, Esq., Alfred E."
            
            ' StartsAt indicates where the suffix(es) start in Part1.  Initialize to -1, indicating no suffix
            StartsAt = -1
            
            ' Set Pattern to look for suffix at end of Part1
            RegX.Pattern = "(, *| +)(" & Suffixes & ")\.?$"
            
            ' Evaluate Part1.  As long as a suffix is found in the portion evaluated, reset the StartsAt variable.
            ' When no more suffixes are found, terminate the loop
            Do
                Set RegXReturn = RegX.Execute(Left(Part1, IIf(StartsAt > -1, StartsAt, Len(Part1))))
                If RegXReturn.count > 0 Then
                    StartsAt = RegXReturn(0).FirstIndex     ' Recall that FirstIndex starts at position zero, not 1 !!!
                Else
                    Exit Do
                End If
            Loop
            
            ' If a suffix is found, then grab the entire suffix
            If StartsAt > -1 Then
                Suffix = Mid(Part1, StartsAt + 1)
                ' remove comma if applicable
                If Left(Suffix, 1) = "," Then Suffix = Mid(Suffix, 2)
                ' remove possible leading space
                Suffix = Trim(Suffix)
                ' adjust Part1 to remove suffixes
                Part1 = Left(Part1, StartsAt)
            End If
            LName = Trim(Part1)
        
            ' Set Pattern to look for titles at beginning of Part2
            RegX.Pattern = "^(" & Titles & ")\.? +"
            
            ' Look for titles.  Use Do...Loop to allow for compound suffixes, as in "Dimmesdale, Rev. Mr. Arthur"
            Do
                Set RegXReturn = RegX.Execute(Mid(Part2, TitleLen + 1))
                If RegXReturn.count > 0 Then
                    TitleLen = TitleLen + Len(RegXReturn(0))
                Else
                    Exit Do
                End If
            Loop
            
            ' Get Title
            Title = Trim(Left(Part2, TitleLen))
            
            ' Adjust Part2
            Part2 = Mid(Part2, TitleLen + 1)
        
            ' Ensure we have space delimiters for remaining Part2
            Part2 = Replace(Part2, ".", ". ")
            
            ' Remove extraneous spaces
            RegX.Pattern = " {2,}"
            Part2 = Trim(RegX.Replace(Part2, " "))
            
            ' Grab first/middle names from Part2
            If UseMiddle And InStr(1, Part2, " ") > 0 Then
                MName = Mid(Part2, InStrRev(Part2, " ") + 1)
                FName = Left(Part2, InStrRev(Part2, " ") - 1)
            Else
                FName = Part2
            End If
        End If
        
    End If
    
    ' determine function's return value
    
    Select Case UCase(NamePart)
        Case "TITLE", "HONORIFIC", "T", "H", "1"
            GetNamePart = Title
        Case "FIRST NAME", "FIRSTNAME", "FNAME", "F NAME", "FIRST", "F", "FN", "F N", "2"
            GetNamePart = FName
        Case "MIDDLE NAME", "MIDDLENAME", "MNAME", "M NAME", "MIDDLE", "M", "MN", "M N", "3"
            GetNamePart = MName
        Case "LAST NAME", "LASTNAME", "LNAME", "L NAME", "LAST", "L", "LN", "L N", "SURNAME", "4"
            GetNamePart = LName
        Case "SUFFIX", "S", "5"
            GetNamePart = Suffix
        Case Else
            GetNamePart = ""
    End Select
    
    ' destroy object variable
    
    Set RegXReturn = Nothing
    Titles = ""
End Function

Public Function GetAllNameParts(NameStr As String, Optional UseMiddle As Boolean = True, _
    Optional FirstMiddleLast As Boolean = True)
    
    ' Function by Patrick Matthews
    
    ' This code may be freely used or distributed so long as you acknowledge authorship and cite the URL
    ' where you found it
    
    ' This function relies on Regular Expressions.  For more information on RegExp, please see:
    ' http://www.experts-exchange.com/articles/Programming/Languages/Visual_Basic/Parsing-Names-in-MS-Office-Visual-Basic-6-and-Visual-Basic-for-Applications.html
    
    ' This function parses a name string, and returns a Dictionary object (Microsoft Scripting Runtime) with
    ' keys corresponding to title, first name, middle name, surname, and suffix.  If a name part is missing from the
    ' full name, the Dictionary item associated with that key is a zero-length string.  The keys are:
    ' Title, First, Middle, Last, and Suffix (not case sensitive)
    
    ' This function can be used in any VBA or VB6 project.  However, it cannot be used directly in an Excel
    ' worksheet formula or an Access query
    
    ' The function takes the following arguments:
    '
    ' 1) NameStr is the full name to be parsed.  Its assumed structure is determined by the FirstMiddleLast
    '       argument
    ' 2) UseMiddle indicates whether or not a middle name *might* be present in the NameStr.  If True or
    '       omitted, the function assumes that a middle name *might* be there.  If False, it assumes there
    '       is never a middle name
    ' 3) FirstMiddleLast indicates the order of the name parts.  If True or omitted, the function assumes:
    '       <Title (optional)> <First name> <Middle Name (optional)> <Surname> <Suffix (optional)>
    '       If False, the function assumes:
    '       <Surname> <Suffix (optional)>, <First name> <Middle Name (optional)>
    
    ' Notes:
    ' 1) The function has a prebuilt list of Titles (see GenerateLists function below), which you can modify to
    '       fit your needs.  The Titles string will be embedded in a RegExp Pattern string, so be sure to follow
    '       proper RegExp Pattern syntax
    ' 2) The function will recognize compound titles, as long as they are delimited by spaces
    ' 3) The function has a prebuilt list of Suffixes (see GenerateLists function below), which you can modify to
    '       fit your needs.  The Suffixes string will be embedded in a RegExp Pattern string, so be sure to
    '       follow proper RegExp Pattern syntax
    ' 4) The function will recognize compound suffixes, as long as they are delimited by commas and/or spaces
    ' 5) If you are using True (or omitting) for FirstMiddleLast:
    '       a) It is always assumed that the first name has a single "word"
    '       b) It is always assumed that the middle name, if present, has a single "word"
    '       c) After the function has identfied the title, first name, middle name, and suffix, it assumes that
    '           whatever is left must be the surname/last name
    '       d) Thus, this function will process compound first/middle names incorrectly
    ' 6) If you are using False for FirstMiddleLast:
    '       a) It is always assumed that the last comma in NameStr delimits the <Surname><Suffix> block
    '           from the <Title><First name><Middle name> block
    '       b) Whatever is left in the <Surname><Suffix> block after the suffix has been removed is assumed
    '           to be the last name
    '       c) After the Title is removed from the <Title><First name><Middle name> block, if there is only
    '           one "word", it is the first name.  If there are 2+ "words" and UseMiddle = True or omitted,
    '           then the last word is the middle name, and the rest is the first name
    '       d) Thus, this function will process compound middle names incorrectly, and may erroneously think
    '           a compound first name is a first name and a middle name
    
    Dim Title As String
    Dim FName As String
    Dim MName As String
    Dim LName As String
    Dim Suffix As String
    Dim RegXReturn As Object
    Dim NameArr As Variant
    Dim Counter As Long
    Dim StartsAt As Long
    Dim TitleLen As Long
    Dim LastComma As Long
    Dim Part1 As String
    Dim Part2 As String
    Dim dic As Object
    
    Static Titles As String
    Static Suffixes As String
    Static RegX As Object 'holding as a Static variable to improve performance
    
    If Titles = "" Then Titles = GenerateLists("Titles")
    If Suffixes = "" Then Suffixes = GenerateLists("Suffixes")
    
    Set dic = CreateObject("Scripting.Dictionary")
    With dic
        .CompareMode = 1
        .Add "Title", ""
        .Add "First", ""
        .Add "Middle", ""
        .Add "Last", ""
        .Add "Suffix", ""
    End With
    
    If Trim(NameStr) = "" Then
        Set GetAllNameParts = dic
        Set dic = Nothing
        Exit Function
    End If
    
    ' remove leading and trailing spaces
    
    NameStr = Trim(NameStr)
    
    ' instantiate RegExp if needed (static variable, so it will remain in between calls)
    
    If RegX Is Nothing Then
        Set RegX = CreateObject("VBScript.RegExp")
        With RegX
            .IgnoreCase = True  ' case insensitive
            .Global = True      ' finds all matches, not just first match
        End With
    End If
    
    ' Determine structure of NameStr
    
    If FirstMiddleLast Then
        
        ' NameStr is structured <Title (optional)> <First name> <Middle Name (optional)> <Surname> <Suffix (optional)>
        
        ' Set Pattern to look for titles at beginning of NameStr
        RegX.Pattern = "^(" & Titles & ")\.? +"
        
        ' Look for titles.  Use Do...Loop to allow for compound suffixes, as in "Rev. Mr. Arthur Dimmesdale"
        Do
            Set RegXReturn = RegX.Execute(Mid(NameStr, TitleLen + 1))
            If RegXReturn.count > 0 Then
                TitleLen = TitleLen + Len(RegXReturn(0))
            Else
                Exit Do
            End If
        Loop
        
        ' Get Title
        Title = Trim(Left(NameStr, TitleLen))
        
        ' Adjust NameStr
        NameStr = Mid(NameStr, TitleLen + 1)
        
        ' Look for suffixes.  Use Do...Loop to allow for compound suffixes, as in "Alfred E. Neumann, PhD, Esq."
        
        ' StartsAt indicates where the suffix(es) start in the NameStr.  Initialize to -1, indicating no suffix
        StartsAt = -1
        
        ' Set Pattern to look for suffix at end of NameStr
        RegX.Pattern = "(, *| +)(" & Suffixes & ")\.?$"
        
        ' Evaluate the NameStr.  As long as a suffix is found in the portion evaluated, reset the StartsAt variable.
        ' When no more suffixes are found, terminate the loop
        Do
            Set RegXReturn = RegX.Execute(Left(NameStr, IIf(StartsAt > -1, StartsAt, Len(NameStr))))
            If RegXReturn.count > 0 Then
                StartsAt = RegXReturn(0).FirstIndex     ' Recall that FirstIndex starts at position zero, not 1 !!!
            Else
                Exit Do
            End If
        Loop
        
        ' If a suffix is found, then grab the entire suffix
        If StartsAt > -1 Then
            Suffix = Mid(NameStr, StartsAt + 1)
            ' remove comma if applicable
            If Left(Suffix, 1) = "," Then Suffix = Mid(Suffix, 2)
            ' remove possible leading space
            Suffix = Trim(Suffix)
            ' adjust NameStr to remove suffixes
            NameStr = Left(NameStr, StartsAt)
        End If
        
        ' Ensure we have space delimiters for remaining NameStr
        NameStr = Replace(NameStr, ".", ". ")
        
        ' Remove extraneous spaces
        RegX.Pattern = " {2,}"
        NameStr = Trim(RegX.Replace(NameStr, " "))
        
        ' Create zero-based array with remaining "words" in the name
        NameArr = Split(NameStr, " ")
        
        ' First name is always assumed to be in position zero
        FName = NameArr(0)
        
        ' Depending on how many "words" are left and whether middle name is assumed to possibly be there
        ' (UseMiddle argument), grab middle/last names
        
        Select Case UBound(NameArr)
            Case 0
                'no middle or last names
            Case 1
                ' first name and last name
                LName = NameArr(1)
            Case Else
                ' potentially first, middle, and last names are present
                If UseMiddle Then
                    ' position 1 is assumed to be middle name, and positions 2 to N the last name
                    MName = NameArr(1)
                    ' remaining "words" are for last name
                    For Counter = 2 To UBound(NameArr)
                        LName = LName & " " & NameArr(Counter)
                    Next
                    ' drop leading space
                    LName = Trim(LName)
                Else
                    ' assume no middle name, and all remaining words are for the last name
                    For Counter = 1 To UBound(NameArr)
                        LName = LName & " " & NameArr(Counter)
                    Next
                    ' drop leading space
                    LName = Trim(LName)
                End If
        End Select
    
    Else
    
        ' NameStr is structured <Surname> <Suffix (optional)>, <Title (optional)> <First name> <Middle Name (optional)>
        
        ' Find position of last comma
        LastComma = InStrRev(NameStr, ",")
        
        If LastComma > 0 Then
            
            ' Part1 will be <Surname> <Suffix (optional)> block;
            ' Part2 is <Title (optional)> <First name> <Middle Name (optional)>
            Part1 = Trim(Left(NameStr, LastComma - 1))
            Part2 = Trim(Mid(NameStr, LastComma + 1))
        
            ' Look for suffixes.  Use Do...Loop to allow for compound suffixes, as in "Neumann, PhD, Esq., Alfred E."
            
            ' StartsAt indicates where the suffix(es) start in Part1.  Initialize to -1, indicating no suffix
            StartsAt = -1
            
            ' Set Pattern to look for suffix at end of Part1
            RegX.Pattern = "(, *| +)(" & Suffixes & ")\.?$"
            
            ' Evaluate Part1.  As long as a suffix is found in the portion evaluated, reset the StartsAt variable.
            ' When no more suffixes are found, terminate the loop
            Do
                Set RegXReturn = RegX.Execute(Left(Part1, IIf(StartsAt > -1, StartsAt, Len(Part1))))
                If RegXReturn.count > 0 Then
                    StartsAt = RegXReturn(0).FirstIndex     ' Recall that FirstIndex starts at position zero, not 1 !!!
                Else
                    Exit Do
                End If
            Loop
            
            ' If a suffix is found, then grab the entire suffix
            If StartsAt > -1 Then
                Suffix = Mid(Part1, StartsAt + 1)
                ' remove comma if applicable
                If Left(Suffix, 1) = "," Then Suffix = Mid(Suffix, 2)
                ' remove possible leading space
                Suffix = Trim(Suffix)
                ' adjust Part1 to remove suffixes
                Part1 = Left(Part1, StartsAt)
            End If
            LName = Trim(Part1)
        
            ' Set Pattern to look for titles at beginning of Part2
            RegX.Pattern = "^(" & Titles & ")\.? +"
            
            ' Look for titles.  Use Do...Loop to allow for compound suffixes, as in "Dimmesdale, Rev. Mr. Arthur"
            Do
                Set RegXReturn = RegX.Execute(Mid(Part2, TitleLen + 1))
                If RegXReturn.count > 0 Then
                    TitleLen = TitleLen + Len(RegXReturn(0))
                Else
                    Exit Do
                End If
            Loop
            
            ' Get Title
            Title = Trim(Left(Part2, TitleLen))
            
            ' Adjust Part2
            Part2 = Mid(Part2, TitleLen + 1)
        
            ' Ensure we have space delimiters for remaining Part2
            Part2 = Replace(Part2, ".", ". ")
            
            ' Remove extraneous spaces
            RegX.Pattern = " {2,}"
            Part2 = Trim(RegX.Replace(Part2, " "))
            
            ' Grab first/middle names from Part2
            If UseMiddle And InStr(1, Part2, " ") > 0 Then
                MName = Mid(Part2, InStrRev(Part2, " ") + 1)
                FName = Left(Part2, InStrRev(Part2, " ") - 1)
            Else
                FName = Part2
            End If
        End If
        
    End If
    
    ' determine function's return value
        
    With dic
        .Item("Title") = Title
        .Item("First") = FName
        .Item("Middle") = MName
        .Item("Last") = LName
        .Item("Suffix") = Suffix
    End With
    
    Set GetAllNameParts = dic
    
    ' destroy object variable
    
    Set RegXReturn = Nothing
    Set dic = Nothing
    
End Function
    
Private Function GenerateLists(ListType As String)
    
    Dim Titles As String
    Dim Suffixes As String
    
    ' In creating the master title and suffix lists, keep in mind that the strings will be passed in as part of a
    ' RegExp pattern, and so typical syntax rules for the VBScript implementation of RegExp will apply for things
    ' such as optional characters and escaping reserved characters.  For example:
    '
    '           M\.? ?D
    '
    ' matches M, then zero or one period, then zero or one space, then D.  Use the pipe character to delimit your
    ' entries
    
    ' If the lists get too long to keep using line continuators, then simply break them up into separate expressions:
    '
    ' Titles = 'Dr|Doctor|Mrs|Ms|Miss|Mr|Mister|Master|'
    ' Titles = Titles & "Reverend|Rev|Right Reverend|Right Rev|Most Reverend|
    ' Titles = Titles & "Most Rev|Honorable|Honourable"
    
    ' Populate master title list.  This can be expanded according to your needs.  There is no need to include a
    ' trailing period here, as the Pattern string built later on includes an optional period at the end.  In cases
    ' where a title may be shortened, list the longer version first.  For example, list Senator before Sen.
    
    Titles = "Dr|Doctor|Mrs|Ms|Miss|Mr|Mister|Master|Reverend|Rev|Right Reverend|Right Rev|Most Reverend|" & _
        "Most Rev|Honorable|Honourable|Hon|Monsignor|Msgr|Father|Fr|Bishop|Sister|Sr|Mother Superior|Mother|" & _
        "Senator|Sen|President|Pres|Vice President|V\.? ?P|Secretary|Sec|General|Gen|Lieutenant General|Lt\.? ?Gen|" & _
        "Major General|Maj\.? ?Gen|Brigadier General|Brig\.? ?Gen|Colonel|Col|Lieutenant Colonel|Lt\.? ?Col|Major|" & _
        "Maj|Sir|Dame|Lord|Lady|Judge|Professor|Prof|Sheikh|Shaikh|H\.? ?H\.? ? Sheikh|Captain|H\.? ?H\.? ? Sheikha|H\.? ?E\.|" & _
        "Eng|Sheikha|Capt|HE|Lt"
    
    ' Populate master suffix list.  This can be expanded according to your needs.  There is no need to include a
    ' trailing period here, as the Pattern string built later on includes an optional period at the end.  In cases
    ' where a title may be shortened, list the longer version first.  For example, list Esquire before Esq.  Also,
    ' list III before II, and II before I
    
    Suffixes = "M\.? ?D|Ph\.? ?D|Esquire|Esq\.?|J\.? ?D|D\.? ?D|Jr|Sr|III|II|I|IV|X|IX|VIII|VII|VI|V|M\.? ?P|" & _
        "M\.? ?S\.? ?W|C\.? P\.? ?A|P\.? M\.? ?P|L\.? ?P\.? ?N|R\.? ?N|A\.? ?S\.? ?E|U\.? ?S\.? ?N|" & _
        "U\.? ?S\.? ?M\.? ?C|R\.? ?G\.? ?C\.? ?E|P\.? ?M\.? ?P|P\.? ?E|M\.? ?O\.? ?S|M\.? ?C\.? ?T\.? ?S|" & _
        "M\.? ?C\.? ?T|M\.? ?C\.? ?S\.? ?E|M\.? ?C\.? ?S\.? ?D\.? ?|M\.? ?C\.? ?S\.? ?A|M\.? ?C\.? ?P\.? ?D|" & _
        "M\.? ?C\.? ?M|M\.? ?C\.? ?L\.? ?T|M\.? ?C\.? ?I\.? ?T\.? ?P|M\.? ?C\.? ?D\.? ?S\.? ?T|" & _
        "M\.? ?C\.? ?D\.? ?B\.? ?A|M\.? ?C\.? ?B\.? ?M\.? ?S\.? ?S|M\.? ?C\.? ?B\.? ?M\.? ?S\.? ?P|" & _
        "M\.? ?C\.? ?A\.? ?S|M\.? ?C\.? ?A\.? ?D|M\.? ?C\.? ?A|I\.? ?T\.? ?I\.? ?L|C\.? ?R\.? ?P|C\.? ?N\.? ?E|" & _
        "C\.? ?N\.? ?A|C\.? ?I\.? ?S\.? ?S\.? ?P|C\.? ?C\.? ?V\.? ?P|C\.? ?C\.? ?S\.? ?P|C\.? ?C\.? ?N\.? ?P|" & _
        "C\.? ?C\.? ?I\.? ?E|C\.? ?A\.? ?P\.? ?M|S\.? ?J|O\.? ?F\.? ?M|C\.? ?N\.? ?D|M\.? ?B\.? ?A|M\.? ?S"
        
    If ListType = "Titles" Then
        GenerateLists = Titles
    Else
        GenerateLists = Suffixes
    End If
    
End Function

Public Sub prepareGDNSheet()

Dim varLastRowNum As String
Dim i As Integer

'Check to make sure that the code runs only when the specified file is open.
If UCase(Left(ActiveWorkbook.name, 8)) <> "LR SALES" Then
'If UCase(Left(ActiveWorkbook.Name, 4)) <> "TEST" Then
    MsgBox "The data workbook with name starting with 'TEST' is not found" & vbCrLf & vbCrLf & _
    "Program cannot run & will exit now.", vbOKOnly
    Exit Sub
End If

Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

varLastRowNum = "0"

Do While varLastRowNum = 0
    varLastRowNum = InputBox("Please advise number of records to prepare", "Record Count")
    'End procedure
    If varLastRowNum = vbNullString Then End
Loop


varLastRowNum = varLastRowNum + 1

    For i = 1 To 4
        Range("D:D").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Next i
    
    Range("D1").FormulaR1C1 = "Client Type"
        With Range("D2")
            .Value = "p"
            .AutoFill Destination:=Range("D2:D" & varLastRowNum), Type:=xlFillCopy
        End With

    Range("E1").FormulaR1C1 = "Title"
    Range("F1").FormulaR1C1 = "F+M Name"
    Range("G1").FormulaR1C1 = "Surname"
    Range("E2").FormulaR1C1 = "=PERSONAL.XLSB!GetNamePart(RC[3],1)"
    
    Application.CutCopyMode = False
    
    Range("F2").FormulaR1C1 = _
        "=CONCATENATE(PERSONAL.XLSB!GetNamePart(RC[2],2),CHAR(32),PERSONAL.XLSB!GetNamePart(RC[2],3))"
    Range("G2").FormulaR1C1 = "=PERSONAL.XLSB!GetNamePart(RC[1],4)"
    Range("E2:G2").AutoFill Destination:=Range("E2:G" & varLastRowNum)
    Range("F2:G" & varLastRowNum).Select
    Range("I:I").Delete Shift:=xlToLeft
    Range("I:I").NumberFormat = "0"
    '**********************************************
    
    '=VLOOKUP($B2,Data!$D:$G,4,FALSE)
    Range("I2").FormulaR1C1 = _
        "=VLOOKUP(RC[-7],Lookup_Tables!C[-5]:C[-2],4,FALSE)"  'Some issue here relating to unsolicited conditional format.
    '**********************************************
    Range("I2").AutoFill Destination:=Range("I2:I" & varLastRowNum)  'Some issue here relating to unsolicited conditional format.
    Range("I2:I" & varLastRowNum).Calculate
    'Range("I2").Select
    'Range(Selection, Selection.End(xlDown)).Copy
    'Range(Selection, Selection.End(xlDown)).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
   
    For i = 1 To 4
        Range("I:I").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Next i
    
    Range("H1").FormulaR1C1 = "Customer Full Name"
    Range("I1").FormulaR1C1 = "Mobile Number"
    Range("J1").FormulaR1C1 = "Registration Number"
    Range("K1").FormulaR1C1 = "P.O. Box"
    Range("L1").FormulaR1C1 = "City"
    Range("M1").FormulaR1C1 = "Base Price"
    Range("N1").FormulaR1C1 = "Vista Sale Type"
    Range("O1").FormulaR1C1 = "Report Date"
    Range("P1").FormulaR1C1 = "Email Address"
    Range("Q1").FormulaR1C1 = "Payment Mode"
        With Range("O2")
            .Value = Date
            .Copy Destination:=Range("O2:O" & varLastRowNum)
        End With
    
    Application.CutCopyMode = False
    Range("A1").Select
    
    Call gDNEntrySheetFormat
    
    ActiveWorkbook.Save
    
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

Public Sub prepareGDNSheet_2()

    Call clearZeros
    Call conditionalFormatROReport

End Sub


Sub checkEntrySheet()

'Check City

Call checkCity

'Check for PO Box#7310 & highlight font in red

'Check Customer Name



End Sub

Sub checkCity()

Dim wi As Worksheet, wd As Worksheet, N As Range, F As Range, r As Long, SX As String

Set wi = Sheets("Input"): Set wd = Sheets("Data")

For Each N In wi.Range("B4:B" & wi.Range("B" & Rows.count).End(xlUp).Row)
    SX = SOUNDEX(CStr(N)): Set F = wd.Range("B:B").Find(SX, , , xlWhole)
    If Not F Is Nothing Then N.Offset(0, 4) = F.Offset(0, 1)
Next

End Sub

Function SOUNDEX(Surname As String) As String
' Developed by Richard J. Yanco
' This function follows the Soundex rules given at
' http://home.utah-inter.net/kinsearch/Soundex.html

    Dim Result As String, c As String * 1
    Dim Location As Integer

    Surname = UCase(Surname)

'   First character must be a letter
    If Asc(Left(Surname, 1)) < 65 Or Asc(Left(Surname, 1)) > 90 Then
        SOUNDEX = ""
        Exit Function
    Else
'       St. is converted to Saint
        If Left(Surname, 3) = "ST." Then
            Surname = "SAINT" & Mid(Surname, 4)
        End If

'       Convert to Soundex: letters to their appropriate digit,
'                     A,E,I,O,U,Y ("slash letters") to slashes
'                     H,W, and everything else to zero-length string

        Result = Left(Surname, 1)
        For Location = 2 To Len(Surname)
            Result = Result & Category(Mid(Surname, Location, 1))
        Next Location
      
'       Remove double letters
        Location = 2
        Do While Location < Len(Result)
            If Mid(Result, Location, 1) = Mid(Result, Location + 1, 1) Then
                Result = Left(Result, Location) & Mid(Result, Location + 2)
            Else
                Location = Location + 1
            End If
        Loop
    
'       If category of 1st letter equals 2nd character, remove 2nd character
        If Category(Left(Result, 1)) = Mid(Result, 2, 1) Then
            Result = Left(Result, 1) & Mid(Result, 3)
        End If
    
'       Remove slashes
        For Location = 2 To Len(Result)
            If Mid(Result, Location, 1) = "/" Then
                Result = Left(Result, Location - 1) & Mid(Result, Location + 1)
            End If
        Next
    
'       Trim or pad with zeroes as necessary
        Select Case Len(Result)
            Case 4
                SOUNDEX = Result
            Case Is < 4
                SOUNDEX = Result & String(4 - Len(Result), "0")
            Case Is > 4
                SOUNDEX = Left(Result, 4)
        End Select
    End If
End Function

Private Function Category(c) As String
'   Returns a Soundex code for a letter
    Select Case True
        Case c Like "[AEIOUY]"
            Category = "/"
        Case c Like "[BPFV]"
            Category = "1"
        Case c Like "[CSKGJQXZ]"
            Category = "2"
        Case c Like "[DT]"
            Category = "3"
        Case c = "L"
            Category = "4"
        Case c Like "[MN]"
            Category = "5"
        Case c = "R"
            Category = "6"
        Case Else 'This includes H and W, spaces, punctuation, etc.
            Category = ""
    End Select
End Function

Private Sub gDNEntrySheetFormat()
    
    
    Dim lastRow As Integer
    
    
    lastRow = totalRecords("A", "2")
    
    'Select Data Range
    Range("A1:Q" & lastRow).Select
    
    'Apply Indentation to left
    With Selection
        .HorizontalAlignment = xlGeneral
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlLeft
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    'FORMAT BORDERS
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    Selection.AutoFilter
    
    'Delete conditional formats if any
    Cells.FormatConditions.Delete

    'Autofit
    Cells.EntireColumn.AutoFit
    
    'Ready cursor at 'C1' for data collection from AutoLine
    Range("C1").Select

End Sub

Private Function clrLZeros(ByVal Target As Range) As String

' Called from SUB clearZeros()

Dim varTemp As String

varTemp = Target.Value

Do While Left(varTemp, 1) = 0

    varTemp = Right(varTemp, Len(varTemp) - 1)

Loop

clrLZeros = varTemp

End Function

Public Sub clearZeros()

    Dim varLastRowNum As String
    
    'Check to make sure that the code runs only when the specified file is open.
    'If UCase(Left(ActiveWorkbook.Name, 4)) <> "TEST" Then
    If UCase(Left(ActiveWorkbook.name, 8)) <> "LR SALES" Then
        MsgBox "The data workbook with name starting with 'TEST' is not found" & vbCrLf & vbCrLf & _
        "Program cannot run & will exit now.", vbOKOnly
        Exit Sub
    End If
    
    varLastRowNum = "0"
    
    'Do While varLastRowNum = "0"
    '    varLastRowNum = InputBox("Please advise number of records to prepare", "Record Count")
    '    'End procedure
    '    If varLastRowNum = "" Then End
    'Loop
    
    
    varLastRowNum = totalRecords("A", "2")
    varLastRowNum = varLastRowNum + 1
    
    'Remove leading zeros on registration column
    Columns("K:K").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "=clrLZeros(RC[-1])" 'Function Call
    Range("K2").Select
    Selection.AutoFill Destination:=Range("K2:K" & varLastRowNum)
    Range("K2:K" & varLastRowNum).Select
    Selection.Copy
    Range("J2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("K:K").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    
    ActiveWorkbook.Save


End Sub

Public Sub conditionalFormatROReport()


    'Highlight PO Box# 7310(ATM) & 35288(RO)
    Columns("K:K").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=7310"
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .Color = -16776961
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    'Highlight duplicates in names section
    Columns("H:H").Select
    Selection.FormatConditions.AddUniqueValues
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
    Selection.FormatConditions(1).DupeUnique = xlDuplicate
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("H1").Select
    
    
Dim checkStrings() As String 'Declaring string array to store strings that needs to be searched.
Dim searchString As String
Dim cnt As Integer

checkStrings = Split("LandRover,Land Rover,Jaguar,Demo,Courtesy,Tayer,Fazza", ",") 'Initializing string array

For cnt = 0 To 6

    searchString = "=SEARCH(" & Chr(34) & checkStrings(cnt) & Chr(34) & ",H1)"
    'Highlight RO & ATM deals
    Columns("H:H").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        searchString
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.599963377788629
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    ActiveWindow.SmallScroll Down:=-3
    
Next cnt

End Sub



