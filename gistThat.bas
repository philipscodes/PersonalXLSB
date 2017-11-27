Attribute VB_Name = "gistThat"
'IMPORTANT - CHANGE gtExampleLoad() to include gtDoit() for each Gist you want to load
' bootstrap code to update VBA modules from gists
' all code is in this module - so no classes etc.
' latebinding is used to avoid need for any references
' can be found at https://gist.github.com/3403537
Option Explicit
' v2.04 - 3403537
' if you are using your own gists - change this
Const gistOwner = "philipscodes"

Public Function gtExampleLoad()
' this is an example of how you would load your VBE with a particular manifest
' you could set the 2nd parameter to overrride conflict checking the first time used-

' we are going to need cDataSet and cjobject stuff as well
  gtDoit "3414394", True
  
' and crest
  gtDoit "5055578", True
  
'uameasure - if instrumenting.
  gtDoit "7471153", True
  
End Function

Public Function gtDeadDropLoad()

' scriptdbcom & deaddrop
  gtDoit "8767201", True
  
' parsecom - may want to use parse.com instead
  gtDoit "7663169", True

End Function
Private Function gtExampleMakeManifestScriptDbCom()
    ' this is an example of how you would create a manifest to be loaded up as a Gist
    '
    Dim dom As Object ' DOMDocument
    Set dom = gtInitManifest("cDataset and associated classes and modules", "bruce@mcpher.com")
    '
    ' call this for each required gist of the manifest
    '---cDataSet
    gtAddToManifest dom, "3414216", "class", "cCell", "cCell.cls"
    gtAddToManifest dom, "3414216", "class", "cDataSet", "cDataSet.cls"
    gtAddToManifest dom, "3414216", "class", "cDataSets", "cDataSets.cls"
    gtAddToManifest dom, "3414216", "class", "cDataColumn", "cDataColumn.cls"
    gtAddToManifest dom, "3414216", "class", "cDataRow", "cDataRow.cls"
    gtAddToManifest dom, "3414216", "class", "cHeadingRow", "cHeadingRow.cls"
    '--Useful modules
    gtAddToManifest dom, "3414346", "module", "usefulStuff"
    '--cJobject
    gtAddToManifest dom, "3414365", "class", "cJobject"
    gtAddToManifest dom, "3414365", "module", "usefulcJobject", "usefulcJobject.vba"

    '--regex
    gtAddToManifest dom, "3414836", "module", "regXLib", "regXLib.vba"
    gtAddToManifest dom, "3414836", "class", "cregXLib", "cregXLib.cls"
    
    '-- browser & auth
    gtAddToManifest dom, "3423912", "class", "cBrowser", "cBrowser.cls"
    gtAddToManifest dom, "6937450", "class", "cOauth2", "cOauth2.cls"
    gtAddToManifest dom, "6937450", "module", "oAuthExamples", "oAuthExamples.vba"
    
    '-- scriptdb and deaddrop
    gtAddToManifest dom, "7943765", "class", "cScriptDbCom", "cScriptDbCom.cls"
    gtAddToManifest dom, "7943765", "class", "cDeadDrop", "cDeadDrop.cls"
    gtAddToManifest dom, "7943765", "module", "scriptDbCom", "scriptDbCom.vba"
    
    '-- encryption
    gtAddToManifest dom, "6937529", "module", "usefulEncrypt", "usefulEncrypt.vba"
    
        '-- analytics
    gtAddToManifest dom, "7453196", "module", "UAMeasure", "UAMeasure.vba"
    gtAddToManifest dom, "7453152", "class", "cUAMeasure", "cUAMeasure.cls"
    
    ' -- references
    gtCreateReferences dom
    

    
    ' cut and paste the result of this into a gist - this will be your manifest
    Debug.Print dom.XML
    
End Function
Private Function gtExampleMakeManifest()
    ' this is an example of how you would create a manifest to be loaded up as a Gist
    '
    Dim dom As Object ' DOMDocument
    Set dom = gtInitManifest("cDataset and associated classes and modules", "bruce@mcpher.com")
    '
    ' call this for each required gist of the manifest
    '---cDataSet
    gtAddToManifest dom, "3414216", "class", "cCell", "cCell.cls"
    gtAddToManifest dom, "3414216", "class", "cDataSet", "cDataSet.cls"
    gtAddToManifest dom, "3414216", "class", "cDataSets", "cDataSets.cls"
    gtAddToManifest dom, "3414216", "class", "cDataColumn", "cDataColumn.cls"
    gtAddToManifest dom, "3414216", "class", "cDataRow", "cDataRow.cls"
    gtAddToManifest dom, "3414216", "class", "cHeadingRow", "cHeadingRow.cls"
    '--Useful modules
    gtAddToManifest dom, "3414346", "module", "usefulStuff"
    gtAddToManifest dom, "5102369", "class", "cStringChunker.cls"

    '--cJobject
    gtAddToManifest dom, "3414365", "class", "cJobject"
    gtAddToManifest dom, "3414365", "module", "usefulcJobject", "usefulcJobject.vba"

    '--regex
    gtAddToManifest dom, "3414836", "module", "regXLib", "regXLib.vba"
    gtAddToManifest dom, "3414836", "class", "cregXLib", "cregXLib.cls"
    
    
    ' -- references
    gtCreateReferences dom
    
    ' cut and paste the result of this into a gist - this will be your manifest
    Debug.Print dom.XML
    
End Function
Private Function gtExampleMakeManifestCrest()
    ' this is an example of how you would create a manifest to be loaded up as a Gist
    '
    Dim dom As Object ' DOMDocument
    Set dom = gtInitManifest("restlibrary associated classes and modules", "bruce@mcpher.com")
    '
    ' call this for each required gist of the manifest

    
    ' crest and restlibrary
    gtAddToManifest dom, "3423885", "module", "restLibrary", "restLibrary.vba"
    gtAddToManifest dom, "3423885", "class", "cRest", "cRest.cls"
    gtAddToManifest dom, "3423912", "class", "cBrowser", "cBrowser.cls"
    
    ' -- references
    gtCreateReferences dom
    
    ' cut and paste the result of this into a gist - this will be your manifest
    Debug.Print dom.XML
    
End Function


Private Function gtClassDocumenter()
    ' this is an example of how you would create a manifest to be loaded up as a Gist
    ' this is modules required for classDoucumenter
    Dim dom As Object ' DOMDocument
    Set dom = gtInitManifest("class", "bruce@mcpher.com")
    '
    ' call this for each required gist of the manifest
    '---cVBA
    gtAddToManifest dom, "5019927", "class", "cVBAProject", "cVBAProject.vba"
    gtAddToManifest dom, "5019927", "class", "cVBAModule", "cVBAModule.vba"
    gtAddToManifest dom, "5019927", "class", "cVBAProcedure", "cVBAProcedure.vba"
    gtAddToManifest dom, "5019927", "module", "classSerializer", "classSerializer.vba"
    '-- browser
    gtAddToManifest dom, "3423912", "class", "cBrowser", "cBrowser.cls"
    
    ' -- references
    gtCreateReferences dom
    
    ' cut and paste the result of this into a gist - this will be your manifest
    Debug.Print dom.XML
    
End Function
Private Function gtCreateReferences(dom As Object) 'DOMDocument)
    ' adds all current references to an xml
    Dim r As Object ' Reference
    
    With ActiveWorkbook.VBProject
        For Each r In .References
            gtAddRefToManifest dom, r
        Next r
    End With

End Function
Private Function gtUpdateAll()
    ' do all updates for known manifests in this project
    Dim modle As Object 'VBComponent
    Dim stampLine As Long, co As New Collection, manifest As String, s As String, v As Variant
    For Each modle In ThisWorkbook.VBProject.VBComponents
        ' do we know this module?
        stampLine = gtManageable(modle)
        If stampLine > 0 Then
            ' yes we do - get the manifest
            manifest = gtStampManifest(modle, stampLine)
            If (manifest = vbNullString) Then
                MsgBox ("gistThat stamp in module " & modle.name & " fiddled with.Run again as greenField")
            Else
            ' add it to the collection of manifests we need to process
              If gtCoIndex(manifest, co) = 0 Then
                co.Add manifest, manifest
              End If
            End If
        End If
    Next modle
    ' todo check if versions are up to date
    If co.count > 0 Then
        s = vbNullString
        For Each v In co
            gtDoit CStr(v)
            If s <> vbNullString Then s = s & ","
            s = s & CStr(v)
        Next v
        MsgBox ("updated " & co.count & " manifests(" & s & ")")
    End If
End Function
Private Function gtCoIndex(sid As Variant, co As Collection) As Long
    ' find index in a collection
    Dim i As Long
    gtCoIndex = 0
    For i = 1 To co.count
        If co(i) = sid Then
          gtCoIndex = i
          Exit Function
        End If
    Next i
End Function
Private Function gtPreventCaching(url As String) As String
    ' this will tweak the url with an extra random parameter to prevent any accidental caching
    Dim p As String
    If (InStr(1, url, "?") > 0) Then
        p = "&"
    Else
        p = "?"
    End If
    Randomize
    gtPreventCaching = url & p & "gtPreventCaching=" & CStr(Int(10000 * Rnd()))
End Function
Public Function gtDoit(gtDoitmanifestID As String, Optional greenField As Boolean = False) As Boolean
    Dim dom As Object ' DOMDocument
    Dim rawUrl As String, t As String, N As String, g As String
    Dim xNode As Object ' IXMLDOMNode
    Dim attrib As Object 'IXMLDOMAttribute
    Dim vbCom As Object 'VBComponent
    ' get the requested manifest
    Set dom = gtRecreateManifest(gtDoitmanifestID)

    ' now we know which gists are needed here
    If (gtWillItWork(dom, greenField)) Then
        'theres a good chance it will work
        ' for each module
        For Each xNode In dom.SelectSingleNode("//gists").ChildNodes
            t = xNode.Attributes.getNamedItem("type").Text
            Select Case t
                Case "class", "module"
                    ' get the gist
                    rawUrl = gtConstructRawUrl(xNode.Attributes.getNamedItem("gistid").Text, _
                                            xNode.Attributes.getNamedItem("filename").Text)
                    ' prevent caching will make it look like a different request
                    g = gtHttpGet(gtPreventCaching(rawUrl))
                    ' module name
                    N = xNode.Attributes.getNamedItem("module").Text
                    ' does it exist - if so then delete it
                    Set vbCom = gtModuleExists(N, ThisWorkbook)
                    If (Not vbCom Is Nothing) Then
                        ' delete everything in it
                        With vbCom.CodeModule
                            .DeleteLines 1, .CountOfLines
                        End With
                    Else
                        Set vbCom = gtAddModule(N, ThisWorkbook, xNode.Attributes.getNamedItem("type").Text)
                    End If
        
                    ' add in the new code
                    With vbCom.CodeModule
                        .AddFromString g
                    End With
        
                    ' stamp it
                    gtInsertStamp vbCom, gtDoitmanifestID, rawUrl
                
                Case "reference"
                    gtAddReference xNode.Attributes.getNamedItem("name").Text, _
                                   xNode.Attributes.getNamedItem("guid").Text, _
                                   xNode.Attributes.getNamedItem("major").Text, _
                                   xNode.Attributes.getNamedItem("minor").Text
                Case Else
                    Debug.Assert False
            
            End Select
        Next xNode
        gtDoit = True
    Else

    End If
End Function

Private Function gtAddReference(name As String, guid As String, major As String, minor As String) As Object ' Reference
    ' add a reference (if its not already there)
    Dim r As Object ' Reference
    On Error GoTo handle
    With ActiveWorkbook.VBProject
        For Each r In .References
            If (r.name = name) Then
                If (r.major < major Or r.major = major And r.minor < minor And Not r.BuiltIn) Then
                    .References.AddFromGuid guid, major, minor
                    .References.Remove (r)
                End If
                Exit Function
            End If
        Next r
    ' if we get here then we need to add it
      Set gtAddReference = .References.AddFromGuid(guid, major, minor)
      Exit Function
    End With
    
handle:
    MsgBox ("warning - tried and failed to add reference to " & name & "v" & major & "." & minor)
    Exit Function
    
End Function
Private Function gtStampManifest(vbCom As Object, line As Long) As String 'VBComponent
    ' the manifest should be on the given line
    Dim s As String, N As Long, p As Long, marker As String
    marker = "manifest:"
    s = vbNullString
    With vbCom.CodeModule
       N = InStr(1, LCase(.Lines(line, 1)), marker)
       If (N > 0) Then
        s = Mid(.Lines(line, 1), N + Len(marker))
        p = InStr(1, s, " ")
        s = Left(s, p - 1)
       End If
    End With
    gtStampManifest = s
End Function
Private Function gtInsertStamp(vbCom As Object, manifest As String, rawUrl As String) As Long 'VBComponent
    Dim stampLine As Long
    stampLine = gtManageable(vbCom)
    ' if it wasnt found then insert at line 1
    With vbCom.CodeModule
        If stampLine <> 0 Then
            .DeleteLines stampLine, 1
        Else
            stampLine = 1
        End If
        .InsertLines stampLine, gtStampLog(manifest, rawUrl)
    End With
    gtInsertStamp = stampLine
    
End Function
Private Function gtWillItWork(dom As Object, _
                Optional greenField As Boolean = False) As Boolean 'DOMDocument
    
    Dim xNode As Object ' IXMLDOMNode
    Dim attrib As Object 'IXMLDOMAttribute
    Dim N As String, s As String, t As String
    Dim modle As Object 'VBComponent
    
    ' check we have something to do
    gtWillItWork = Not dom Is Nothing
    If Not gtWillItWork Then
        Exit Function
    End If
    ' first we check if these are new modules
    s = vbNullString
    For Each xNode In dom.SelectSingleNode("//gists").ChildNodes
        ' the target module
        t = xNode.Attributes.getNamedItem("type").Text
        Select Case t
            Case "class", "module"
                N = xNode.Attributes.getNamedItem("module").Text
                Set modle = gtModuleExists(N, ThisWorkbook)
                
                If (Not modle Is Nothing) Then
                    ' it exists - validate its not somethig else with the same name
                    If (gtManageable(modle) = 0 And Not greenField) Then
                        s = gtAddStr(s, N)
                    End If
                End If
            Case "reference"
            Case Else
                s = gtAddStr(s, "unknown type " & t)
        End Select
    Next xNode
    
    If (s <> vbNullString) Then
        MsgBox ("there may be a conflict with these modules names (" & s & _
            ") and some others in your project. " & _
            "If this is the first time you have run this - run with greenfield set to true to override this check")
        gtWillItWork = False
        Exit Function
    End If
   
   ' now check all gists are getable
   ' todo
   
End Function
Private Function gtAddStr(t As String, N As String) As String
    Dim s As String
    s = t
    If (s <> vbNullString) Then s = s & ","
    gtAddStr = s & N
End Function

Private Function gtRecreateManifest(manifestID As String) As Object 'DOMDocument
    Dim dom As Object 'DOMDocument
    Dim manifest As String
    
    ' get the xml string
    manifest = gtHttpGet(gtPreventCaching(gtConstructRawUrl(manifestID)))
    
    If manifest <> vbNullString Then
    ' parse the xml
        Set dom = CreateObject("MSXML.DOMDocument")
        dom.LoadXML (manifest)
        Set gtRecreateManifest = dom
    Else
        MsgBox ("Could not get manifest for " & manifestID)
    End If
End Function

Private Function gtModuleExists(name As String, wb As Workbook) As Object 'VBComponent
    ' determine whether this module exists in the given workbook
    Dim modle As Object 'VBComponent
    For Each modle In wb.VBProject.VBComponents
       If Trim(LCase(modle.name)) = Trim(LCase(name)) Then
        Set gtModuleExists = modle
        Exit Function
       End If
    Next modle
End Function

Private Function gtAddModule(name As String, wb As Workbook, modType As String) As Object ' VBComponent
    ' determine whether this module exists in the given workbook
    Dim modle As Object, t As String ' VBComponent, t As Long

    Select Case LCase(modType)
        Case "class"
            t = 2
        Case "module"
            t = 1
        Case Else
            MsgBox ("unknown module type " & modType)
    End Select
        
    Set modle = wb.VBProject.VBComponents.Add(t)
    modle.name = name
    
    ' added by andypope.info
    If modle.CodeModule.CountOfLines > 1 Then
        ' remove Option Explict lines if it was added automatically
        modle.CodeModule.DeleteLines 1, modle.CodeModule.CountOfLines
    End If
    
    Set gtAddModule = modle
End Function

Private Function gtConstructRawUrl(gistID As String, _
                Optional gistFileName As String = vbNullString) As String
    ' given a gist, where is it?
    Dim s As String
    ' raw URL
    s = "https://gist.github.com/" & gistOwner & "/" & gistID & "/raw"

    ' a gist can have multiple files in it
    If gistFileName <> vbNullString Then s = s & "/" & gistFileName
    ' TODO - specific versions
    gtConstructRawUrl = s
End Function

Private Function gtAddToManifest(dom As Object, _
                                 gistID As String, _
                                 modType As String, _
                                 modle As String, _
                                 Optional Filename As String = vbNullString, _
                                 Optional version As String = vbNullString _
                        ) As Object ' DOMDocument
                                 
    Dim Element As Object 'IXMLDOMElement
    Dim attrib As Object 'IXMLDOMAttribute
    Dim elements As Object 'IXMLDOMNodeList
    Dim head As Object 'IXMLDOMElement
    ' add an item to the manifest element - returns the dom for chaining
    Set elements = dom.getElementsByTagName("gists")
    Set head = elements.NextNode
    Set Element = dom.createElement("item" & CStr(head.ChildNodes.Length + 1))
    head.appendChild Element
    
    Set attrib = dom.createAttribute("gistid")
    attrib.NodeValue = gistID
    Element.setAttributeNode attrib
    
    Set attrib = dom.createAttribute("version")
    attrib.NodeValue = version
    Element.setAttributeNode attrib
    
    Set attrib = dom.createAttribute("filename")
    attrib.NodeValue = Filename
    Element.setAttributeNode attrib
    
    Set attrib = dom.createAttribute("module")
    attrib.NodeValue = modle
    Element.setAttributeNode attrib
    
    Set attrib = dom.createAttribute("type")
    attrib.NodeValue = modType
    Element.setAttributeNode attrib
    
    Set gtAddToManifest = dom
End Function
Private Function gtAddRefToManifest(dom As Object, r As Object) As Object   ' DOMDocument, reference, domdocument
                                 
    Dim Element As Object 'IXMLDOMElement
    Dim attrib As Object 'IXMLDOMAttribute
    Dim elements As Object 'IXMLDOMNodeList
    Dim head As Object 'IXMLDOMElement
    
    ' add an item to the manifest element - returns the dom for chaining
    Set elements = dom.getElementsByTagName("gists")
    Set head = elements.NextNode
    Set Element = dom.createElement("item" & CStr(head.ChildNodes.Length + 1))
    head.appendChild Element
    'r.GUID, r.name, r.Major, r.Minor, r.description
    Set attrib = dom.createAttribute("guid")
    attrib.NodeValue = r.guid
    Element.setAttributeNode attrib
    
    Set attrib = dom.createAttribute("name")
    attrib.NodeValue = r.name
    Element.setAttributeNode attrib
    
    Set attrib = dom.createAttribute("major")
    attrib.NodeValue = r.major
    Element.setAttributeNode attrib
    
    Set attrib = dom.createAttribute("minor")
    attrib.NodeValue = r.minor
    Element.setAttributeNode attrib

    Set attrib = dom.createAttribute("description")
    attrib.NodeValue = r.description
    Element.setAttributeNode attrib
    
    Set attrib = dom.createAttribute("type")
    attrib.NodeValue = "reference"
    Element.setAttributeNode attrib
    
    Set gtAddRefToManifest = dom
End Function
Private Function gtInitManifest(Optional description As String = vbNullString, _
                                 Optional contact As String = vbNullString) As Object ' DOMDocument
    Dim Element As Object 'IXMLDOMElement
    Dim attrib As Object 'IXMLDOMAttribute
    Dim dom As Object ' DOMDocument
    Dim e2 As Object ' IXMLDOMElement
   
 ' creates an xml manifest of required gists
   Set dom = CreateObject("MSXML.DOMDocument")

    Set Element = dom.createElement("gistThat")
    Set attrib = dom.createAttribute("info")
    attrib.NodeValue = _
            "this is a manifest for gistThat VBA code distribution " & _
            " - see ramblings.mcpher.com for details"
    Element.setAttributeNode attrib

    
    With dom.appendChild(Element)
        Set Element = dom.createElement("manifest")
        .appendChild Element
        

        
        Set attrib = dom.createAttribute("description")
        attrib.NodeValue = description
        Element.setAttributeNode attrib

        Set attrib = dom.createAttribute("contact")
        attrib.NodeValue = contact
        Element.setAttributeNode attrib

        Element.appendChild dom.createElement("gists")

    End With
    Set gtInitManifest = dom
   
End Function
 
Private Function gtHttpGet(url As String) As String
    ' TODO oAuth
    Dim ohttp As Object
    Set ohttp = CreateObject("Msxml2.ServerXMLHTTP.6.0")
    Call ohttp.Open("GET", url, False)
    Call ohttp.Send("")
    gtHttpGet = ohttp.ResponseText
    Set ohttp = Nothing
End Function
Private Function gtStampLog(manifest As String, rawUrl As String) As String
    ' create a comment to identify this as manageable
    gtStampLog = gtStamp & _
        " updated on " & Now() & " : from manifest:" & _
        manifest & _
        " gist " & rawUrl
End Function
Private Function gtStamp() As String
' this marks a module as manageable
    gtStamp = "'gistThat@mcpher.com :do not modify this line" & _
    " - see ramblings.mcpher.com for details:"
End Function
Private Function gtManageable(vbCom As Object) As Long  ' VBComponent
    ' return the line number of the gtStamp
    ' parameters as passed by ref in .find method
    Dim startLine As Long, startColumn As Long, endLine As Long, endColumn As Long
    startLine = 1: endLine = vbCom.CodeModule.CountOfLines: startColumn = 1: endColumn = 255
    
    If (vbCom.CodeModule.Find(gtStamp(), startLine, startColumn, endLine, endColumn)) Then
        gtManageable = startLine
    End If
    
End Function




