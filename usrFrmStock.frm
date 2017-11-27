VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usrFrmStock 
   Caption         =   "Stock Reconciliation"
   ClientHeight    =   1860
   ClientLeft      =   6048
   ClientTop       =   3372
   ClientWidth     =   5760
   OleObjectBlob   =   "usrFrmStock.frx":0000
End
Attribute VB_Name = "usrFrmStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdExit_Click()

        'Call closingProc
        Unload Me
        'Call mainExit
    
End Sub

Private Sub cmdOK_Click()

Dim foundRange As Range
Dim check As VbMsgBoxResult
    
   
    If Len(Me.txtVIN) < 6 Then
        MsgBox "Minimum 6 characters required to check stock...", vbOKOnly, "Invalid Entry"
        Call txtBxFocus
        GoTo commonPoint
    End If
        
    
    'Call revertHighlightedEntry     '***FUNCTION CALL***
    
    'Call function findVIN
    Set foundRange = findvalue(Me.txtVIN)   '***FUNCTION CALL***
    
    ' If the input VIN is not found in Excel Stock Sheet then call function 'excessEntryQuery'
    If foundRange Is Nothing Then
      
        MsgBox "The input value is not found", vbOKOnly, "Not Found"
        Call txtBxFocus
        
    Else
        'Activate found cell and highlight the VIN in Red Font
        foundRange.Activate
        Call StartBlinking
        
    End If
        

commonPoint:

    'ActiveWorkbook.Save
    'Call txtBxFocus
    
End Sub

Private Sub txtVIN_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    
    If Len(Me.txtVIN) = 17 Then
        KeyAscii = 0
        MsgBox "VIN is only 17 characters maximum.", vbOKOnly, "VIN - 17 Characters Length"
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If

End Sub

