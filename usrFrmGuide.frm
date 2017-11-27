VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usrFrmGuide 
   Caption         =   "Sales File Usage Guide"
   ClientHeight    =   10410
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   6432
   OleObjectBlob   =   "usrFrmGuide.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "usrFrmGuide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Activate()
    With usrFrmGuide
        .Top = Application.Top + 50 '< change 125 to what u want
        .Left = Application.Left + 825 '< change 25 to what u want
    End With
End Sub

Private Sub UserForm_Initialize()

    AlwaysOnTop Me.caption
    
End Sub

Private Sub AlwaysOnTop(caption As String)

    Dim ret As Long
    Dim hWnd As Long
    
    hWnd = FindWindow(vbNullString, caption)
    ret = SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
     
    If CloseMode = 0 Then
        Cancel = True
        MsgBox "The X is disabled, please close this form using the Excel Ribbon.", vbCritical
    End If
     
End Sub

