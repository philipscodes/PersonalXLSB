VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usrFrmNCReports 
   Caption         =   "Report Options"
   ClientHeight    =   4050
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   1992
   OleObjectBlob   =   "usrFrmNCReports.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usrFrmNCReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkBxSelect_Click()

    If Me.chkBxSelect = True Then
        Me.chkBxOrders = True
        Me.chkBxArrivals = True
        Me.chkBxStock = True
        Me.chkBxVCC = True
        Me.chkBxSales = True
    Else
        Me.chkBxOrders = False
        Me.chkBxArrivals = False
        Me.chkBxStock = False
        Me.chkBxVCC = False
        Me.chkBxSales = False
    
    End If

End Sub

Private Sub cmdOK_Click()

    If Me.chkBxOrders = True Then
        varPrepareOrders = True
    Else
        varPrepareOrders = False
    End If
    
    If Me.chkBxArrivals = True Then
        varPrepareArrivals = True
    Else
        varPrepareArrivals = False
    End If
    
    If Me.chkBxStock = True Then
        varPrepareStock = True
    Else
        varPrepareStock = False
    End If
    
    If Me.chkBxVCC = True Then
        varPrepareVCCRelease = True
    Else
        varPrepareVCCRelease = False
    End If
    
    If Me.chkBxSales = True Then
        varPrepareSales = True
    Else
        varPrepareSales = False
    End If
    
    Unload Me

End Sub

