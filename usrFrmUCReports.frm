VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usrFrmUCReports 
   Caption         =   "Report Options"
   ClientHeight    =   3765
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   1992
   OleObjectBlob   =   "usrFrmUCReports.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usrFrmUCReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub chkBxSelect_Click()

    If Me.chkBxSelect = True Then
    
        Me.chkBxUCStock = True
        Me.chkBxPurchase = True
        Me.chkBxHiyaza = True
        Me.chkBxSales = True
        Me.chkBxSalesYTD = True
        
    Else
    
        Me.chkBxUCStock = False
        Me.chkBxPurchase = False
        Me.chkBxHiyaza = False
        Me.chkBxSales = False
        Me.chkBxSalesYTD = False
        
    End If

End Sub

Private Sub cmdOK_Click()

    
    If Me.chkBxUCStock = True Then
    
        varPrepareUCStock = True
        
    Else
    
        varPrepareUCStock = False
        
    End If
    
    
    
    If Me.chkBxPurchase = True Then
    
        varPrepareUCPurchase = True
        
    Else
    
        varPrepareUCPurchase = False
        
    End If
    
    
    
    If Me.chkBxHiyaza = True Then
    
        varPrepareHiyaza = True
        
    Else
    
        varPrepareHiyaza = False
        
    End If
    
    
    
    If Me.chkBxSales = True Then
    
        varPrepareUCSales = True
        
    Else
    
        varPrepareUCSales = False
        
    End If
    
    
    
    If Me.chkBxSalesYTD = True Then
    
        varPrepareUCSalesYTD = True
        
    Else
    
        varPrepareUCSalesYTD = False
        
    End If
    
    
    Unload Me

End Sub
