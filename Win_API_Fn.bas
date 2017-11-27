Attribute VB_Name = "Win_API_Fn"
Public Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Public Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Public Declare PtrSafe Function CloseClipboard Lib "user32" () As Long

Public Function ClearClipboard()
    OpenClipboard (0&)
    EmptyClipboard
    CloseClipboard
End Function
