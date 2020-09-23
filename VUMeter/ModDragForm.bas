Attribute VB_Name = "ModDragForm"
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Private Declare Function ReleaseCapture Lib "user32" () As Long
    Private Const WM_NCLBUTTONDOWN = &HA1
    Private Const HTCAPTION = 2


Public Sub FormDrag(frm As Form)
    ReleaseCapture
    Call SendMessage(frm.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

'-----------------------------------------
' Add this code to the form to be draged

'    If Button = 1 Then Call FormDrag(Me)
'-----------------------------------------
