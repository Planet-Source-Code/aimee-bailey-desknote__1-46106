Attribute VB_Name = "ontop"
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2


Public Sub AlwaysOnTop(myfrm As Form, SetOnTop As Boolean)


    If SetOnTop Then
        lFlag = HWND_TOPMOST
    Else
        lFlag = HWND_NOTOPMOST
    End If
 
    SetWindowPos myfrm.hWnd, lFlag, _
    myfrm.Left / Screen.TwipsPerPixelX, _
    myfrm.Top / Screen.TwipsPerPixelY, _
    myfrm.Width / Screen.TwipsPerPixelX, _
    myfrm.Height / Screen.TwipsPerPixelY, _
    SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub

