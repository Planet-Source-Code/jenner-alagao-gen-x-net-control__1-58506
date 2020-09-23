Attribute VB_Name = "position"
'set the form top
Public username As String

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
'Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
 
Public Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal X As Long, ByVal wFlags As Long)
