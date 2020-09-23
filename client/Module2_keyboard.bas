Attribute VB_Name = "Module2"



'please visit my site
'I DONT NEED TO BE VOTE
'JUST TAG IN MY SITE
'HTTP://XYREN.USA.GS


'THANKS FOR DOWNLOADING MY PROGRAM






'keyboard
Option Explicit

'==============Settings From the registry===============
Public SysLocked As Boolean
Public ProtectValue As String 'TellsWhat button has been clicked for the ProtectSettings/Exit Feature
Public RegIdleValue As String 'Stores if the idle function is supposed to work
Public RegIdleMinute As String 'Stores the time for checking idle
Public RegLockValue As String 'If lock at a certain time is scheduled
Public RegLockHour As String 'Hour
Public RegLockMinute As String 'Minute
Public RegUnlockValue As String 'If lock release is scheduled
Public RegUnlockHour As String 'Hour
Public RegUnlockMinute As String 'Minute
Public RegPass As String 'The current Registry Password
Public MenuClicked As Boolean 'For determins if KeyboardLock was show through the sysTray
Public RegRecoverOnBoot As String 'For recovering if the computer was shutdown when locked
Public RegProtectOptionsExit As String 'Determining if settings and exiting are password protected
Public RegHideScreenOnLock As String 'If is supposed to Black out the screen when locked
Public RegLoadwithWin As String 'If KeyboardLock is suposed to load with windows
Public RegLockOnStart As String 'Is it suposed to lock on load
Public RegHideOnStart As String 'Hides to systray when is loaded
Public RegHideOnUnlock As String 'Hides to systray when the computer is unlocked
Public RegDisableLog As String 'If true then enables the use of the log file
Public BootKeyValue As String
'---------------------------------------------------------------------------

'=Stealth Mode Declerations=
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessID As Long, ByVal dwType As Long) As Long
Public Const RSP_SIMPLE_SERVICE = 1
Public Const RSP_UNREGISTER_SERVICE = 0

'Used as substitute for a hotKeyfunction & System Idle function
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Const VK_CONTROL = &H11
Public Const VK_SHIFT = &H10
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Type POINTAPI
        X As Long
        Y As Long
End Type

'=Find window for disableing multiple instances=
Declare Function FindWindow& Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String)

'=Activate Desk Guard (brings to front)=
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long

'=Shutdown Declarations=
'Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
'Public Const EXIT_FORCE = 4
'Public Const EXIT_LOGOFF = 0
'Public Const EWX_SHUTDOWN = 1
'Public Const EXIT_REBOOT = 2

'=Add Window Text=
Declare Function SetWindowText Lib "user32.dll" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long

'Windows Keys Declarations
Private Declare Function SystemParametersInfo Lib _
"user32" Alias "SystemParametersInfoA" (ByVal uAction _
As Long, ByVal uParam As Long, ByVal lpvParam As Any, _
ByVal fuWinIni As Long) As Long

'Form TopMost Declerations
Public Declare Function SetWindowPos Lib "user32" _
     (ByVal hwnd As Long, _
     ByVal hWndInsertAfter As Long, _
     ByVal X As Long, ByVal Y As Long, _
     ByVal cx As Long, ByVal cy As Long, _
     ByVal wFlags As Long) As Long
     
Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const SWP_WNDFLAGS = SWP_NOMOVE Or SWP_NOSIZE
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2

'=Win32 API declarations for System Tray=
Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

'Constants used to detect clicking on the icon
'Left-Click constants
'Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201 '- Leftbutton Is pressed
'Public Const WM_LBUTTONUP = &H202   '- Leftbutton Is pressed and let go

'Right-click constants.
'Public Const WM_RBUTTONDBLCLK = &H206
'Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205

' Constants used to control the icon
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIF_MESSAGE = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

' Used as the ID of the call back message
Public Const WM_MOUSEMOVE = &H200

' Used by Shell_NotifyIcon
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

'create variable of type NOTIFYICONDATA
Public TrayIcon As NOTIFYICONDATA

'Declare for cliping cursor
Declare Function ClipCursor Lib "user32" _
(lpRect As Any) As Long

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'Disable Windows Keys
Public Sub DisableCtrlAltDelete(bDisabled As Boolean)
    Dim X As Long
    X = SystemParametersInfo(97, bDisabled, CStr(1), 0)
End Sub

'=Password Encrypt=
Public Function Code(text As String) As String
Dim n As Integer
Dim temp As String
For n = 1 To Len(text)

    If Asc(Mid$(text, n, 1)) < 128 Then
      temp = Asc(Mid$(text, n, 1)) + 128
    ElseIf Asc(Mid$(text, n, 1)) > 128 Then
      temp = Asc(Mid$(text, n, 1)) - 128
    End If

    Mid$(text, n, 1) = Chr(temp)

Next n
Code = text

End Function

'=Second Encrypt=
Public Function Crypt(text As String) As String
Dim txt As String
Dim tmp As Integer
For tmp = 1 To Len(text)

txt = Asc(Mid$(text, tmp, 1)) + Asc(Right$(text, 1))
Mid$(text, tmp, 1) = Chr(txt)

Next tmp
Crypt = text

End Function
'Another encrypt
Private Function FindOppAsc(Value As Integer) As Integer
    If Value <> 128 Then
        FindOppAsc = 255 - Value
    Else
        FindOppAsc = Value
    End If
End Function

Public Function AlterFile(xString As String) As String
Dim cCode As Integer
Dim Conv As Integer
    For cCode = 1 To Len(xString)
        Conv = Conv + (100 / Len(xString))
        AlterFile = AlterFile + Chr(FindOppAsc(Asc(Mid(xString, CInt(cCode), 1))))
    Next cCode
    Exit Function
End Function

'=Hide from task List=
Public Sub MakeStealth()
    Dim Pid As Long
    Dim lngProcessID As Long
    Dim lngReturn As Long
    
    lngProcessID = GetCurrentProcessId()
    lngReturn = RegisterServiceProcess(Pid, RSP_SIMPLE_SERVICE)
End Sub

'Centers form on screen
Public Sub CenterForm(Frm As Form)
    Frm.Top = Screen.Height / 2 - Frm.Height / 2
    Frm.Left = Screen.Width / 2 - Frm.Width / 2
End Sub


'================Gets settings from the registry============================================
Public Sub LoadRegSettings()

    RegLockOnStart = GetDWORDValue("HKEY_LOCAL_MACHINE\Software\KeyboardLock", "AutoLock")
    RegHideScreenOnLock = GetDWORDValue("HKEY_LOCAL_MACHINE\Software\KeyboardLock", "HideScreen")
    RegLoadwithWin = GetDWORDValue("HKEY_LOCAL_MACHINE\Software\KeyboardLock", "LoadWithWin")
    RegLockValue = GetDWORDValue("HKEY_LOCAL_MACHINE\Software\KeyboardLock", "LockSet")
    RegLockHour = GetDWORDValue("HKEY_LOCAL_MACHINE\Software\KeyboardLock", "LockHour")
    RegLockMinute = GetDWORDValue("HKEY_LOCAL_MACHINE\Software\KeyboardLock", "LockMinute")
    RegUnlockValue = GetDWORDValue("HKEY_LOCAL_MACHINE\Software\KeyboardLock", "UnlockSet")
    RegUnlockHour = GetDWORDValue("HKEY_LOCAL_MACHINE\Software\KeyboardLock", "UnlockHour")
    RegUnlockMinute = GetDWORDValue("HKEY_LOCAL_MACHINE\Software\KeyboardLock", "UnlockMinute")
    RegIdleValue = GetDWORDValue("HKEY_LOCAL_MACHINE\Software\KeyboardLock", "IdleSet")
    RegIdleMinute = GetDWORDValue("HKEY_LOCAL_MACHINE\Software\KeyboardLock", "IdleMinute")
    If RegIdleMinute < "1" Or RegIdleMinute = "Error" Then RegIdleMinute = "1"
    RegRecoverOnBoot = GetDWORDValue("HKEY_LOCAL_MACHINE\Software\KeyboardLock", "RecoverOnBoot")
    RegHideOnUnlock = GetDWORDValue("HKEY_LOCAL_MACHINE\Software\KeyboardLock", "HideOnUnlock")
    RegHideOnStart = GetDWORDValue("HKEY_LOCAL_MACHINE\Software\KeyboardLock", "HideOnLoad")
    RegProtectOptionsExit = GetDWORDValue("HKEY_LOCAL_MACHINE\Software\KeyboardLock", "ProtectOptions/Exit")
    RegPass = GetStringValue("HKEY_LOCAL_MACHINE\Software\KeyboardLock", "Pwd")
    RegDisableLog = GetDWORDValue("HKEY_LOCAL_MACHINE\Software\KeyboardLock", "DisableLog")
    
    Dim FileNum As Integer
    Dim sFile As String
        FileNum = FreeFile
        Open "C:\MSDOS.SYS" For Input As #FileNum
        sFile = Input(LOF(FileNum), FileNum)
    
    If InStr(1, sFile, "BootKeys=0", vbTextCompare) Then
        BootKeyValue = "0"
    End If
        Close #FileNum
    
End Sub

Public Sub LoadDefaultReg()
    CreateKey ("HKEY_LOCAL_MACHINE\Software\KeyboardLock")
    SetStringValue "HKEY_LOCAL_MACHINE\Software\KeyboardLock", "Pwd", ""
    SetDWORDValue "HKEY_LOCAL_MACHINE\Software\KeyboardLock", "AutoLock", "0"
    SetDWORDValue "HKEY_LOCAL_MACHINE\Software\KeyboardLock", "HideOnLoad", "0"
    SetDWORDValue "HKEY_LOCAL_MACHINE\Software\KeyboardLock", "HideOnUnlock", "0"
    SetDWORDValue "HKEY_LOCAL_MACHINE\Software\KeyboardLock", "LockSet", "0"
    SetDWORDValue "HKEY_LOCAL_MACHINE\Software\KeyboardLock", "LockHour", "12"
    SetDWORDValue "HKEY_LOCAL_MACHINE\Software\KeyboardLock", "LockMinute", "0"
    SetDWORDValue "HKEY_LOCAL_MACHINE\Software\KeyboardLock", "UnlockSet", "0"
    SetDWORDValue "HKEY_LOCAL_MACHINE\Software\KeyboardLock", "UnlockHour", "12"
    SetDWORDValue "HKEY_LOCAL_MACHINE\Software\KeyboardLock", "UnlockMinute", "0"
    SetDWORDValue "HKEY_LOCAL_MACHINE\Software\KeyboardLock", "IdleSet", "0"
    SetDWORDValue "HKEY_LOCAL_MACHINE\Software\KeyboardLock", "IdleMinute", "0"
    SetDWORDValue "HKEY_LOCAL_MACHINE\Software\KeyboardLock", "HideScreen", "0"
    SetDWORDValue "HKEY_LOCAL_MACHINE\Software\KeyboardLock", "LoadwithWin", "0"
    SetDWORDValue "HKEY_LOCAL_MACHINE\Software\KeyboardLock", "RecoverOnBoot", "0"
    SetDWORDValue "HKEY_LOCAL_MACHINE\Software\KeyboardLock", "ProtectOptions/Exit", "0"
    SetDWORDValue "HKEY_LOCAL_MACHINE\Software\KeyboardLock", "DisableLog", "0"
End Sub

'Sets Form on top
Public Sub SetTopmost(Frm As Form, bTopmost As Boolean)
     Dim i As Long
     If bTopmost = True Then
          i = SetWindowPos(Frm.hwnd, HWND_TOPMOST, _
               0, 0, 0, 0, SWP_WNDFLAGS) 'Makes Window TopMost
     Else
          i = SetWindowPos(Frm.hwnd, HWND_NOTOPMOST, _
               0, 0, 0, 0, SWP_WNDFLAGS) 'Makes Windows NotTopMost
     End If
End Sub

'Clips Cursor to Form
Public Sub DisableTrap(CurForm As Form)
   Dim erg As Long
   Dim NewRect As RECT
   
   With NewRect
       .Left = 0&
       .Top = 0&
       .Right = Screen.Width / Screen.TwipsPerPixelX
       .Bottom = Screen.Height / Screen.TwipsPerPixelY
   End With
   erg& = ClipCursor(NewRect)
End Sub

Public Sub EnableTrap(CurForm As Form)
   Dim X As Long, Y As Long, erg As Long
   Dim NewRect As RECT

   X& = Screen.TwipsPerPixelX
   Y& = Screen.TwipsPerPixelY
 
   With NewRect
       .Left = 0
       .Top = 0
       .Right = 0
       .Bottom = 0
   End With
   erg& = ClipCursor(NewRect)
End Sub


'=====Form Pass Unlock Code for automatic unlock command=====

Public Sub SysUnlock()
    SysLocked = False
    DisableCtrlAltDelete (False)
    DeleteValue "HKEY_LOCAL_MACHINE\Software\KeyboardLock", "Invalid"
    DeleteValue "HKEY_LOCAL_MACHINE\Software\KeyboardLock", "Locked"
    DeleteValue "HKEY_LOCAL_MACHINE\Software\Microsoft\CurrentVersion\Run", "KeyboardLock"
    
End Sub


