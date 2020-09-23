Attribute VB_Name = "system_color"
'please visit my site
'I DONT NEED TO BE VOTE
'JUST TAG IN MY SITE
'HTTP://XYREN.USA.GS


'THANKS FOR DOWNLOADING MY PROGRAM
  
Option Explicit

'For Settings
Global sys_color        As String
Global show_menu        As String
Global show_toolbar     As String

Global original_menu_color              As Long
Global original_buttonface_color        As Long
Global original_buttonshadow_color      As Long
Global original_buttonhighlight_color      As Long

Public Declare Function SetSysColors Lib "user32" (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Const COLOR_WINDOW = 5
Public Const COLOR_WINDOWFRAME = 6
Public Const COLOR_MENUTEXT = 7
Public Const COLOR_WINDOWTEXT = 8
Public Const COLOR_CAPTIONTEXT = 9
Public Const COLOR_ACTIVEBORDER = 10
Public Const COLOR_SCROLLBAR = 0
Public Const COLOR_INACTIVECAPTION = 3
Public Const COLOR_MENU = 4
Public Const COLOR_BACKGROUND = 1
Public Const COLOR_ACTIVECAPTION = 2
Public Const COLOR_HIGHLIGHT = 13
Public Const COLOR_HIGHLIGHTTEXT = 14
Public Const COLOR_BTNFACE = 15
Public Const COLOR_BTNSHADOW = 16
Public Const COLOR_GRAYTEXT = 17
Public Const COLOR_BTNTEXT = 18
Public Const COLOR_INACTIVECAPTIONTEXT = 19
Public Const COLOR_BTNHIGHLIGHT = 20
Public Const COLOR_2NDACTIVECAPTION = 27
Public Const COLOR_2NDINACTIVECAPTION = 28
Public Const COLOR_INACTIVEBORDER = 11
Public Const COLOR_APPWORKSPACE = 12


Public Type ColorSystem
    SelectColor(0 To 20) As Long
End Type

Public New_System_Color As ColorSystem

Public Sub use_settings()
On Error Resume Next
Open parent_dir(Environ("TMP")) & "\GEN-X\Settings\Settings.bin" For Input As #1
    Input #1, sys_color
    Input #1, show_menu
    Input #1, show_toolbar
Close #1
End Sub

Public Sub select_color_type(ByVal sColorOption As Byte)
Select Case sColorOption
    Case 0: '[ XP Default ]
            New_System_Color.SelectColor(4) = RGB(239, 238, 224)  'Menu
            New_System_Color.SelectColor(15) = RGB(240, 240, 224) 'Button
            New_System_Color.SelectColor(16) = RGB(216, 210, 189) 'Button Shadow
            New_System_Color.SelectColor(20) = RGB(255, 255, 255) 'Button Highlight

            Call change_system_color
            
    Case 1: '[ Mac Grey ]
            New_System_Color.SelectColor(4) = RGB(235, 235, 235)  'Menu
            New_System_Color.SelectColor(15) = RGB(235, 235, 235) 'Button
            New_System_Color.SelectColor(16) = RGB(186, 186, 186) 'Button Shadow
            New_System_Color.SelectColor(20) = RGB(255, 255, 255) 'Button Highlight
            
            Call change_system_color
            
    Case 2: '[ XP Blue ]
            New_System_Color.SelectColor(4) = RGB(211, 229, 251)    'Menu
            New_System_Color.SelectColor(15) = RGB(211, 229, 251) 'Button
            New_System_Color.SelectColor(16) = RGB(139, 188, 254)  'Button Shadow
            New_System_Color.SelectColor(20) = RGB(255, 255, 255) 'Button Highlight
            
            Call change_system_color
    
    Case 3: '[ Cool Green ]
            New_System_Color.SelectColor(4) = RGB(217, 238, 205)   'Menu
            New_System_Color.SelectColor(15) = RGB(217, 238, 205) 'Button
            New_System_Color.SelectColor(16) = RGB(149, 207, 114)  'Button Shadow
            New_System_Color.SelectColor(20) = RGB(255, 255, 255) 'Button Highlight
            
            Call change_system_color
            
    Case 4: '[ Light Violet ]
            New_System_Color.SelectColor(4) = RGB(220, 220, 223)  'Menu
            New_System_Color.SelectColor(15) = RGB(220, 220, 223) 'Button
            New_System_Color.SelectColor(16) = RGB(185, 191, 199)  'Button Shadow
            New_System_Color.SelectColor(20) = RGB(235, 244, 255) 'Button Highlight
            
            Call change_system_color
            
    Case 5: '[ Light Brown ]
            New_System_Color.SelectColor(4) = RGB(218, 214, 206)   'Menu
            New_System_Color.SelectColor(15) = RGB(218, 214, 206) 'Button
            New_System_Color.SelectColor(16) = RGB(167, 163, 155)  'Button Shadow
            New_System_Color.SelectColor(20) = RGB(235, 231, 223)  'Button Highlight
            
            Call change_system_color
        
    Case 6: '[ Win Classic ]
            New_System_Color.SelectColor(4) = RGB(212, 208, 200)    'Menu
            New_System_Color.SelectColor(15) = RGB(212, 208, 200) 'Button
            New_System_Color.SelectColor(16) = RGB(128, 128, 128) 'Button Shadow
            New_System_Color.SelectColor(20) = RGB(255, 255, 255)  'Button Highlight
            
            Call change_system_color
            
End Select
End Sub
Private Sub create_save_setting_dir()
On Error Resume Next
MkDir parent_dir(Environ("TMP")) & "\GEN-X"
MkDir parent_dir(Environ("TMP")) & "\GEN-X\Settings"
End Sub
Public Function parent_dir(ByVal sDir As String) As String
Dim c As Integer
Dim tmp_h As String
For c = 1 To Len(sDir)
    tmp_h = Left(Right(sDir, c), 1)
    If tmp_h = "\" Then Exit For
    parent_dir = Left(sDir, Len(sDir) - (c + 1))
Next c
tmp_h = ""
c = 0
End Function
Public Sub change_system_color()

Call SetSysColors(1, 4, New_System_Color.SelectColor(4))   'Menu
Call SetSysColors(1, 15, New_System_Color.SelectColor(15)) 'Button
Call SetSysColors(1, 16, New_System_Color.SelectColor(16)) 'Button Shadow
Call SetSysColors(1, 20, New_System_Color.SelectColor(20)) 'Button Highlight

End Sub

Public Sub save_settings()
On Error Resume Next
Call create_save_setting_dir
Open parent_dir(Environ("TMP")) & "\Media Tracker\Settings\Settings.bin" For Output As #1
    Print #1, sys_color
    Print #1, show_menu
    Print #1, show_toolbar
Close #1
End Sub
