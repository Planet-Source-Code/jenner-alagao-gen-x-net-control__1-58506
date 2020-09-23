VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00EFF7F7&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SERVER"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4890
   ForeColor       =   &H00FFFFFF&
   Icon            =   "server_frm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   4890
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Log-off (force)"
      Height          =   375
      Left            =   840
      TabIndex        =   49
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   720
      Top             =   4320
   End
   Begin VB.Timer Timer4 
      Interval        =   100
      Left            =   2160
      Top             =   4320
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   240
      Top             =   4320
   End
   Begin VB.Timer Timer3 
      Interval        =   200
      Left            =   1200
      Top             =   4320
   End
   Begin VB.Timer Timer5 
      Interval        =   20
      Left            =   1680
      Top             =   4320
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   2640
      Top             =   4320
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   7223
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   15726583
      TabCaption(0)   =   "&Connection"
      TabPicture(0)   =   "server_frm.frx":0ECA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Power "
      TabPicture(1)   =   "server_frm.frx":0EE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Open Application"
      TabPicture(2)   =   "server_frm.frx":0F02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame14"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "&Accessories"
      TabPicture(3)   =   "server_frm.frx":0F1E
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Frame15"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame1 
         BackColor       =   &H00EFF7F7&
         BorderStyle     =   0  'None
         Height          =   3615
         Left            =   -74880
         TabIndex        =   25
         Top             =   360
         Width           =   4695
         Begin VB.CommandButton Command6 
            Caption         =   "Hibernate"
            Height          =   375
            Left            =   2520
            TabIndex        =   26
            Top             =   1800
            Width           =   1935
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Log-off"
            Height          =   375
            Left            =   2520
            TabIndex        =   50
            Top             =   1440
            Width           =   1935
         End
         Begin VB.CommandButton standby_cmd 
            Caption         =   "Standby"
            Height          =   375
            Left            =   2520
            TabIndex        =   30
            Top             =   1080
            Width           =   1935
         End
         Begin VB.CommandButton restart_cmd 
            Caption         =   "Restart"
            Height          =   375
            Left            =   2520
            TabIndex        =   32
            Top             =   720
            Width           =   1935
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Hibernate (force)"
            Height          =   375
            Left            =   3000
            TabIndex        =   27
            Top             =   3000
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Standby (force)"
            Height          =   375
            Left            =   3000
            TabIndex        =   29
            Top             =   2640
            Width           =   1455
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Restart (force)"
            Height          =   375
            Left            =   1560
            TabIndex        =   28
            Top             =   3000
            Width           =   1455
         End
         Begin VB.CommandButton shutdown_cmd 
            Caption         =   "Shutdown"
            Height          =   375
            Left            =   2520
            TabIndex        =   33
            Top             =   360
            Width           =   1935
         End
         Begin VB.CommandButton Force_shut 
            Caption         =   "Shutdown (Force)"
            Height          =   375
            Left            =   1560
            TabIndex        =   31
            Top             =   2640
            Width           =   1455
         End
         Begin VB.Image Image10 
            Height          =   480
            Index           =   1
            Left            =   840
            Picture         =   "server_frm.frx":0F3A
            Top             =   1560
            Width           =   480
         End
         Begin VB.Image Image10 
            Height          =   720
            Index           =   2
            Left            =   1200
            Picture         =   "server_frm.frx":1804
            Top             =   1440
            Width           =   720
         End
         Begin VB.Image Image11 
            Height          =   720
            Index           =   2
            Left            =   960
            Picture         =   "server_frm.frx":26CE
            Stretch         =   -1  'True
            Top             =   960
            Width           =   825
         End
         Begin VB.Image Image10 
            Height          =   480
            Index           =   3
            Left            =   1800
            Picture         =   "server_frm.frx":3598
            Top             =   1560
            Width           =   480
         End
         Begin VB.Image Image10 
            Height          =   480
            Index           =   0
            Left            =   360
            Picture         =   "server_frm.frx":3E62
            Top             =   1560
            Width           =   480
         End
         Begin VB.Image Image11 
            Height          =   720
            Index           =   0
            Left            =   600
            Picture         =   "server_frm.frx":472C
            Stretch         =   -1  'True
            Top             =   600
            Width           =   825
         End
         Begin VB.Image Image9 
            Height          =   720
            Index           =   1
            Left            =   840
            Picture         =   "server_frm.frx":55F6
            Top             =   2760
            Width           =   720
         End
         Begin VB.Image Image9 
            Height          =   480
            Index           =   0
            Left            =   720
            Picture         =   "server_frm.frx":64C0
            Top             =   2640
            Width           =   480
         End
         Begin VB.Image Image8 
            Height          =   720
            Left            =   240
            Picture         =   "server_frm.frx":6D8A
            Top             =   2640
            Width           =   720
         End
         Begin VB.Shape Shape5 
            BorderColor     =   &H00FF8000&
            FillColor       =   &H00CEF3FF&
            FillStyle       =   0  'Solid
            Height          =   975
            Left            =   120
            Top             =   2520
            Width           =   4455
         End
         Begin VB.Image Image11 
            Height          =   1080
            Index           =   1
            Left            =   1320
            Picture         =   "server_frm.frx":7C54
            Stretch         =   -1  'True
            Top             =   480
            Width           =   1185
         End
         Begin VB.Shape Shape7 
            BorderColor     =   &H00FF8000&
            FillColor       =   &H00CEF3FF&
            FillStyle       =   0  'Solid
            Height          =   2055
            Left            =   120
            Top             =   240
            Width           =   4455
         End
      End
      Begin VB.Frame Frame15 
         BackColor       =   &H00EFF7F7&
         BorderStyle     =   0  'None
         Height          =   3615
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   4695
         Begin VB.CheckBox minimize_chk 
            BackColor       =   &H00CEF3FF&
            Caption         =   "Minimize"
            Height          =   255
            Left            =   1680
            TabIndex        =   48
            Top             =   720
            Width           =   1215
         End
         Begin VB.CheckBox desktop_chk 
            BackColor       =   &H00CEF3FF&
            Caption         =   "Desktop Icon"
            Height          =   255
            Left            =   1680
            TabIndex        =   47
            Top             =   240
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.CheckBox taskbar_chk 
            BackColor       =   &H00CEF3FF&
            Caption         =   "Taskbar"
            Height          =   255
            Left            =   1680
            TabIndex        =   46
            Top             =   480
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.CheckBox ctrl_chk 
            BackColor       =   &H00CEF3FF&
            Caption         =   "Ctrl-Alt-Del"
            Height          =   255
            Left            =   2760
            TabIndex        =   45
            Top             =   1320
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.CheckBox block_chk 
            BackColor       =   &H00CEF3FF&
            Caption         =   "Block PC"
            Height          =   255
            Left            =   1440
            TabIndex        =   44
            Top             =   1320
            Width           =   1335
         End
         Begin VB.CheckBox mmove_chk 
            BackColor       =   &H00CEF3FF&
            Caption         =   "Control mouse move"
            Height          =   255
            Left            =   2640
            TabIndex        =   43
            Top             =   2640
            Width           =   1815
         End
         Begin VB.CheckBox mouse_chk 
            BackColor       =   &H00CEF3FF&
            Caption         =   "Mouse"
            Height          =   255
            Left            =   1320
            TabIndex        =   42
            Top             =   2640
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox sub_chk 
            BackColor       =   &H00CEF3FF&
            Caption         =   "Mouse func."
            Height          =   255
            Left            =   1320
            TabIndex        =   41
            Top             =   3120
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00CEF3FF&
            Caption         =   "Button Control"
            Enabled         =   0   'False
            Height          =   255
            Left            =   2640
            TabIndex        =   40
            Top             =   2880
            Width           =   1335
         End
         Begin VB.CheckBox keyboard_chk 
            BackColor       =   &H00CEF3FF&
            Caption         =   "Keyboard"
            Height          =   255
            Left            =   2760
            TabIndex        =   24
            Top             =   1560
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox swap_chk 
            BackColor       =   &H00CEF3FF&
            Caption         =   "Swap button"
            Height          =   255
            Left            =   1320
            TabIndex        =   23
            Top             =   2880
            Width           =   1215
         End
         Begin VB.CheckBox cd_chk 
            BackColor       =   &H00CEF3FF&
            Caption         =   "Open CD..."
            Height          =   255
            Left            =   1440
            TabIndex        =   22
            Top             =   1560
            Width           =   1215
         End
         Begin VB.CheckBox speaker_chk 
            BackColor       =   &H00CEF3FF&
            Caption         =   "Speaker"
            Enabled         =   0   'False
            Height          =   255
            Left            =   1440
            TabIndex        =   21
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Image Image7 
            Height          =   480
            Index           =   4
            Left            =   960
            Picture         =   "server_frm.frx":8B1E
            Top             =   480
            Width           =   480
         End
         Begin VB.Image Image7 
            Height          =   720
            Index           =   5
            Left            =   360
            Picture         =   "server_frm.frx":93E8
            Top             =   240
            Width           =   720
         End
         Begin VB.Image Image7 
            Height          =   360
            Index           =   2
            Left            =   960
            Picture         =   "server_frm.frx":A2B2
            Top             =   1800
            Width           =   360
         End
         Begin VB.Image Image7 
            Height          =   480
            Index           =   1
            Left            =   240
            Picture         =   "server_frm.frx":A99C
            Top             =   1680
            Width           =   480
         End
         Begin VB.Image Image7 
            Height          =   720
            Index           =   3
            Left            =   600
            Picture         =   "server_frm.frx":B266
            Top             =   1320
            Width           =   720
         End
         Begin VB.Image Image7 
            Height          =   720
            Index           =   0
            Left            =   360
            Picture         =   "server_frm.frx":C130
            Top             =   2520
            Width           =   720
         End
         Begin VB.Shape Shape3 
            BorderColor     =   &H00FF8000&
            FillColor       =   &H00CEF3FF&
            FillStyle       =   0  'Solid
            Height          =   1095
            Index           =   0
            Left            =   120
            Top             =   1200
            Width           =   4455
         End
         Begin VB.Shape Shape3 
            BorderColor     =   &H00FF8000&
            FillColor       =   &H00CEF3FF&
            FillStyle       =   0  'Solid
            Height          =   1095
            Index           =   1
            Left            =   120
            Top             =   2400
            Width           =   4455
         End
         Begin VB.Shape Shape3 
            BorderColor     =   &H00FF8000&
            FillColor       =   &H00CEF3FF&
            FillStyle       =   0  'Solid
            Height          =   975
            Index           =   2
            Left            =   120
            Top             =   120
            Width           =   4455
         End
      End
      Begin VB.Frame Frame14 
         BackColor       =   &H00EFF7F7&
         BorderStyle     =   0  'None
         Height          =   3615
         Left            =   -74880
         TabIndex        =   2
         Top             =   360
         Width           =   4695
         Begin VB.CommandButton sub_cmd 
            Caption         =   "Sub Option"
            Height          =   375
            Left            =   3000
            TabIndex        =   8
            Top             =   1200
            Width           =   1455
         End
         Begin VB.CommandButton findf_cmd 
            Caption         =   "Find Files"
            Height          =   375
            Left            =   3000
            TabIndex        =   5
            Top             =   840
            Width           =   1455
         End
         Begin VB.CommandButton run_cmd 
            Caption         =   "Run Files"
            Height          =   375
            Left            =   1560
            TabIndex        =   4
            Top             =   1200
            Width           =   1455
         End
         Begin VB.CommandButton dos_cmd 
            Caption         =   "Ms Dos"
            Height          =   375
            Left            =   1560
            TabIndex        =   7
            Top             =   840
            Width           =   1455
         End
         Begin VB.CommandButton explorer_cmd 
            Caption         =   "Explorer"
            Height          =   375
            Left            =   3000
            TabIndex        =   3
            Top             =   480
            Width           =   1455
         End
         Begin VB.CommandButton c_cmd 
            Caption         =   "Open C:\"
            Height          =   375
            Left            =   1560
            TabIndex        =   6
            Top             =   480
            Width           =   1455
         End
         Begin VB.CommandButton access_cmd 
            Caption         =   "Ms Access"
            Height          =   375
            Left            =   3000
            TabIndex        =   39
            Top             =   2880
            Width           =   1455
         End
         Begin VB.CommandButton word_cmd 
            Caption         =   "Ms Word"
            Height          =   375
            Left            =   3000
            TabIndex        =   36
            Top             =   2520
            Width           =   1455
         End
         Begin VB.CommandButton notepad_cmd 
            Caption         =   "Notepad"
            Height          =   375
            Left            =   1560
            TabIndex        =   35
            Top             =   2880
            Width           =   1455
         End
         Begin VB.CommandButton paint_cmd 
            Caption         =   "Paint"
            Height          =   375
            Left            =   1560
            TabIndex        =   34
            Top             =   2520
            Width           =   1455
         End
         Begin VB.CommandButton excel_cmd 
            Caption         =   "Ms Excel"
            Height          =   375
            Left            =   3000
            TabIndex        =   38
            Top             =   2160
            Width           =   1455
         End
         Begin VB.CommandButton calcu_cmd 
            Caption         =   "Calculator"
            Height          =   375
            Left            =   1560
            TabIndex        =   37
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Image Image6 
            Height          =   480
            Index           =   4
            Left            =   960
            Picture         =   "server_frm.frx":DDFA
            Top             =   1200
            Width           =   480
         End
         Begin VB.Image Image6 
            Height          =   720
            Index           =   7
            Left            =   240
            Picture         =   "server_frm.frx":E6C4
            Top             =   960
            Width           =   720
         End
         Begin VB.Image Image6 
            Height          =   720
            Index           =   5
            Left            =   600
            Picture         =   "server_frm.frx":1038E
            Top             =   2640
            Width           =   720
         End
         Begin VB.Image Image6 
            Height          =   720
            Index           =   2
            Left            =   240
            Picture         =   "server_frm.frx":11258
            Top             =   2400
            Width           =   720
         End
         Begin VB.Image Image6 
            Height          =   720
            Index           =   3
            Left            =   840
            Picture         =   "server_frm.frx":12122
            Top             =   2160
            Width           =   720
         End
         Begin VB.Image Image6 
            Height          =   720
            Index           =   0
            Left            =   840
            Picture         =   "server_frm.frx":12FEC
            Top             =   600
            Width           =   720
         End
         Begin VB.Image Image6 
            Height          =   720
            Index           =   1
            Left            =   360
            Picture         =   "server_frm.frx":13EB6
            Top             =   360
            Width           =   720
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FF8000&
            FillColor       =   &H00CEF3FF&
            FillStyle       =   0  'Solid
            Height          =   1455
            Left            =   120
            Top             =   2040
            Width           =   4455
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H00FF8000&
            FillColor       =   &H00CEF3FF&
            FillStyle       =   0  'Solid
            Height          =   1695
            Left            =   120
            Top             =   240
            Width           =   4455
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00EFF7F7&
         BorderStyle     =   0  'None
         Height          =   3615
         Left            =   -74880
         TabIndex        =   1
         Top             =   360
         Width           =   4695
         Begin VB.CommandButton clear_cmd 
            Caption         =   "Clear"
            Height          =   375
            Left            =   3120
            TabIndex        =   16
            Top             =   3000
            Width           =   1215
         End
         Begin VB.CommandButton send_cmd 
            Caption         =   "Send"
            Height          =   375
            Left            =   3120
            TabIndex        =   15
            Top             =   2640
            Width           =   1215
         End
         Begin VB.TextBox msg_txt 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   14
            Top             =   2640
            Width           =   2775
         End
         Begin VB.CommandButton disconnect_cmd 
            Caption         =   "&Disconnect"
            Height          =   375
            Left            =   3480
            TabIndex        =   13
            Top             =   1920
            Width           =   975
         End
         Begin VB.TextBox ip_txt 
            Height          =   285
            Left            =   1080
            TabIndex        =   12
            Text            =   "Unit3"
            Top             =   1560
            Width           =   2295
         End
         Begin VB.TextBox status_txt 
            Height          =   285
            Left            =   1080
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   11
            Top             =   1920
            Width           =   2295
         End
         Begin VB.CommandButton connect_cmd 
            BackColor       =   &H00000000&
            Caption         =   "Connect"
            Height          =   375
            Left            =   3480
            MaskColor       =   &H00000080&
            TabIndex        =   10
            Top             =   1440
            Width           =   975
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   800
            Picture         =   "server_frm.frx":14D80
            Top             =   360
            Width           =   480
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Status"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   19
            Top             =   1920
            Width           =   735
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "IP Address"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   18
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Shape Shape6 
            BorderColor     =   &H00FF8000&
            FillColor       =   &H00CEF3FF&
            FillStyle       =   0  'Solid
            Height          =   975
            Index           =   2
            Left            =   120
            Top             =   2520
            Width           =   4455
         End
         Begin VB.Image Image5 
            Height          =   480
            Left            =   2640
            Picture         =   "server_frm.frx":1564A
            Stretch         =   -1  'True
            Top             =   360
            Width           =   480
         End
         Begin VB.Image Image4 
            Height          =   720
            Left            =   1320
            Picture         =   "server_frm.frx":15BD4
            Top             =   240
            Width           =   720
         End
         Begin VB.Image Image3 
            Height          =   360
            Left            =   3240
            Picture         =   "server_frm.frx":16A9E
            Top             =   480
            Width           =   360
         End
         Begin VB.Image Image2 
            Height          =   720
            Left            =   1680
            Picture         =   "server_frm.frx":17208
            Stretch         =   -1  'True
            Top             =   360
            Width           =   960
         End
         Begin VB.Shape Shape6 
            BorderColor     =   &H00FF8000&
            FillColor       =   &H00CEF3FF&
            FillStyle       =   0  'Solid
            Height          =   1095
            Index           =   0
            Left            =   120
            Top             =   120
            Width           =   4455
         End
         Begin VB.Shape Shape6 
            BorderColor     =   &H00FF8000&
            FillColor       =   &H00CEF3FF&
            FillStyle       =   0  'Solid
            Height          =   1095
            Index           =   1
            Left            =   120
            Top             =   1320
            Width           =   4455
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H80000007&
            ForeColor       =   &H0000FF00&
            Height          =   375
            Left            =   960
            TabIndex        =   17
            Top             =   600
            Width           =   495
         End
      End
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   495
      Left            =   480
      TabIndex        =   9
      Top             =   1560
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      _Version        =   393216
      Max             =   65535
      TickStyle       =   3
      TextPosition    =   1
   End
   Begin MSWinsockLib.Winsock clijent 
      Left            =   3120
      Top             =   4320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'please visit my site
'I DONT NEED TO BE VOTE
'JUST TAG IN MY SITE
'HTTP://XYREN.USA.GS


'THANKS FOR DOWNLOADING MY PROGRAM




'mouse function
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Dim mouse As POINTAPI
Private Type POINTAPI
    x As Long
    y As Long
End Type
Dim xPos As Long
Dim yPos As Long
Dim cflag As Boolean

Dim control As Boolean

'open msaccess to client
Private Sub access_cmd_Click()
On Error Resume Next
clijent.SendData "access"
DoEvents
End Sub

'block the pc by black screen
Private Sub block_chk_Click()
If block_chk.Value = 0 Then
    On Error Resume Next
    clijent.SendData "unblack"
    DoEvents
ElseIf block_chk.Value = 1 Then
    On Error Resume Next
    msg_txt.Text = ""
    clijent.SendData "black"
    DoEvents
End If
End Sub

'open drive c:
Private Sub c_cmd_Click()
On Error Resume Next
clijent.SendData "c:\"
DoEvents
End Sub

'open calculator
Private Sub calcu_cmd_Click()
On Error Resume Next
clijent.SendData "calculator"
DoEvents
End Sub

'CD-ROM control
'open and close
Private Sub cd_chk_Click()
If cd_chk.Value = 0 Then
    clijent.SendData "cdc"
    DoEvents
ElseIf cd_chk.Value = 1 Then
    clijent.SendData "cdo"
    DoEvents
End If
End Sub


'Button control
'RIGHT CLICK  and LEFT CLICK
'but i terminate this program to client
Private Sub Check1_Click()
If Check1.Value = 0 Then
    Timer2.Enabled = False
    control = False
ElseIf Check1.Value = 1 Then
Timer2.Enabled = True
    control = True
End If
End Sub

'clear the textbox of message
Private Sub clear_cmd_Click()
msg_txt.Text = ""
End Sub

'standby force control
Private Sub Command1_Click()
On Error Resume Next
    clijent.SendData "standbyforce"
    DoEvents
End Sub


'restart force
Private Sub Command2_Click()
On Error Resume Next
clijent.SendData "restartforce"
DoEvents
End Sub

'command to log-ff the computer
Private Sub Command3_Click()
On Error Resume Next
    clijent.SendData "logoff"
    DoEvents
End Sub

'log-off by force
Private Sub Command4_Click()
On Error Resume Next
    clijent.SendData "logoffforce"
    DoEvents
End Sub


'hibernate its like a standby option
'BY FORCE
Private Sub Command5_Click()
On Error Resume Next
    clijent.SendData "Hibernateforce"
    DoEvents
End Sub

'hibernate
Private Sub Command6_Click()
On Error Resume Next
    clijent.SendData "Hibernate"
    DoEvents
End Sub

'Connect the winsock
Private Sub connect_cmd_Click()
On Error Resume Next
clijent.Close
clijent.Connect ip_txt.Text, 10003
End Sub


'CTRL-ALT-DEl control
'Enable and Disable
Private Sub ctrl_chk_Click()
If ctrl_chk.Value = 1 Then
    On Error Resume Next
    clijent.SendData "ena-del"
    DoEvents
ElseIf ctrl_chk.Value = 0 Then
    On Error Resume Next
    clijent.SendData "dis-del"
    DoEvents
End If
End Sub

'DESKTOP COntrol
'hide and show
Private Sub desktop_chk_Click()
If desktop_chk.Value = 0 Then
    On Error Resume Next
    clijent.SendData "hd"
    DoEvents
ElseIf desktop_chk.Value = 1 Then
    On Error Resume Next
    clijent.SendData "sd"
    DoEvents
End If
End Sub

'Disconnect the Client
Private Sub disconnect_cmd_Click()
If Label1.Caption = "7" Then
    If MsgBox("Are you sure you want to Disconnect Now ?", vbSystemModal + vbQuestion + vbYesNoCancel, "Disconnect Now") = vbYes Then
        On Error Resume Next
        clijent.SendData "disconnect"
        DoEvents
    
        clijent.Close
        status_txt.Text = "disconnected"
    End If
ElseIf (Label1.Caption = "8") Or (Label1.Caption = "9") Or (Label1.Caption = "6") Then
    MsgBox "No connection in other LAN ", vbApplicationModal + vbExclamation, "GEN-X Error"
Else
    MsgBox "There's no connection found", vbApplicationModal + vbExclamation, "GEN-X Error"
End If
End Sub

'open the MS-DOS-PROMPT
Private Sub dos_cmd_Click()
On Error Resume Next
clijent.SendData "dos"
DoEvents
End Sub

'OPEN THE MS EXCEL
Private Sub excel_cmd_Click()
On Error Resume Next
clijent.SendData "excel"
DoEvents
End Sub

'OPEN THE EXPLORER
'START WITHIN DRIVE C:
Private Sub explorer_cmd_Click()
On Error Resume Next
clijent.SendData "explorer"
DoEvents
End Sub

'FIND FILES
Private Sub findf_cmd_Click()
On Error Resume Next
clijent.SendData "findfiles"
DoEvents
End Sub

'FORCE SHUTDOWN
Private Sub Force_shut_Click()
On Error Resume Next
clijent.SendData "shutdownforce"
DoEvents
End Sub

Private Sub Form_Load()
'system tray
Me.Show 'form must be fully visible
    Me.Refresh
        
        With nid 'with system tray
            .cbSize = Len(nid)
            .hwnd = Me.hwnd
            .uId = vbNull
            .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
            .uCallBackMessage = WM_MOUSEMOVE
            .hIcon = Me.Icon 'use form's icon in tray
            .szTip = "System Tray Example" & vbNullChar 'tooltip text
        End With
        
    Shell_NotifyIcon NIM_ADD, nid 'add to tray



'=======CHANGE THE SYSTEM COLOR======
'Initialize variable

'i terminate it
'sys_color = "2"
'show_menu = "True"
'show_toolbar = "True"

'Use the save settings
'Call use_settings

'Get the original system color
'original_menu_color = GetSysColor(4)
'original_buttonface_color = GetSysColor(0)
'original_buttonshadow_color = GetSysColor(16)
'original_buttonhighlight_color = GetSysColor(20)

'Set the system color
'Call select_color_type(Val(sys_color))
'Slider1.Value = 30000

'status_txt.Text = "Not Connected"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Result, Action As Long
    
    'there are two display modes and we need to find out
    'which one the application is using
    
    If Me.ScaleMode = vbPixels Then
        Action = x
    Else
        Action = x / Screen.TwipsPerPixelX
    End If
    
'IF THE MOUSE CONTROL BUTTON WAS CHECKED THIS CODES WILL RUN
'DETECT EITHER LEFT CLICK OR RIGHT CLICK
If control = True Then
    If Button = vbKeyRButton Then
        On Error Resume Next
        clijent.SendData "right"
        DoEvents
    Else
        On Error Resume Next
        clijent.SendData "left"
        DoEvents
    End If
End If


'MOUSE ACTION
Select Case Action

    Case WM_LBUTTONDBLCLK 'Left Button Double Click
        Me.WindowState = vbNormal 'put into taskbar
            Result = SetForegroundWindow(Me.hwnd)
        Me.Show 'show form
    
    Case WM_RBUTTONUP 'Right Button Up
        Result = SetForegroundWindow(Me.hwnd)
        PopupMenu mnuFile 'popup menu, cool eh?
    
    End Select
e:
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
MsgBox "Thank you for using a GEN-X Net Control Program", vbInformation + vbSystemModal, "Quit GEN-X Net Control"
 Shell_NotifyIcon NIM_DELETE, nid 'remove from tray


'==========RESTORE THE SYSTEM COLOR======'
'terminate it
'Save the settings
'Call save_settings

'Restore the orignal system color
'New_System_Color.SelectColor(4) = original_menu_color
'New_System_Color.SelectColor(15) = original_buttonface_color
'New_System_Color.SelectColor(16) = original_buttonshadow_color
'New_System_Color.SelectColor(20) = original_buttonhighlight_color
'Call change_system_color

End
End Sub

'OPEN THE INFORATIOM ABOUT ME JENNER ALAGAO
Private Sub Image1_Click()
Form3.Show
End Sub
'OPEN THE INFORATIOM ABOUT ME JENNER ALAGAO
Private Sub Image2_Click()
Form3.Show
End Sub
'OPEN THE INFORATIOM ABOUT ME JENNER ALAGAO
Private Sub Image3_Click()
Form3.Show
End Sub
'OPEN THE INFORATIOM ABOUT ME JENNER ALAGAO
Private Sub Image4_Click()
Form3.Show
End Sub
'OPEN THE INFORATIOM ABOUT ME JENNER ALAGAO
Private Sub Image5_Click()
Form3.Show
End Sub


'IN IP TEXTBOX
'IF PRESS THE ENTER ITS
'AUTOMATICALLY GO TO THE CODES OF CONNECT BUTTON
Private Sub ip_txt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call connect_cmd_Click
End If
End Sub


'KEYBOARD CONTROL
'DISABLE OR ENABLE
'
'
'BUT THIS CODE IS NOT ACCESIBLE

Private Sub keyboard_chk_Click()
If keyboard_chk.Value = 0 Then
    On Error Resume Next
    clijent.SendData "keyunlock"
    DoEvents
ElseIf keyboard_chk.Value = 1 Then
    On Error Resume Next
    clijent.SendData "keylock"
    DoEvents
End If
End Sub


'GETMINIMIZE ALL APPLICATION
Private Sub minimize_chk_Click()
If minimize_chk.Value = 1 Then
    On Error Resume Next
    clijent.SendData "minimize"
    DoEvents
ElseIf minimize_chk.Value = 0 Then
    On Error Resume Next
    clijent.SendData "undominimize"
    DoEvents
End If
End Sub

'THE SERVER WILL CONTROL THE
'MOVEMENT OF CURSOR IN CLIENT
Private Sub mmove_chk_Click()
If mmove_chk.Value = 0 Then
    Timer2.Enabled = False
ElseIf mmove_chk.Value = 1 Then
    Timer2.Enabled = True
End If
End Sub


'SET THE MOUSE FUNCTION OR NOT
Private Sub mouse_chk_Click()
If mouse_chk.Value = 0 Then
    On Error Resume Next
    clijent.SendData "mouselock"
    DoEvents
ElseIf mouse_chk.Value = 1 Then
    On Error Resume Next
    clijent.SendData "mouseunlock"
    DoEvents
End If
End Sub

'OPEN THE NOTEPAD
Private Sub notepad_cmd_Click()
On Error Resume Next
clijent.SendData "notepad"
DoEvents
End Sub

'OPEN THE MSPAINT
Private Sub paint_cmd_Click()
On Error Resume Next
clijent.SendData "mspaint"
DoEvents
End Sub

'OPEN THR RUN WINDOW
Private Sub run_cmd_Click()
On Error Resume Next
clijent.SendData "run"
DoEvents
End Sub

'SEND THE MESSAGE TO CLIENT
Private Sub send_cmd_Click()
If msg_txt.Text = "" Then MsgBox "Theres no message to send", vbSystemModal + vbExclamation, "GEN-X Error"
On Error Resume Next
clijent.SendData msg_txt.Text
msg_txt.Text = ""
DoEvents
End Sub

'SHUTDOWN THE COMPUTER
Private Sub shutdown_cmd_Click()
On Error Resume Next
clijent.SendData "shutdown"
DoEvents
End Sub


'CHANGE THE VOLUME CONTROL
'BUT I TERMINATE THIS FUNCTION TO CLIENT
' TRY TO RUN THIS
Private Sub Slider1_Change()
On Error Resume Next
clijent.SendData "vol" & Slider1.Value
DoEvents
End Sub

'SPEAKER VOLUME WILL BECAME ZERO
'SO IT YTHINK ITS NO SPEAKER
Private Sub speaker_chk_Click()
If speaker_chk.Value = 0 Then
    Slider1.Value = 0
ElseIf speaker_chk.Value = 1 Then
    Slider1.Value = 30000
End If
End Sub

'RESTART CONTROL
Private Sub restart_cmd_Click()
On Error Resume Next
clijent.SendData "restart"
DoEvents
End Sub

'IF RIGHT ON TAB
'THEN THE MOUSE BUTTON CONTROL IS = TRUE
'THE CLIENT WILL BE RIGHT CLICK OR LEFT
'EITHER THE CLIENT DID NOT CLICK
'BUT I TERMINATE IT
Private Sub SSTab1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If control = True Then
    If Button = vbKeyRButton Then
        On Error Resume Next
        clijent.SendData "right"
        DoEvents
    Else
        On Error Resume Next
        clijent.SendData "left"
        DoEvents
    End If
End If
End Sub


'STANDBY CONTROL
Private Sub standby_cmd_Click()
  On Error Resume Next
    clijent.SendData "standby"
    DoEvents
End Sub

'MOUSE FUNCTION
Private Sub sub_chk_Click()
If sub_chk.Value = 0 Then
    On Error Resume Next
    clijent.SendData "mousefunc"
    DoEvents
ElseIf sub_chk.Value = 1 Then
    On Error Resume Next
    clijent.SendData "mousenofunc"
    DoEvents
End If
End Sub

'SWAP THE BUTTON FUNCTION OF A MOUSE
Private Sub swap_chk_Click()
If swap_chk.Value = 0 Then
    On Error Resume Next
    clijent.SendData "co"
    DoEvents
ElseIf swap_chk.Value = 1 Then
    On Error Resume Next
    clijent.SendData "c"
    DoEvents
End If
End Sub

'HIDE OR SHOW THE TASKBAR
Private Sub taskbar_chk_Click()
If taskbar_chk.Value = 0 Then
    On Error Resume Next
    clijent.SendData "ht"
    DoEvents
ElseIf taskbar_chk.Value = 1 Then
    On Error Resume Next
    clijent.SendData "st"
    DoEvents
End If
End Sub


'CHECKING OF CONNECTION
Private Sub Timer1_Timer()
On Error Resume Next
Label1.Caption = clijent.State
If Label1.Caption = "6" Then
    status_txt.Text = "connecting"
ElseIf Label1.Caption = "9" Then
    status_txt.Text = "not connected"
ElseIf Label1.Caption = "7" Then
    status_txt.Text = "connected"
ElseIf Label1.Caption = "8" Then
    status_txt.Text = "connection terminated"
End If
End Sub

'GETTING THE MOUSE CURSOR POSITION
Private Sub Timer2_Timer()
Call GetCursorPos(mouse)
x = mouse.x
y = mouse.y
Dim str(2), a As Long
str(0) = x
str(1) = y
    On Error Resume Next
    a = Len("mousepos" & str(0) & "-" & str(1))
    clijent.SendData "mousepos" & a & "-" & str(0) & "-" & str(1)
    DoEvents
Exit Sub
End Sub

'GETTING THE LEFT CLICK OR RIGHT OF SERVER
'TRY TO UNDERSTAND
'ITS FOR SYSTEM ICON TRAY BUT NOT EXECUTED CODES
Private Sub Timer4_Timer()
Static Sant As Boolean
If Sant <> True Then
        If GetX < 54 And GetY > 576 Then LeftClick: Sant = True
    Else
        
        If GetX > 54 And GetY < 576 Then Sant = False
End If
End Sub


'MOVING THE ICONS IN LEFT GOING TO RIGHT
'PANG PAGANDA LANG PERO HINDI NAMAN MAGANDA
Private Sub Timer5_Timer()
Dim a, b, c, d, e As Long
a = Image1.Left
b = Image2.Left
c = Image3.Left
d = Image4.Left
e = Image5.Left
If a >= 1500 Then
    Timer6.Enabled = True
    Timer5.Enabled = False
Else
    Image1.Move a + 20
    Image2.Move b + 20
    Image3.Move c + 20
    Image4.Move d + 20
    Image5.Move e + 20
End If
Label1.Caption = a
End Sub

'MOVING THE ICONS IN RIGHT GOING TO LEFT
'PANG PAGANDA LANG PERO HINDI NAMAN MAGANDA
Private Sub Timer6_Timer()
Dim a, b, c, d, e As Long
a = Image1.Left
b = Image2.Left
c = Image3.Left
d = Image4.Left
e = Image5.Left
If a <= 800 Then
    Timer5.Enabled = True
    Timer6.Enabled = False
Else
    Image1.Move a - 20
    Image2.Move b - 20
    Image3.Move c - 20
    Image4.Move d - 20
    Image5.Move e - 20
End If
End Sub

'OPEN THE MS WORD
Private Sub word_cmd_Click()
On Error Resume Next
clijent.SendData "word"
DoEvents
End Sub
