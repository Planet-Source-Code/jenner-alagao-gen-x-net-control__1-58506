VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Information"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   Icon            =   "server_info.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&OK"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   2
      Left            =   360
      Picture         =   "server_info.frx":08CA
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   720
      Index           =   1
      Left            =   480
      Picture         =   "server_info.frx":1194
      Top             =   2760
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   360
      Picture         =   "server_info.frx":205E
      Top             =   480
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "The GEN-X Net Control Program is created by JENNER F. ALAGAO that access the network Lan Computer."
      Height          =   615
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   480
      Width           =   3135
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00FF8000&
      FillColor       =   &H00CEF3FF&
      FillStyle       =   0  'Solid
      Height          =   1095
      Index           =   0
      Left            =   120
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "http://xyren.usa.gs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   1800
      TabIndex        =   5
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "jhonce_xyren@yahoo.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   1800
      TabIndex        =   4
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "GENERATION XYREN NET CONTROL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "I made this program to play the server of computer rental shop "
      Height          =   615
      Index           =   1
      Left            =   960
      TabIndex        =   2
      Top             =   1920
      Width           =   3255
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00FF8000&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   1
      Left            =   120
      Top             =   1440
      Width           =   4335
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00FF8000&
      FillColor       =   &H00CEF3FF&
      FillStyle       =   0  'Solid
      Height          =   1095
      Index           =   2
      Left            =   120
      Top             =   2640
      Width           =   4335
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00FF8000&
      FillColor       =   &H00CEF3FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   3
      Left            =   120
      Top             =   1800
      Width           =   4335
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'please visit my site
'I DONT NEED TO BE VOTE
'JUST TAG IN MY SITE
'HTTP://XYREN.USA.GS


'THANKS FOR DOWNLOADING MY PROGRAM




Private Sub Command1_Click()
Form1.Show
Unload Me
End Sub

