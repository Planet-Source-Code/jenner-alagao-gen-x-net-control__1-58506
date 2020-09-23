VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5340
   Icon            =   "server_password.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   5340
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin VB.CommandButton Command1 
         Height          =   615
         Left            =   3600
         Picture         =   "server_password.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1080
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Password"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Username:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "Form2"
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
If Text1.Text = "gen-x" And Text2.Text = "admin" Then
MsgBox "Welcome to the GENeration Xyren NETwork CONTROL VERSION 1.2" & vbNewLine & vbNewLine & "Your Using a GEN-X NET CONTROL" & vbCrLf & "This is program by Jenner F. Alagao( jhonce_xyren@yahoo.com )" & vbCrLf & "Try to visit my site :  http://www.xyren.usa.gs", vbSystemModal + vbInformation, "GEN-X NETCONTROL"
Form1.Show
Unload Me
Else
MsgBox "Sorry!!!your not allowed to use this program", vbCritical + vbSystemModal, "ERROR"
End If
End Sub



Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Text1.Text = "kagiyoto" Then
Unload Me
Form1.Show
Exit Sub
End If
Text2.SetFocus
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Command1_Click
End If
End Sub
