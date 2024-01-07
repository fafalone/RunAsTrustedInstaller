VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Run as ACTUAL administrator"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5220
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   5220
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   4515
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1875
      Width           =   5130
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Run As TrustedInstaller"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   360
      Left            =   2805
      TabIndex        =   2
      Top             =   1485
      Width           =   2370
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   30
      TabIndex        =   1
      Top             =   1095
      Width           =   5145
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":0000
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   30
      TabIndex        =   5
      Top             =   45
      Width           =   4725
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Must run as admin"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   210
      TabIndex        =   4
      Top             =   1530
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Program to run:"
      Height          =   255
      Left            =   30
      TabIndex        =   0
      Top             =   825
      Width           =   1230
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const BCM_FIRST = &H1600
Private Const BCM_SETSHIELD = (BCM_FIRST + &HC)
Private Const SB_BOTTOM = 7
Private Const EM_SCROLL As Integer = &HB5
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function IsUserAnAdmin Lib "shell32" () As Long

Private Sub Command1_Click()
Dim lRet As Long

lRet = LaunchAsTI(Text1.Text)

If lRet Then
    AppendLog "LaunchAsTI return code=0x" & Hex$(lRet) & " (SUCCESS)"
Else
    AppendLog "LaunchAsTI return code=0x" & Hex$(lRet) & " (FAIL)"
End If

End Sub

Private Sub Form_Load()
If IsUserAnAdmin() Then
    Command1.Enabled = True
    SendMessage Command1.hWnd, BCM_SETSHIELD, 0&, ByVal 1&
    AppendLog "Waiting..."
Else
    Label2.Visible = True
    AppendLog "Please exit and restart with 'Run As Administrator'"
End If
End Sub

Public Sub AppendLog(smsg As String)
Text2.Text = Text2.Text & smsg & vbCrLf
SendMessage Text2.hWnd, EM_SCROLL, SB_BOTTOM, ByVal 0&
End Sub

Private Sub Form_Unload(Cancel As Integer)
ReleaseToken
End Sub

