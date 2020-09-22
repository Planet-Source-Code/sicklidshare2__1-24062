VERSION 5.00
Begin VB.Form Login 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SickidSoft's ChatShare Login"
   ClientHeight    =   1125
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   664.687
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   960
      TabIndex        =   1
      Top             =   135
      Width           =   2685
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   960
      TabIndex        =   2
      Top             =   660
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2520
      TabIndex        =   3
      Top             =   660
      Width           =   1140
   End
   Begin VB.Image Image1 
      Height          =   345
      Left            =   2160
      Picture         =   "frmLogin.frx":0000
      Top             =   720
      Width           =   315
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H80000007&
      Caption         =   "&UserName:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean
Public UserName As String

Private USER As frmSplash
Private User2 As SicklidServer

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    MsgBox "Fine, Screw Off Then!", vbOKOnly, "BYE!"
    LoginSucceeded = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    'check for correct password
    UserName = txtUserName.Text
    
    If txtUserName = "" Then
        MsgBox "Please enter in your Username for this Session!", , "Startup"
        txtUserName.SetFocus
        SendKeys "{Home}+{End}"
    Else
        LoginSucceeded = True
        Set USER = New frmSplash
        USER.UserName = txtUserName.Text
        USER.Show
        Unload Me
    End If
End Sub

