VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H80000008&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   8295
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000008&
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   8025
      Begin VB.Timer Timer1 
         Interval        =   5000
         Left            =   1680
         Top             =   3000
      End
      Begin VB.Image imgLogo 
         Height          =   3705
         Left            =   120
         Picture         =   "frmSplash.frx":0000
         Stretch         =   -1  'True
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label lblCopyright 
         BackColor       =   &H00000000&
         Caption         =   "Copyright"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5520
         TabIndex        =   3
         Top             =   3540
         Width           =   2415
      End
      Begin VB.Label lblCompany 
         BackColor       =   &H80000008&
         Caption         =   "Company"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5520
         TabIndex        =   2
         Top             =   3750
         Width           =   2415
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000008&
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6900
         TabIndex        =   4
         Top             =   2700
         Width           =   915
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000008&
         Caption         =   "Platform"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6435
         TabIndex        =   5
         Top             =   2340
         Width           =   1380
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackColor       =   &H80000008&
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   780
         Left            =   3480
         TabIndex        =   7
         Top             =   1140
         Width           =   2550
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000008&
         Caption         =   "Licensed To: Sicklid"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3480
         TabIndex        =   1
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackColor       =   &H80000008&
         Caption         =   "CompanyProduct"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   3480
         TabIndex        =   6
         Top             =   705
         Width           =   3090
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public UserName As String
Private User2 As SicklidServer

Private Sub Form_KeyPress(KeyAscii As Integer)
        SicklidServer.UserName = UserName
        SicklidServer.Show

        Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
    lblCompanyProduct.Caption = App.ProductName
    lblCopyright.Caption = "Copyright: " & App.LegalCopyright
    lblCompany.Caption = "Company Name: " & App.CompanyName
    lblPlatform.Caption = "Best Used In Windows 2000"
    lblLicenseTo.Caption = "Licensed To: " & UserName
End Sub

Private Sub Frame1_Click()
        SicklidServer.UserName = UserName
        SicklidServer.Show

        Unload Me
End Sub


Private Sub imgLogo_Click()
        SicklidServer.UserName = UserName
        SicklidServer.Show

        Unload Me
End Sub

Private Sub lblCompanyProduct_Click()
        SicklidServer.UserName = UserName
        SicklidServer.Show

        Unload Me
End Sub

Private Sub lblLicenseTo_Click()
        SicklidServer.UserName = UserName
        SicklidServer.Show

        Unload Me
End Sub

Private Sub lblPlatform_Click()
        SicklidServer.UserName = UserName
        SicklidServer.Show

        Unload Me
End Sub

Private Sub lblProductName_Click()
        SicklidServer.UserName = UserName
        SicklidServer.Show

        Unload Me
End Sub

Private Sub lblVersion_Click()
        SicklidServer.UserName = UserName
        SicklidServer.Show

        Unload Me
        
        'Set User2 = New SicklidServer
        'User2.UserName = UserName
        'User2.Show

        'Unload Me
End Sub

Private Sub Timer1_Timer()
        'Set User2 = New SicklidServer
        SicklidServer.UserName = UserName
        SicklidServer.Show

        Unload Me
End Sub

