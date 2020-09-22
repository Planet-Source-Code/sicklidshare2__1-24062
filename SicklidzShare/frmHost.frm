VERSION 5.00
Begin VB.Form frmHost 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Host Room"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   Icon            =   "frmHost.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      Begin VB.Frame Frame1 
         BackColor       =   &H80000012&
         Caption         =   "Hosting"
         ForeColor       =   &H8000000E&
         Height          =   2175
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   4695
         Begin VB.CommandButton cmdCancel 
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
            Height          =   375
            Left            =   840
            TabIndex        =   5
            Top             =   1680
            Width           =   1215
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
            Height          =   375
            Left            =   2520
            TabIndex        =   4
            Top             =   1680
            Width           =   1215
         End
         Begin VB.TextBox txtMaxPeeps 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   7
            Text            =   "20"
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtAlias 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            MaxLength       =   14
            TabIndex        =   2
            Top             =   720
            Width           =   3135
         End
         Begin VB.TextBox txtRoomName 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   1
            Top             =   360
            Width           =   3135
         End
         Begin VB.VScrollBar vsSpin 
            Height          =   255
            Left            =   1920
            Max             =   40
            TabIndex        =   3
            Top             =   1080
            Value           =   10
            Width           =   255
         End
         Begin VB.Label Label3 
            BackColor       =   &H80000012&
            Caption         =   "Max Participants:"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label2 
            BackColor       =   &H80000012&
            Caption         =   "Name/Alias:"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000012&
            Caption         =   "Room name:"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "frmHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Me.Hide

End Sub

Private Sub cmdOK_Click()
If Len(txtRoomName.Text) > 0 And Len(txtAlias.Text) > 0 Then

    'If SicklidServer.sckAccept.State = 2 Then
    '    SicklidServer.sckAccept.Close
    'End If
    
    
    If SicklidServer.sckClient.State = 8 Then
        SicklidServer.sckClient.Close
    End If

'MsgBox "Client Winsock State: " & SicklidServer.sckClient.State
'MsgBox SicklidServer.sckServer(n).State
'MsgBox "Accept Winsock State: " & SicklidServer.sckAccept.State

Me.Hide
    SicklidServer.Tag = txtAlias.Text
    SicklidServer.sckAccept.Tag = txtRoomName.Text
    SicklidServer.cmdConnect.Enabled = False
    SicklidServer.cmdSend.Enabled = True
    SicklidServer.cmdClrScreen.Enabled = True
    SicklidServer.cmdHost.Caption = "Close Room"
    SicklidServer.Caption = "Hosting: " & txtRoomName.Text
    SicklidServer.lstPeople.AddItem SicklidServer.Tag
    SicklidServer.txtIncoming.Text = ""
    
    If SicklidServer.sckAccept.State = 0 Then
        SicklidServer.sckAccept.LocalPort = 5150
        SicklidServer.sckAccept.Listen
        SicklidServer.lblStatus2.Caption = "Hosting Chatroom..."
        
        '# This notifies the user if there was an error and resets the connection:
        If err Then
          SicklidServer.lblStatus2.Caption = err.Description
        End If


    End If

    If SicklidServer.sckAccept.LocalPort <> 5150 Then
        'SicklidServer.sckAccept.LocalPort = 5150
        'SicklidServer.sckAccept.Listen
        SicklidServer.lblChatroom.Caption = "You are hosting the: " & SicklidServer.sckClient.Tag & " Chat Room."

        SicklidServer.txtIncoming.Text = "Status:" & vbTab & vbTab & "Hosting Room: " & txtRoomName.Text & vbCrLf
        'MsgBox "Accept Winsock State In IF-THEN: " & SicklidServer.sckAccept.State
        Exit Sub
    End If
    'SicklidServer.sckAccept.Listen
    SicklidServer.txtIncoming.Text = "Status:" & vbTab & vbTab & "Hosting Room: " & txtRoomName.Text & vbCrLf
    SicklidServer.lblChatroom.Caption = "You are hosting the: < " & txtRoomName.Text & " > Chat Room."

End If
End Sub

Private Sub Form_Load()
vsSpin.Value = 18
End Sub

Private Sub vsSpin_Change()
'just dont ask.
If Abs(40 - CStr(vsSpin.Value)) + 2 > 40 Then Exit Sub
txtMaxPeeps.Text = Abs(40 - CStr(vsSpin.Value)) + 2
End Sub
