VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "swflash.ocx"
Begin VB.Form frmConnect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Connect"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3330
   Icon            =   "frmConnect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   3330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin VB.Frame Frame2 
         BackColor       =   &H80000012&
         Caption         =   "Connect"
         ForeColor       =   &H8000000E&
         Height          =   2055
         Left            =   120
         TabIndex        =   21
         Top             =   120
         Width           =   3015
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
            Left            =   240
            TabIndex        =   26
            Top             =   1440
            Width           =   1215
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "OK"
            Enabled         =   0   'False
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
            Left            =   1560
            TabIndex        =   25
            Top             =   1440
            Width           =   1215
         End
         Begin VB.TextBox txtAlias 
            Alignment       =   2  'Center
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
            Left            =   1200
            MaxLength       =   14
            TabIndex        =   24
            Top             =   1080
            Width           =   1575
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "Search"
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
            Left            =   1200
            TabIndex        =   23
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox txtHost 
            Alignment       =   2  'Center
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
            Left            =   1200
            TabIndex        =   22
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label2 
            BackColor       =   &H80000012&
            Caption         =   "Name/Alias:"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000012&
            Caption         =   "Connect to:"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   720
            Width           =   855
         End
      End
      Begin VB.Frame frmIPSniff 
         BackColor       =   &H80000012&
         Caption         =   "Port Sniffer"
         ForeColor       =   &H8000000E&
         Height          =   2775
         Left            =   120
         TabIndex        =   1
         Top             =   2640
         Visible         =   0   'False
         Width           =   7095
         Begin VB.CommandButton Command7 
            BackColor       =   &H80000015&
            DownPicture     =   "frmConnect.frx":0442
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2640
            MaskColor       =   &H80000015&
            Picture         =   "frmConnect.frx":3160
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   2160
            UseMaskColor    =   -1  'True
            Width           =   1815
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H80000012&
            Height          =   1935
            Left            =   4800
            TabIndex        =   12
            Top             =   720
            Width           =   2175
            Begin VB.CommandButton Command4 
               Caption         =   "Pause"
               Enabled         =   0   'False
               Height          =   255
               Left            =   1200
               TabIndex        =   15
               Top             =   600
               Width           =   855
            End
            Begin VB.ListBox List1 
               Height          =   840
               Left            =   120
               TabIndex        =   14
               Top             =   960
               Width           =   1935
            End
            Begin VB.TextBox Text4 
               Height          =   285
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   13
               Top             =   240
               Width           =   855
            End
            Begin MSWinsockLib.Winsock Winsock8 
               Left            =   1440
               Top             =   360
               _ExtentX        =   741
               _ExtentY        =   741
               _Version        =   393216
            End
            Begin MSWinsockLib.Winsock Winsock7 
               Left            =   960
               Top             =   360
               _ExtentX        =   741
               _ExtentY        =   741
               _Version        =   393216
            End
            Begin MSWinsockLib.Winsock Winsock6 
               Left            =   480
               Top             =   360
               _ExtentX        =   741
               _ExtentY        =   741
               _Version        =   393216
            End
            Begin MSWinsockLib.Winsock Winsock5 
               Left            =   0
               Top             =   360
               _ExtentX        =   741
               _ExtentY        =   741
               _Version        =   393216
            End
            Begin VB.Label Label10 
               BackColor       =   &H80000012&
               Caption         =   "Open ports:"
               ForeColor       =   &H8000000E&
               Height          =   255
               Left            =   120
               TabIndex        =   17
               Top             =   720
               Width           =   975
            End
            Begin VB.Label Label11 
               BackColor       =   &H80000012&
               Caption         =   "Scanning port:"
               ForeColor       =   &H8000000E&
               Height          =   255
               Left            =   120
               TabIndex        =   16
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H80000012&
            ForeColor       =   &H8000000E&
            Height          =   1935
            Left            =   120
            TabIndex        =   2
            Top             =   720
            Width           =   2055
            Begin VB.TextBox Text5 
               Height          =   285
               Left            =   1440
               TabIndex        =   7
               Text            =   "1"
               Top             =   1560
               Width           =   495
            End
            Begin VB.Timer Timer3 
               Enabled         =   0   'False
               Interval        =   500
               Left            =   120
               Top             =   840
            End
            Begin VB.CommandButton Command5 
               Caption         =   "Start"
               Enabled         =   0   'False
               Height          =   255
               Left            =   120
               TabIndex        =   6
               Top             =   1560
               Width           =   615
            End
            Begin VB.TextBox Text3 
               Height          =   285
               Left            =   1200
               TabIndex        =   5
               Text            =   "32000"
               Top             =   1200
               Width           =   735
            End
            Begin VB.TextBox Text2 
               Height          =   285
               Left            =   120
               TabIndex        =   4
               Text            =   "1"
               Top             =   1200
               Width           =   735
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Left            =   120
               TabIndex        =   3
               Top             =   480
               Width           =   1815
            End
            Begin MSWinsockLib.Winsock Winsock10 
               Left            =   480
               Top             =   480
               _ExtentX        =   741
               _ExtentY        =   741
               _Version        =   393216
            End
            Begin MSWinsockLib.Winsock Winsock9 
               Left            =   0
               Top             =   480
               _ExtentX        =   741
               _ExtentY        =   741
               _Version        =   393216
            End
            Begin MSWinsockLib.Winsock Winsock4 
               Left            =   1440
               Top             =   1200
               _ExtentX        =   741
               _ExtentY        =   741
               _Version        =   393216
            End
            Begin MSWinsockLib.Winsock Winsock3 
               Left            =   960
               Top             =   1200
               _ExtentX        =   741
               _ExtentY        =   741
               _Version        =   393216
            End
            Begin MSWinsockLib.Winsock Winsock2 
               Left            =   480
               Top             =   1200
               _ExtentX        =   741
               _ExtentY        =   741
               _Version        =   393216
            End
            Begin MSWinsockLib.Winsock Winsock1 
               Left            =   0
               Top             =   1200
               _ExtentX        =   741
               _ExtentY        =   741
               _Version        =   393216
            End
            Begin VB.Label Label12 
               Alignment       =   2  'Center
               BackColor       =   &H80000012&
               Caption         =   "Timeout:"
               ForeColor       =   &H8000000E&
               Height          =   255
               Left            =   720
               TabIndex        =   11
               Top             =   1605
               Width           =   735
            End
            Begin VB.Label Label13 
               Alignment       =   2  'Center
               Caption         =   "To"
               Height          =   255
               Left            =   840
               TabIndex        =   10
               Top             =   1245
               Width           =   375
            End
            Begin VB.Label Label14 
               BackColor       =   &H80000012&
               Caption         =   "Range of ports to scan:"
               ForeColor       =   &H8000000E&
               Height          =   255
               Left            =   120
               TabIndex        =   9
               Top             =   960
               Width           =   1815
            End
            Begin VB.Label Label15 
               BackColor       =   &H80000012&
               Caption         =   "IP Address to scan:"
               ForeColor       =   &H8000000E&
               Height          =   255
               Left            =   120
               TabIndex        =   8
               Top             =   240
               Width           =   1815
            End
         End
         Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash3 
            Height          =   1215
            Left            =   2640
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   840
            Width           =   1815
            _cx             =   22678657
            _cy             =   22677599
            Movie           =   "c:\sicksm.swf"
            Src             =   "c:\sicksm.swf"
            WMode           =   "Window"
            Play            =   -1  'True
            Loop            =   -1  'True
            Quality         =   "High"
            SAlign          =   ""
            Menu            =   -1  'True
            Base            =   ""
            Scale           =   "ShowAll"
            DeviceFont      =   0   'False
            EmbedMovie      =   -1  'True
            BGColor         =   ""
            SWRemote        =   ""
            Stacking        =   "below"
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            BackColor       =   &H80000012&
            Caption         =   "Scan a Remote Port!"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   375
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   6855
         End
      End
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'# **********************
'# ****PORTSNIFF VAR*****
'# **********************

Dim PortToScan As Integer
Dim PortToStopOn As Integer
Dim Port1 As Integer
Dim Port2 As Integer
Dim Port3 As Integer
Dim Port4 As Integer
Dim Port5 As Integer
Dim Port6 As Integer
Dim Port7 As Integer
Dim Port8 As Integer
Dim Port9 As Integer
Dim Port10 As Integer
Dim TimeOut1
Dim TimeOut2
Dim TimeOut3
Dim TimeOut4
Dim TimeOut5
Dim TimeOut6
Dim TimeOut7
Dim TimeOut8
Dim TimeOut9
Dim TimeOut10
Dim Paused
Dim FileText
Dim TimerThing

Private Sub cmdCancel_Click()
txtHost.Text = ""
Me.Hide
SicklidServer.lblStatus2.Caption = "Connection Terminated By User"
End Sub

'********************************************************
'IsAlpha - locate illegal charactes in the subject field
 
Public Function IsAlpha(strText As String) As Boolean
 
    Dim intLen As Integer
    Dim intCounter As Integer
    Dim blnAlpha As Boolean
    Dim strChar As String
    
    intLen = Len(strText)
    
    'Set the condition of our loop
    For intCounter = 1 To intLen
        strChar = Mid(strText, intCounter, 1)
        
        'Accept A through z"
        If strChar >= "A" And strChar <= "z" Then
            blnAlpha = True
        Else
            'Accept "spaces", "periods", and "commas"
            If strChar = " " Or strChar = "." Or strChar = "," Then
                blnAlpha = True
            'Reject everything else.
            Else
                blnAlpha = False
            End If
        End If
        
        'We caught an illegal character (not A-z, " ", ".", or ","
        If blnAlpha = False Then
            IsAlpha = False
            Exit Function
        End If
    Next intCounter
    IsAlpha = True
End Function
                      

Private Sub cmdOK_Click()
If IsAlpha(txtHost) Then
    MsgBox "Enter a valid IP address, no letters!!", vbExclamation, "Error"
Else
    If Len(txtHost) > 0 And Len(txtAlias) > 0 Then
        Me.Hide
        SicklidServer.cmdConnect.Caption = "Cancel"
        SicklidServer.cmdHost.Enabled = False
        SicklidServer.sckClient.Close
        SicklidServer.sckAccept.Close
        SicklidServer.sckClient.Connect txtHost.Text, 5150
        SicklidServer.txtIncoming.Text = "Status:" & vbTab & vbTab & "...Connecting...."
        SicklidServer.Tag = txtAlias.Text

    Else
        MsgBox "Fill in ALL the fields", vbExclamation, "Error"
    End If
End If
End Sub

Private Sub cmdSearch_Click()
'MsgBox "This feature is not active.", vbInformation, "Sorry!"
'# PORT SNIFFER PORTION OF PROGRAM
frmIPSniff.Visible = False
SicklidServer.TabSheet.Tab = 2
SicklidServer.TabSheet.SetFocus
Me.Hide
End Sub




Private Sub Command5_Click()
TimerThing = 0
Winsock1.Close
Winsock2.Close
Winsock3.Close
Winsock4.Close
Winsock5.Close
Winsock6.Close
Winsock7.Close
Winsock8.Close
Winsock9.Close
Winsock10.Close
If Text2.Text = "" Then
Text2.Text = "1"
End If
If Text3.Text = "" Then
Text3.Text = "32000"
End If
Winsock1.RemoteHost = Text1.Text
Winsock2.RemoteHost = Text1.Text
Winsock3.RemoteHost = Text1.Text
Winsock4.RemoteHost = Text1.Text
Winsock5.RemoteHost = Text1.Text
Winsock6.RemoteHost = Text1.Text
Winsock7.RemoteHost = Text1.Text
Winsock8.RemoteHost = Text1.Text
Winsock9.RemoteHost = Text1.Text
Winsock10.RemoteHost = Text1.Text
PortToScan = Text2.Text
PortToStopOn = Text3.Text
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text5.Enabled = False
If Text5.Text = "" Then
Text5.Text = "1"
End If
Timer3.Enabled = True
Command5.Enabled = False
If Paused <> 1 Then
List1.Clear
Text4.Text = "Starting..."
End If
Command4.Enabled = True
Paused = 0
If SicklidServer.mnuLogtoFile.Checked = True Then
Open "PortLog.txt" For Output As #1
Close #1
CopyFileText
Open "PortLog.txt" For Output As #2
Write #2, FileText & vbCrLf & vbCrLf & "The remote IP at " & Winsock1.RemoteHost & " has the following open ports: " & vbCrLf & vbCrLf
Close #2
End If
SicklidServer.mnuLogtoFile.Enabled = False
End Sub

Private Sub Command4_Click()
Winsock1.Close
Text1.Enabled = True
Text2.Enabled = True
Text2.Text = PortToScan - 1
Text3.Enabled = True
Text5.Enabled = True
Command5.Enabled = True
Text4.Text = "Paused"
Timer3.Enabled = False
Command4.Enabled = True
Command5.Enabled = True
Paused = 1
End Sub

Private Sub Text1_Change()
If Text1.Text <> "" Then
Command5.Enabled = True
Else
Command5.Enabled = False
End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Command5_Click
    End If
End Sub

Private Sub Text2_Click()
Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)
End Sub

Private Sub Text3_Click()
Text3.SelStart = 0
Text3.SelLength = Len(Text3.Text)
End Sub

Private Sub Text5_Change()
If Text5.Text <> "" Then
Timer1.Interval = Text5.Text
End If
End Sub

Private Sub Text5_Click()
Text5.SelStart = 0
Text5.SelLength = Len(Text5.Text)
End Sub

Private Sub Timer3_Timer()
Text4.Text = PortToScan
TimerThing = TimerThing + 1

If TimerThing = 10 Then
 If TimeOut1 = 1 Then
  If Winsock1.State = sckConnected Then
  Winsock1.Close
  List1.AddItem Port1
   If SicklidServer.mnuLogtoFile.Checked = True Then
   CopyFileText
   Open "PortLog.txt" For Output As #3
   Write #3, FileText & Port1 & vbCrLf
   Close #3
   End If
  Beep
  TimeOut1 = 0
  PortToScan = PortToScan + 1
  Else
  Winsock1.Close
  TimeOut1 = 0
  PortToScan = PortToScan + 1
  End If
 End If
End If

If TimerThing = 1 Then
 If TimeOut1 <> 1 Then
  If PortToScan <= PortToStopOn Then
  Winsock1.RemotePort = PortToScan
  Port1 = PortToScan
  Winsock1.Connect
  TimeOut1 = 1
  Else
  Text1.Enabled = True
  Text2.Enabled = True
  Text3.Enabled = True
  Text5.Enabled = True
  Command5.Enabled = True
  Text4.Text = "Done"
  Command4.Enabled = False
  Timer3.Enabled = False
  SicklidServer.mnuLogtoFile.Enabled = True
  End If
 End If
End If

'

If TimerThing = 1 Then
 If TimeOut2 = 1 Then
  If Winsock2.State = sckConnected Then
  Winsock2.Close
  List1.AddItem Port2
   If SicklidServer.mnuLogtoFile.Checked = True Then
   CopyFileText
   Open "PortLog.txt" For Output As #3
   Write #3, FileText & Port2 & vbCrLf
   Close #3
   End If
  Beep
  TimeOut2 = 0
  PortToScan = PortToScan + 1
  Else
  Winsock2.Close
  TimeOut2 = 0
  PortToScan = PortToScan + 1
  End If
 End If
End If

If TimerThing = 2 Then
 If TimeOut2 <> 1 Then
  If PortToScan <= PortToStopOn Then
  Winsock2.RemotePort = PortToScan
  Port2 = PortToScan
  Winsock2.Connect
  TimeOut2 = 1
  Else
  Text1.Enabled = True
  Text2.Enabled = True
  Text3.Enabled = True
  Text5.Enabled = True
  Command5.Enabled = True
  Text4.Text = "Done"
  Command4.Enabled = False
  Timer3.Enabled = False
  SicklidServer.mnuLogtoFile.Enabled = True
  End If
 End If
End If

'

If TimerThing = 2 Then
 If TimeOut3 = 1 Then
  If Winsock3.State = sckConnected Then
  Winsock3.Close
  List1.AddItem Port3
   If SicklidServer.mnuLogtoFile.Checked = True Then
   CopyFileText
   Open "PortLog.txt" For Output As #3
   Write #3, FileText & Port3 & vbCrLf
   Close #3
   End If
  Beep
  TimeOut3 = 0
  PortToScan = PortToScan + 1
  Else
  Winsock3.Close
  TimeOut3 = 0
  PortToScan = PortToScan + 1
  End If
 End If
End If

If TimerThing = 3 Then
 If TimeOut3 <> 1 Then
  If PortToScan <= PortToStopOn Then
  Winsock3.RemotePort = PortToScan
  Port3 = PortToScan
  Winsock3.Connect
  TimeOut3 = 1
  Else
  Text1.Enabled = True
  Text2.Enabled = True
  Text3.Enabled = True
  Text5.Enabled = True
  Command5.Enabled = True
  Text4.Text = "Done"
  Command4.Enabled = False
  Timer3.Enabled = False
  SicklidServer.mnuLogtoFile.Enabled = True
  End If
 End If
End If

'

If TimerThing = 3 Then
 If TimeOut4 = 1 Then
  If Winsock4.State = sckConnected Then
  Winsock4.Close
  List1.AddItem Port4
   If SicklidServer.mnuLogtoFile.Checked = True Then
   CopyFileText
   Open "PortLog.txt" For Output As #3
   Write #3, FileText & Port4 & vbCrLf
   Close #3
   End If
  Beep
  TimeOut4 = 0
  PortToScan = PortToScan + 1
  Else
  Winsock4.Close
  TimeOut4 = 0
  PortToScan = PortToScan + 1
  End If
 End If
End If

If TimerThing = 4 Then
 If TimeOut4 <> 1 Then
  If PortToScan <= PortToStopOn Then
  Winsock4.RemotePort = PortToScan
  Port4 = PortToScan
  Winsock4.Connect
  TimeOut4 = 1
  Else
  Text1.Enabled = True
  Text2.Enabled = True
  Text3.Enabled = True
  Text5.Enabled = True
  Command5.Enabled = True
  Text4.Text = "Done"
  Command4.Enabled = False
  Timer3.Enabled = False
  SicklidServer.mnuLogtoFile.Enabled = True
  End If
 End If
End If

'

If TimerThing = 4 Then
 If TimeOut5 = 1 Then
  If Winsock5.State = sckConnected Then
  Winsock5.Close
  List1.AddItem Port5
   If SicklidServer.mnuLogtoFile.Checked = True Then
   CopyFileText
   Open "PortLog.txt" For Output As #3
   Write #3, FileText & Port5 & vbCrLf
   Close #3
   End If
  Beep
  TimeOut5 = 0
  PortToScan = PortToScan + 1
  Else
  Winsock5.Close
  TimeOut5 = 0
  PortToScan = PortToScan + 1
  End If
 End If
End If

If TimerThing = 5 Then
 If TimeOut5 <> 1 Then
  If PortToScan <= PortToStopOn Then
  Winsock5.RemotePort = PortToScan
  Port5 = PortToScan
  Winsock5.Connect
  TimeOut5 = 1
  Else
  Text1.Enabled = True
  Text2.Enabled = True
  Text3.Enabled = True
  Text5.Enabled = True
  Command5.Enabled = True
  Text4.Text = "Done"
  Command4.Enabled = False
  Timer3.Enabled = False
  SicklidServer.mnuLogtoFile.Enabled = True
  End If
 End If
End If

'

If TimerThing = 5 Then
 If TimeOut6 = 1 Then
  If Winsock6.State = sckConnected Then
  Winsock6.Close
  List1.AddItem Port6
   If SicklidServer.mnuLogtoFile.Checked = True Then
   CopyFileText
   Open "PortLog.txt" For Output As #3
   Write #3, FileText & Port6 & vbCrLf
   Close #3
   End If
  Beep
  TimeOut6 = 0
  PortToScan = PortToScan + 1
  Else
  Winsock6.Close
  TimeOut6 = 0
  PortToScan = PortToScan + 1
  End If
 End If
End If

If TimerThing = 6 Then
 If TimeOut6 <> 1 Then
  If PortToScan <= PortToStopOn Then
  Winsock6.RemotePort = PortToScan
  Port6 = PortToScan
  Winsock6.Connect
  TimeOut6 = 1
  Else
  Text1.Enabled = True
  Text2.Enabled = True
  Text3.Enabled = True
  Text5.Enabled = True
  Command5.Enabled = True
  Text4.Text = "Done"
  Command4.Enabled = False
  Timer3.Enabled = False
  SicklidServer.mnuLogtoFile.Enabled = True
  End If
 End If
End If

'

If TimerThing = 6 Then
 If TimeOut7 = 1 Then
  If Winsock7.State = sckConnected Then
  Winsock7.Close
  List1.AddItem Port7
   If SicklidServer.mnuLogtoFile.Checked = True Then
   CopyFileText
   Open "PortLog.txt" For Output As #3
   Write #3, FileText & Port7 & vbCrLf
   Close #3
   End If
  Beep
  TimeOut7 = 0
  PortToScan = PortToScan + 1
  Else
  Winsock7.Close
  TimeOut7 = 0
  PortToScan = PortToScan + 1
  End If
 End If
End If

If TimerThing = 7 Then
 If TimeOut7 <> 1 Then
  If PortToScan <= PortToStopOn Then
  Winsock7.RemotePort = PortToScan
  Port7 = PortToScan
  Winsock7.Connect
  TimeOut7 = 1
  Else
  Text1.Enabled = True
  Text2.Enabled = True
  Text3.Enabled = True
  Text5.Enabled = True
  Command5.Enabled = True
  Text4.Text = "Done"
  Command4.Enabled = False
  Timer3.Enabled = False
  SicklidServer.mnuLogtoFile.Enabled = True
  End If
 End If
End If

'

If TimerThing = 7 Then
 If TimeOut8 = 1 Then
  If Winsock8.State = sckConnected Then
  Winsock8.Close
  List1.AddItem Port8
   If SicklidServer.mnuLogtoFile.Checked = True Then
   CopyFileText
   Open "PortLog.txt" For Output As #3
   Write #3, FileText & Port8 & vbCrLf
   Close #3
   End If
  Beep
  TimeOut8 = 0
  PortToScan = PortToScan + 1
  Else
  Winsock8.Close
  TimeOut8 = 0
  PortToScan = PortToScan + 1
  End If
 End If
End If

If TimerThing = 8 Then
 If TimeOut8 <> 1 Then
  If PortToScan <= PortToStopOn Then
  Winsock8.RemotePort = PortToScan
  Port8 = PortToScan
  Winsock8.Connect
  TimeOut8 = 1
  Else
  Text1.Enabled = True
  Text2.Enabled = True
  Text3.Enabled = True
  Text5.Enabled = True
  Command5.Enabled = True
  Text4.Text = "Done"
  Command4.Enabled = False
  Timer3.Enabled = False
  SicklidServer.mnuLogtoFile.Enabled = True
  End If
 End If
End If

'

If TimerThing = 8 Then
 If TimeOut9 = 1 Then
  If Winsock9.State = sckConnected Then
  Winsock9.Close
  List1.AddItem Port9
   If SicklidServer.mnuLogtoFile.Checked = True Then
   CopyFileText
   Open "PortLog.txt" For Output As #3
   Write #3, FileText & Port9 & vbCrLf
   Close #3
   End If
  Beep
  TimeOut9 = 0
  PortToScan = PortToScan + 1
  Else
  Winsock9.Close
  TimeOut9 = 0
  PortToScan = PortToScan + 1
  End If
 End If
End If

If TimerThing = 9 Then
 If TimeOut9 <> 1 Then
  If PortToScan <= PortToStopOn Then
  Winsock9.RemotePort = PortToScan
  Port9 = PortToScan
  Winsock9.Connect
  TimeOut9 = 1
  Else
  Text1.Enabled = True
  Text2.Enabled = True
  Text3.Enabled = True
  Text5.Enabled = True
  Command5.Enabled = True
  Text4.Text = "Done"
  Command4.Enabled = False
  Timer3.Enabled = False
  SicklidServer.mnuLogtoFile.Enabled = True
  End If
 End If
End If

'

If TimerThing = 9 Then
 If TimeOut10 = 1 Then
  If Winsock10.State = sckConnected Then
  Winsock10.Close
  List1.AddItem Port10
   If SicklidServer.mnuLogtoFile.Checked = True Then
   CopyFileText
   Open "PortLog.txt" For Output As #3
   Write #3, FileText & Port10 & vbCrLf
   Close #3
   End If
  Beep
  TimeOut10 = 0
  PortToScan = PortToScan + 1
  Else
  Winsock10.Close
  TimeOut10 = 0
  PortToScan = PortToScan + 1
  End If
 End If
End If

If TimerThing = 10 Then
 If TimeOut10 <> 1 Then
  If PortToScan <= PortToStopOn Then
  Winsock10.RemotePort = PortToScan
  Port10 = PortToScan
  Winsock10.Connect
  TimeOut10 = 1
  TimerThing = 0
  Else
  Text1.Enabled = True
  Text2.Enabled = True
  Text3.Enabled = True
  Text5.Enabled = True
  Command5.Enabled = True
  Text4.Text = "Done"
  Command4.Enabled = False
  Timer3.Enabled = False
  SicklidServer.mnuLogtoFile.Enabled = True
  End If
 End If
End If
End Sub

Public Sub CopyFileText()
Open "PortLog.txt" For Binary As #4
Close #4
Open "PortLog.txt" For Input As #5
 If Not EOF(5) Then
 Input #5, FileText
 End If
Close #5
End Sub


Private Sub Form_Load()
Me.Caption = "Connect (Your IP: " & SicklidServer.sckAccept.LocalIP & ")"
End Sub

Private Sub txtHost_Change()
If txtHost.Text <> "" Then
    cmdOK.Enabled = True
Else
    cmdOK.Enabled = False
End If

End Sub
