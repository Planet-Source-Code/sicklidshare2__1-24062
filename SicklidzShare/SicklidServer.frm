VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "swflash.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form SicklidServer 
   BackColor       =   &H00000000&
   Caption         =   "A Sicklidz ChatShare Program"
   ClientHeight    =   8415
   ClientLeft      =   4020
   ClientTop       =   2130
   ClientWidth     =   7335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SicklidServer.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8415
   ScaleWidth      =   7335
   Begin TabDlg.SSTab TabSheet 
      Height          =   9135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   16113
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   582
      BackColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "File Sharing"
      TabPicture(0)   =   "SicklidServer.frx":4F0A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Line2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Winsock_Send"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Command2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Btn_Listen"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Btn_Send"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Fra_Advanced"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "FraServer"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Btn_Browse"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Private Chat"
      TabPicture(1)   =   "SicklidServer.frx":4F26
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Fun Stuff"
      TabPicture(2)   =   "SicklidServer.frx":4F42
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame7"
      Tab(2).Control(1)=   "frmIPSniff"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Chatroom"
      TabPicture(3)   =   "SicklidServer.frx":4F5E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame8"
      Tab(3).Control(1)=   "Chatroom"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "MultiMedia"
      TabPicture(4)   =   "SicklidServer.frx":4F7A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame9"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Frame10"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).ControlCount=   2
      Begin VB.Frame Frame10 
         BackColor       =   &H80000007&
         Caption         =   "MP3 Player"
         ForeColor       =   &H8000000E&
         Height          =   7815
         Left            =   -74880
         TabIndex        =   124
         Top             =   480
         Width           =   7095
         Begin VB.CommandButton Command15 
            BackColor       =   &H80000015&
            DownPicture     =   "SicklidServer.frx":4F96
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
            Left            =   4440
            MaskColor       =   &H80000015&
            Picture         =   "SicklidServer.frx":7CB4
            Style           =   1  'Graphical
            TabIndex        =   141
            Top             =   1440
            UseMaskColor    =   -1  'True
            Width           =   1815
         End
         Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash4 
            Height          =   1215
            Left            =   4440
            TabIndex        =   142
            TabStop         =   0   'False
            Top             =   120
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
         Begin VB.TextBox txtTitle 
            Height          =   315
            Left            =   3600
            TabIndex        =   136
            Top             =   2160
            Width           =   3375
         End
         Begin VB.TextBox txtArtist 
            Height          =   315
            Left            =   3600
            TabIndex        =   135
            Top             =   2880
            Width           =   3375
         End
         Begin VB.TextBox txtAlbum 
            Height          =   315
            Left            =   3600
            TabIndex        =   134
            Top             =   3600
            Width           =   3375
         End
         Begin VB.TextBox txtYear 
            Height          =   315
            Left            =   3600
            TabIndex        =   133
            Top             =   4320
            Width           =   3375
         End
         Begin VB.TextBox txtComment 
            Height          =   315
            Left            =   3600
            TabIndex        =   132
            Top             =   5040
            Width           =   3375
         End
         Begin VB.TextBox txtGenreCode 
            Height          =   315
            Left            =   6360
            TabIndex        =   131
            Top             =   5760
            Width           =   615
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   3600
            TabIndex        =   130
            Text            =   "Genres"
            Top             =   5760
            Width           =   2655
         End
         Begin VB.Timer Timer6 
            Interval        =   60
            Left            =   1560
            Top             =   3720
         End
         Begin VB.DriveListBox Drive1 
            Height          =   315
            Left            =   120
            TabIndex        =   129
            Top             =   3120
            Width           =   3375
         End
         Begin VB.CommandButton cmdOpen 
            Caption         =   "Open MP3"
            Height          =   495
            Left            =   120
            TabIndex        =   128
            Top             =   3720
            Width           =   1215
         End
         Begin VB.CommandButton cmdWriteTag 
            Caption         =   "Save New Info"
            Height          =   495
            Left            =   2280
            TabIndex        =   127
            Top             =   3720
            Width           =   1215
         End
         Begin VB.DirListBox Dir1 
            Height          =   2790
            Left            =   120
            TabIndex        =   126
            Top             =   240
            Width           =   3375
         End
         Begin VB.FileListBox File12 
            Height          =   2625
            Left            =   120
            Pattern         =   "*.mp3"
            TabIndex        =   125
            Top             =   4320
            Width           =   3375
         End
         Begin VB.Frame Frame11 
            BackColor       =   &H80000012&
            BorderStyle     =   0  'None
            Height          =   6735
            Left            =   120
            TabIndex        =   138
            Top             =   240
            Width           =   6855
            Begin VB.Label Label31 
               BackColor       =   &H80000012&
               Caption         =   " Code"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000E&
               Height          =   255
               Left            =   6240
               TabIndex        =   149
               Top             =   5280
               Width           =   615
            End
            Begin VB.Label Label30 
               BackColor       =   &H80000012&
               Caption         =   "Genre: "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000E&
               Height          =   255
               Left            =   3480
               TabIndex        =   148
               Top             =   5280
               Width           =   855
            End
            Begin VB.Label Label29 
               BackColor       =   &H80000012&
               Caption         =   "Comment: "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000E&
               Height          =   255
               Left            =   3480
               TabIndex        =   147
               Top             =   4560
               Width           =   855
            End
            Begin VB.Label Label28 
               BackColor       =   &H80000012&
               Caption         =   "Year: "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000E&
               Height          =   255
               Left            =   3480
               TabIndex        =   146
               Top             =   3840
               Width           =   855
            End
            Begin VB.Label Label27 
               BackColor       =   &H80000012&
               Caption         =   "Album:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000E&
               Height          =   255
               Left            =   3480
               TabIndex        =   145
               Top             =   3120
               Width           =   855
            End
            Begin VB.Label Label26 
               BackColor       =   &H80000012&
               Caption         =   "Artist: "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000E&
               Height          =   255
               Left            =   3480
               TabIndex        =   144
               Top             =   2400
               Width           =   855
            End
            Begin VB.Label Label25 
               BackColor       =   &H80000012&
               Caption         =   "Title: "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000E&
               Height          =   255
               Left            =   3480
               TabIndex        =   143
               Top             =   1680
               Width           =   855
            End
            Begin VB.Label lblElapsedTime 
               BackColor       =   &H80000012&
               Caption         =   "Elapsed Time:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000E&
               Height          =   255
               Left            =   3720
               TabIndex        =   140
               Top             =   6360
               Width           =   3135
            End
            Begin VB.Label lblTotalTime 
               BackColor       =   &H80000012&
               Caption         =   "Total Time:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000E&
               Height          =   255
               Left            =   3720
               TabIndex        =   139
               Top             =   6000
               Width           =   3135
            End
         End
         Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
            Height          =   7455
            Left            =   120
            TabIndex        =   137
            Top             =   240
            Width           =   6855
            AudioStream     =   -1
            AutoSize        =   0   'False
            AutoStart       =   -1  'True
            AnimationAtStart=   0   'False
            AllowScan       =   -1  'True
            AllowChangeDisplaySize=   -1  'True
            AutoRewind      =   0   'False
            Balance         =   0
            BaseURL         =   ""
            BufferingTime   =   5
            CaptioningID    =   ""
            ClickToPlay     =   -1  'True
            CursorType      =   0
            CurrentPosition =   -1
            CurrentMarker   =   0
            DefaultFrame    =   ""
            DisplayBackColor=   0
            DisplayForeColor=   16777215
            DisplayMode     =   0
            DisplaySize     =   4
            Enabled         =   -1  'True
            EnableContextMenu=   -1  'True
            EnablePositionControls=   -1  'True
            EnableFullScreenControls=   0   'False
            EnableTracker   =   -1  'True
            Filename        =   ""
            InvokeURLs      =   -1  'True
            Language        =   -1
            Mute            =   0   'False
            PlayCount       =   1
            PreviewMode     =   0   'False
            Rate            =   1
            SAMILang        =   ""
            SAMIStyle       =   ""
            SAMIFileName    =   ""
            SelectionStart  =   -1
            SelectionEnd    =   -1
            SendOpenStateChangeEvents=   -1  'True
            SendWarningEvents=   -1  'True
            SendErrorEvents =   -1  'True
            SendKeyboardEvents=   0   'False
            SendMouseClickEvents=   0   'False
            SendMouseMoveEvents=   0   'False
            SendPlayStateChangeEvents=   -1  'True
            ShowCaptioning  =   0   'False
            ShowControls    =   -1  'True
            ShowAudioControls=   -1  'True
            ShowDisplay     =   0   'False
            ShowGotoBar     =   0   'False
            ShowPositionControls=   -1  'True
            ShowStatusBar   =   0   'False
            ShowTracker     =   -1  'True
            TransparentAtStart=   0   'False
            VideoBorderWidth=   0
            VideoBorderColor=   0
            VideoBorder3D   =   0   'False
            Volume          =   -600
            WindowlessVideo =   0   'False
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H80000012&
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   8055
         Left            =   -75000
         TabIndex        =   123
         Top             =   360
         Width           =   7335
      End
      Begin VB.Frame frmIPSniff 
         BackColor       =   &H80000012&
         Caption         =   "Port Sniffer"
         ForeColor       =   &H8000000E&
         Height          =   2775
         Left            =   -74880
         TabIndex        =   71
         Top             =   480
         Width           =   7095
         Begin VB.CommandButton Command7 
            BackColor       =   &H80000015&
            DownPicture     =   "SicklidServer.frx":A9D2
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
            Picture         =   "SicklidServer.frx":D6F0
            Style           =   1  'Graphical
            TabIndex        =   94
            Top             =   2160
            UseMaskColor    =   -1  'True
            Width           =   1815
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H80000012&
            Height          =   1935
            Left            =   4800
            TabIndex        =   82
            Top             =   720
            Width           =   2175
            Begin VB.CommandButton Command4 
               Caption         =   "Pause"
               Enabled         =   0   'False
               Height          =   255
               Left            =   1200
               TabIndex        =   85
               Top             =   600
               Width           =   855
            End
            Begin VB.ListBox List1 
               Height          =   840
               Left            =   120
               TabIndex        =   84
               Top             =   960
               Width           =   1935
            End
            Begin VB.TextBox Text99 
               Height          =   285
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   83
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
               TabIndex        =   87
               Top             =   720
               Width           =   975
            End
            Begin VB.Label Label11 
               BackColor       =   &H80000012&
               Caption         =   "Scanning port:"
               ForeColor       =   &H8000000E&
               Height          =   255
               Left            =   120
               TabIndex        =   86
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H80000012&
            ForeColor       =   &H8000000E&
            Height          =   1935
            Left            =   120
            TabIndex        =   72
            Top             =   720
            Width           =   2055
            Begin VB.TextBox Text5 
               Height          =   285
               Left            =   1440
               TabIndex        =   77
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
               TabIndex        =   76
               Top             =   1560
               Width           =   615
            End
            Begin VB.TextBox Text3 
               Height          =   285
               Left            =   1200
               TabIndex        =   75
               Text            =   "32000"
               Top             =   1200
               Width           =   735
            End
            Begin VB.TextBox Text2 
               Height          =   285
               Left            =   120
               TabIndex        =   74
               Text            =   "1"
               Top             =   1200
               Width           =   735
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Left            =   120
               TabIndex        =   73
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
               TabIndex        =   81
               Top             =   1605
               Width           =   735
            End
            Begin VB.Label Label13 
               Alignment       =   2  'Center
               Caption         =   "To"
               Height          =   255
               Left            =   840
               TabIndex        =   80
               Top             =   1245
               Width           =   375
            End
            Begin VB.Label Label14 
               BackColor       =   &H80000012&
               Caption         =   "Range of ports to scan:"
               ForeColor       =   &H8000000E&
               Height          =   255
               Left            =   120
               TabIndex        =   79
               Top             =   960
               Width           =   1815
            End
            Begin VB.Label Label15 
               BackColor       =   &H80000012&
               Caption         =   "IP Address to scan:"
               ForeColor       =   &H8000000E&
               Height          =   255
               Left            =   120
               TabIndex        =   78
               Top             =   240
               Width           =   1815
            End
         End
         Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash3 
            Height          =   1215
            Left            =   2640
            TabIndex        =   95
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
            Enabled         =   0   'False
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
            TabIndex        =   90
            Top             =   240
            Width           =   6855
         End
      End
      Begin VB.Frame Chatroom 
         BackColor       =   &H00000000&
         Caption         =   "Chatroom"
         ForeColor       =   &H8000000E&
         Height          =   7815
         Left            =   -74880
         TabIndex        =   58
         Top             =   480
         Width           =   7095
         Begin VB.CommandButton Command9 
            Caption         =   "Winsock Accept"
            Height          =   495
            Left            =   3480
            TabIndex        =   101
            TabStop         =   0   'False
            Top             =   6000
            Width           =   1575
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Close Client"
            Height          =   495
            Left            =   1800
            TabIndex        =   100
            TabStop         =   0   'False
            Top             =   6000
            Width           =   1575
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Close Accept"
            Height          =   495
            Left            =   5160
            TabIndex        =   99
            TabStop         =   0   'False
            Top             =   6000
            Width           =   1575
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Winsock Client"
            Height          =   495
            Left            =   120
            TabIndex        =   98
            TabStop         =   0   'False
            Top             =   6000
            Width           =   1575
         End
         Begin VB.CommandButton cmdClrScreen 
            Caption         =   "Clear Screen"
            Height          =   375
            Left            =   5400
            Style           =   1  'Graphical
            TabIndex        =   91
            Top             =   4440
            Width           =   1335
         End
         Begin VB.ListBox lstPeople 
            BackColor       =   &H00FFFFFF&
            Height          =   2400
            Left            =   5160
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox txtSend 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   120
            TabIndex        =   64
            Top             =   5040
            Width           =   4095
         End
         Begin VB.CommandButton cmdSend 
            Caption         =   "Send"
            Default         =   -1  'True
            Enabled         =   0   'False
            Height          =   285
            Left            =   4320
            Style           =   1  'Graphical
            TabIndex        =   63
            Top             =   5040
            Width           =   735
         End
         Begin VB.CommandButton cmdConnect 
            Caption         =   "Connect"
            Height          =   375
            Left            =   5400
            Style           =   1  'Graphical
            TabIndex        =   62
            Top             =   3480
            Width           =   1335
         End
         Begin VB.CommandButton cmdHost 
            Caption         =   "Host a Room"
            Height          =   375
            Left            =   5400
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   3960
            Width           =   1335
         End
         Begin VB.ListBox lstbye 
            Height          =   450
            Left            =   0
            TabIndex        =   60
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton cmdSearchTxt 
            Caption         =   "Search Text"
            Height          =   375
            Left            =   5400
            Style           =   1  'Graphical
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   4920
            Width           =   1335
         End
         Begin MSComDlg.CommonDialog cdialog 
            Left            =   240
            Top             =   4200
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin MSWinsockLib.Winsock sckAccept 
            Left            =   0
            Top             =   1080
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   393216
         End
         Begin MSWinsockLib.Winsock sckServer 
            Index           =   0
            Left            =   0
            Top             =   720
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   393216
         End
         Begin MSWinsockLib.Winsock sckClient 
            Left            =   0
            Top             =   120
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   393216
         End
         Begin RichTextLib.RichTextBox txtIncoming 
            Height          =   4215
            Left            =   120
            TabIndex        =   102
            TabStop         =   0   'False
            Top             =   600
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   7435
            _Version        =   393217
            Enabled         =   -1  'True
            ScrollBars      =   2
            TextRTF         =   $"SicklidServer.frx":1040E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.TextBox txtIncoming2 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000017&
            Height          =   1455
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   92
            Top             =   600
            Visible         =   0   'False
            Width           =   4335
         End
         Begin VB.Label lblStatus2 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Status"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   97
            Top             =   5520
            Width           =   6855
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            BackColor       =   &H80000017&
            Caption         =   "People Chatting:"
            ForeColor       =   &H80000014&
            Height          =   255
            Left            =   5160
            TabIndex        =   96
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label lblChatroom 
            Alignment       =   2  'Center
            BackColor       =   &H80000017&
            Caption         =   "Not Connected or Hosting a Chatroom"
            ForeColor       =   &H80000014&
            Height          =   255
            Left            =   120
            TabIndex        =   93
            Top             =   360
            Width           =   4335
         End
         Begin VB.Label Label55 
            BackStyle       =   0  'Transparent
            Caption         =   "People Chatting:"
            Height          =   255
            Left            =   7800
            TabIndex        =   66
            Top             =   120
            Width           =   1815
         End
      End
      Begin VB.CommandButton Btn_Browse 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Select File"
         Height          =   375
         Left            =   1800
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   56
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000007&
         Caption         =   "Chat Window"
         ForeColor       =   &H8000000E&
         Height          =   7815
         Left            =   -74880
         TabIndex        =   88
         Top             =   480
         Width           =   7095
         Begin VB.CommandButton Command12 
            Caption         =   "wdata State"
            Height          =   495
            Left            =   5520
            TabIndex        =   104
            Top             =   600
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H80000015&
            DownPicture     =   "SicklidServer.frx":10489
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
            Picture         =   "SicklidServer.frx":131A7
            Style           =   1  'Graphical
            TabIndex        =   52
            Top             =   1440
            UseMaskColor    =   -1  'True
            Width           =   1815
         End
         Begin VB.Frame frmConnection 
            BackColor       =   &H00000000&
            Caption         =   "Connection"
            ForeColor       =   &H00FFFFFF&
            Height          =   1575
            Left            =   120
            TabIndex        =   89
            Top             =   2040
            Width           =   6855
            Begin VB.TextBox txtNick 
               Alignment       =   2  'Center
               Height          =   288
               Left            =   1200
               TabIndex        =   40
               Text            =   "Guest"
               Top             =   720
               Width           =   1455
            End
            Begin VB.TextBox txtPort 
               Alignment       =   2  'Center
               Height          =   288
               Left            =   3600
               TabIndex        =   45
               Text            =   "69"
               Top             =   240
               Width           =   492
            End
            Begin VB.CommandButton cmdDisconnect 
               Caption         =   "&Disconnect"
               Height          =   372
               Left            =   3720
               TabIndex        =   43
               Top             =   480
               Visible         =   0   'False
               Width           =   1212
            End
            Begin VB.CommandButton cmdListen 
               Caption         =   "&Listen"
               Height          =   372
               Left            =   5520
               TabIndex        =   42
               Top             =   720
               Width           =   1212
            End
            Begin VB.CommandButton cmdConnect2 
               Caption         =   "&Connect"
               Enabled         =   0   'False
               Height          =   372
               Left            =   3720
               TabIndex        =   41
               Top             =   720
               Width           =   1212
            End
            Begin VB.TextBox txtLocal 
               Alignment       =   2  'Center
               Height          =   288
               Left            =   5280
               Locked          =   -1  'True
               TabIndex        =   44
               TabStop         =   0   'False
               Text            =   "127.0.0.1"
               Top             =   240
               Width           =   1452
            End
            Begin VB.TextBox txtRemote 
               Alignment       =   2  'Center
               Height          =   288
               Left            =   1200
               TabIndex        =   39
               Text            =   "0.0.0.0"
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label llnNick 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nickname:"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Left            =   360
               TabIndex        =   51
               Top             =   720
               Width           =   735
            End
            Begin VB.Label lblStatus 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Status"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   50
               Top             =   1200
               Width           =   6615
            End
            Begin VB.Label lblPort 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Port:"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Left            =   3240
               TabIndex        =   49
               Top             =   240
               Width           =   360
            End
            Begin VB.Label lblLocal 
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               Caption         =   "Local IP:"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Left            =   4560
               TabIndex        =   48
               Top             =   240
               Width           =   615
            End
            Begin VB.Label lblRemote 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Remote Host:"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Left            =   120
               TabIndex        =   47
               Top             =   360
               Width           =   990
            End
         End
         Begin VB.CommandButton cmdSend2 
            Caption         =   "&Send"
            Enabled         =   0   'False
            Height          =   372
            Left            =   5880
            TabIndex        =   46
            Top             =   7200
            Width           =   1092
         End
         Begin VB.TextBox txtText 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   38
            Top             =   7200
            Width           =   5655
         End
         Begin MSWinsockLib.Winsock wData 
            Left            =   4200
            Top             =   5040
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   393216
         End
         Begin RichTextLib.RichTextBox txtData 
            Height          =   3375
            Left            =   120
            TabIndex        =   53
            Top             =   3720
            Width           =   6855
            _ExtentX        =   12091
            _ExtentY        =   5953
            _Version        =   393217
            Enabled         =   -1  'True
            ScrollBars      =   2
            TextRTF         =   $"SicklidServer.frx":15EC5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash2 
            Height          =   1215
            Left            =   2640
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   120
            Width           =   1815
            _cx             =   22744193
            _cy             =   22743135
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
         Begin VB.Image Image1 
            Height          =   2535
            Left            =   120
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00000000&
         Caption         =   "Client Side"
         ForeColor       =   &H00FFFFFF&
         Height          =   1815
         Left            =   120
         TabIndex        =   29
         Top             =   6480
         Width           =   7095
         Begin VB.TextBox Txt_File2 
            Height          =   285
            Left            =   120
            TabIndex        =   30
            Text            =   "C:\"
            Top             =   480
            Width           =   6855
         End
         Begin VB.Timer Timer2 
            Interval        =   1000
            Left            =   2760
            Top             =   120
         End
         Begin MSComctlLib.ProgressBar FileBar2 
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   1440
            Width           =   6855
            _ExtentX        =   12091
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   0
            Scrolling       =   1
         End
         Begin MSWinsockLib.Winsock Winsock_Receive 
            Left            =   3360
            Top             =   120
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   393216
         End
         Begin VB.Label Lbl_FileSize2 
            BackStyle       =   0  'Transparent
            Caption         =   "Filesize: 0 kb"
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
            Left            =   120
            TabIndex        =   37
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label Lbl_Complete2 
            BackStyle       =   0  'Transparent
            Caption         =   "Complete: 0%"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Save File to:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   4335
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Filename:"
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
            Left            =   2280
            TabIndex        =   34
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Lbl_Averages2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Average: 0 / KBps"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   4080
            TabIndex        =   33
            Top             =   1200
            Width           =   2055
         End
         Begin VB.Label Lbl_FileName2 
            BackStyle       =   0  'Transparent
            Caption         =   "-"
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
            Left            =   3240
            TabIndex        =   32
            Top             =   840
            Width           =   3135
         End
      End
      Begin VB.Frame FraServer 
         BackColor       =   &H00000000&
         Caption         =   "Server Side"
         ForeColor       =   &H00FFFFFF&
         Height          =   1815
         Left            =   120
         TabIndex        =   20
         Top             =   4560
         Width           =   7095
         Begin VB.TextBox Txt_File 
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   22
            Text            =   "C:\"
            Top             =   480
            Width           =   6855
         End
         Begin VB.Timer Timer1 
            Interval        =   1000
            Left            =   2880
            Top             =   120
         End
         Begin MSComctlLib.ProgressBar FileBar 
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   1440
            Width           =   6855
            _ExtentX        =   12091
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   0
            Scrolling       =   1
         End
         Begin MSComDlg.CommonDialog Dlg_Browser 
            Left            =   3840
            Top             =   120
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Selected File:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   4335
         End
         Begin VB.Label Lbl_FileSize 
            BackStyle       =   0  'Transparent
            Caption         =   "Filesize: 0 kb"
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
            Left            =   120
            TabIndex        =   27
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Lbl_FileName 
            BackStyle       =   0  'Transparent
            Caption         =   "-"
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
            Left            =   3240
            TabIndex        =   26
            Top             =   840
            Width           =   3135
         End
         Begin VB.Label Lbl_Complete 
            BackStyle       =   0  'Transparent
            Caption         =   "Complete: 0%"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   1200
            Width           =   2175
         End
         Begin VB.Label Lbl_Averages 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Average: 0 / KBps"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   4200
            TabIndex        =   24
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Filename:"
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
            Left            =   2280
            TabIndex        =   23
            Top             =   840
            Width           =   855
         End
      End
      Begin VB.Frame Fra_Advanced 
         BackColor       =   &H00000000&
         Caption         =   "Advanced Settings"
         ForeColor       =   &H00FFFFFF&
         Height          =   3495
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   7095
         Begin VB.CommandButton Command6 
            BackColor       =   &H80000015&
            DownPicture     =   "SicklidServer.frx":15F40
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
            Picture         =   "SicklidServer.frx":18C5E
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   2160
            UseMaskColor    =   -1  'True
            Width           =   1815
         End
         Begin VB.TextBox Txt_Port 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            TabIndex        =   9
            Text            =   "420"
            Top             =   960
            Width           =   1575
         End
         Begin VB.TextBox Txt_RemoteIP 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            TabIndex        =   8
            Text            =   "0.0.0.0"
            Top             =   2040
            Width           =   1575
         End
         Begin VB.TextBox Txt_CurrentIP 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   5520
            TabIndex        =   7
            Top             =   2160
            Width           =   1455
         End
         Begin VB.TextBox Txt_Port2 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   5520
            TabIndex        =   6
            Text            =   "420"
            Top             =   960
            Width           =   1455
         End
         Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
            Height          =   1215
            Left            =   2640
            TabIndex        =   54
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
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "IP to Connect To:"
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
            Left            =   120
            TabIndex        =   103
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Line Line5 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   3
            X1              =   3480
            X2              =   3480
            Y1              =   240
            Y2              =   840
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   3
            X1              =   3600
            X2              =   3600
            Y1              =   240
            Y2              =   840
         End
         Begin VB.Label Lbl_Port 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Port to Connect To:"
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
            Left            =   0
            TabIndex        =   19
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Lbl_Info 
            BackStyle       =   0  'Transparent
            Caption         =   "(the port has to be the same in the client form)"
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   120
            TabIndex        =   18
            Top             =   1320
            Width           =   2175
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "(the IP address you connect to: local 127.0.0.1)"
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   120
            TabIndex        =   17
            Top             =   2400
            Width           =   2175
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   3
            X1              =   3480
            X2              =   3480
            Y1              =   2760
            Y2              =   3360
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   " SEND"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label lbl_ExternalIP 
            BackStyle       =   0  'Transparent
            Caption         =   "Your External IP:"
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
            Left            =   5520
            TabIndex        =   15
            Top             =   1920
            Width           =   1455
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Port to listen to:"
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
            Left            =   5520
            TabIndex        =   14
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "(0 = free port, the port has to be the same in the server form)"
            ForeColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   4800
            TabIndex        =   13
            Top             =   1320
            Width           =   2175
         End
         Begin VB.Label Lbl_Status 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "winsock State"
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   120
            TabIndex        =   12
            Top             =   2880
            Width           =   2055
         End
         Begin VB.Label Lbl_Status2 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "winsock State"
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   5280
            TabIndex        =   11
            Top             =   2520
            Width           =   1695
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   3
            X1              =   3600
            X2              =   3600
            Y1              =   2760
            Y2              =   3360
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "LISTEN"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   5520
            TabIndex        =   10
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.CommandButton Btn_Send 
         Caption         =   "Send"
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   4080
         Width           =   1215
      End
      Begin VB.CommandButton Btn_Listen 
         Caption         =   "Listen for file"
         Height          =   375
         Left            =   4440
         TabIndex        =   3
         Top             =   4080
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Stop Transfer"
         Height          =   375
         Left            =   5760
         TabIndex        =   2
         Top             =   4080
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Close Program"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   1
         Top             =   4080
         Width           =   1215
      End
      Begin MSWinsockLib.Winsock Winsock_Send 
         Left            =   3480
         Top             =   5400
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H80000012&
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   8895
         Left            =   -75000
         TabIndex        =   67
         Top             =   360
         Width           =   10815
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H80000012&
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   8775
         Left            =   0
         TabIndex        =   68
         Top             =   360
         Width           =   10815
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H80000012&
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   8775
         Left            =   -75120
         TabIndex        =   70
         Top             =   360
         Width           =   10335
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H80000012&
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   8775
         Left            =   -75000
         TabIndex        =   69
         Top             =   360
         Width           =   10815
         Begin VB.Frame IPFrame 
            BackColor       =   &H80000012&
            Caption         =   "IP Pinger"
            ForeColor       =   &H8000000E&
            Height          =   4935
            Left            =   120
            TabIndex        =   105
            Top             =   3000
            Width           =   7095
            Begin VB.TextBox Text4 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   375
               Index           =   2
               Left            =   4800
               TabIndex        =   122
               Top             =   1080
               Width           =   2175
            End
            Begin VB.CommandButton Command14 
               Caption         =   "Ping IP"
               Height          =   375
               Left            =   120
               TabIndex        =   115
               Top             =   360
               Width           =   855
            End
            Begin VB.TextBox Text7 
               Alignment       =   2  'Center
               Height          =   375
               Left            =   1080
               TabIndex        =   113
               Text            =   "IP Address"
               Top             =   360
               Width           =   1935
            End
            Begin VB.TextBox Text6 
               Height          =   375
               Left            =   3120
               TabIndex        =   114
               Text            =   "Message"
               Top             =   360
               Width           =   2775
            End
            Begin VB.TextBox Text4 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   112
               Top             =   1080
               Width           =   2175
            End
            Begin VB.TextBox Text4 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   375
               Index           =   1
               Left            =   2400
               TabIndex        =   111
               Top             =   1080
               Width           =   2295
            End
            Begin VB.TextBox Text4 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   375
               Index           =   3
               Left            =   120
               TabIndex        =   110
               Top             =   1800
               Width           =   2175
            End
            Begin VB.TextBox Text4 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   375
               Index           =   4
               Left            =   2400
               TabIndex        =   109
               Top             =   1800
               Width           =   2295
            End
            Begin VB.TextBox Text4 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   375
               Index           =   5
               Left            =   4800
               TabIndex        =   108
               Top             =   1800
               Width           =   2175
            End
            Begin VB.CommandButton Command13 
               Caption         =   "Get host IP"
               Height          =   375
               Left            =   6000
               TabIndex        =   107
               Top             =   360
               Width           =   975
            End
            Begin MSComctlLib.ListView ListView1 
               Height          =   2535
               Left            =   120
               TabIndex        =   106
               Top             =   2280
               Width           =   6855
               _ExtentX        =   12091
               _ExtentY        =   4471
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   0
            End
            Begin VB.Label Label24 
               Alignment       =   2  'Center
               BackColor       =   &H80000012&
               Caption         =   "Data Pointer"
               ForeColor       =   &H8000000E&
               Height          =   255
               Left            =   4800
               TabIndex        =   121
               Top             =   1560
               Width           =   2175
            End
            Begin VB.Label Label23 
               Alignment       =   2  'Center
               BackColor       =   &H80000012&
               Caption         =   "Data returned"
               ForeColor       =   &H8000000E&
               Height          =   255
               Left            =   2400
               TabIndex        =   120
               Top             =   1560
               Width           =   2295
            End
            Begin VB.Label Label22 
               Alignment       =   2  'Center
               BackColor       =   &H80000012&
               Caption         =   "Data packet size"
               ForeColor       =   &H8000000E&
               Height          =   255
               Left            =   120
               TabIndex        =   119
               Top             =   1560
               Width           =   2175
            End
            Begin VB.Label Label21 
               Alignment       =   2  'Center
               BackColor       =   &H80000012&
               Caption         =   "Round trip time:"
               ForeColor       =   &H8000000E&
               Height          =   255
               Left            =   4800
               TabIndex        =   118
               Top             =   840
               Width           =   2175
            End
            Begin VB.Label Label20 
               Alignment       =   2  'Center
               BackColor       =   &H80000012&
               Caption         =   "IP Address Being Pinged:"
               ForeColor       =   &H8000000E&
               Height          =   255
               Left            =   2400
               TabIndex        =   117
               Top             =   840
               Width           =   2295
            End
            Begin VB.Label Label19 
               Alignment       =   2  'Center
               BackColor       =   &H80000012&
               Caption         =   "Return Status:"
               ForeColor       =   &H8000000E&
               Height          =   255
               Left            =   120
               TabIndex        =   116
               Top             =   840
               Width           =   2175
            End
         End
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         X1              =   2280
         X2              =   2280
         Y1              =   720
         Y2              =   3960
      End
   End
   Begin VB.Label Label32 
      BackColor       =   &H80000012&
      Caption         =   "Genre: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   150
      Top             =   0
      Width           =   855
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLogtoFile 
         Caption         =   "&Log to File (Port Sniffer ONLY)"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuTabs 
      Caption         =   "&Tabs"
      Begin VB.Menu mnuFileSharing 
         Caption         =   "File &Sharing"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuPrivateChat 
         Caption         =   "&Private Chat"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuMisc 
         Caption         =   "&Fun Stuff"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuChatroom 
         Caption         =   "&Chatroom"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuMP3Player 
         Caption         =   "&MP3 Player"
         Shortcut        =   ^M
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "SicklidServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'# **********************
'# ****MP3 PLAYER VAR****
'# **********************

Dim GenresTypes
Dim Min As Integer
Dim Sec As Integer

Dim FileName As String
Dim FileOpen As Boolean
Dim CurrentTag As TagInfo

Private Type TagInfo
    Tag As String * 3
    Songname As String * 30
    artist As String * 30
    album As String * 30
    year As String * 4
    comment As String * 30
    genre As String * 1
End Type

'# **********************
'# ****IP PINGER VAR*****
'# **********************
Private Const ERROR_SUCCESS As Long = 0

Private Type MIB_IPADDRROW
  dwAddr As Long        'IP address
  dwIndex As Long       'index of interface associated with this IP
  dwMask As Long        'subnet mask for the IP address
  dwBCastAddr As Long   'broadcast address (typically the IP
                        'with host portion set to either all
                        'zeros or all ones)
  dwReasmSize As Long   'reassembly size for received datagrams
  unused1 As Integer    'not currently used (but shown anyway)
  unused2 As Integer    'not currently used (but shown anyway)
End Type

Private Declare Function GetIpAddrTable Lib "IPHLPAPI.DLL" _
  (ByRef ipAddrTable As Byte, _
   ByRef dwSize As Long, _
   ByVal bOrder As Long) As Long
   
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
  (dst As Any, src As Any, ByVal bcount As Long)
  
Private Declare Function inet_ntoa Lib "WSOCK32.DLL" (ByVal addr As Long) As Long

Private Declare Function lstrcpyA Lib "kernel32" _
  (ByVal RetVal As String, ByVal Ptr As Long) As Long
                        
Private Declare Function lstrlenA Lib "kernel32" _
  (ByVal Ptr As Any) As Long
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
'# **********************
'# *****CHATROOM VAR*****
'# **********************
Public okgo As Boolean '6
Public lastdata$
Public Peeps As Integer
'# **********************

'# **********************
'# *****FILESHAREVAR*****
'# **********************

Dim DoneBytes As Long
Dim NextPart As Boolean
Dim DoneRec As Long                     '# This is for calculating the KB per second receiving
Dim DownloadingFile As Integer          '# This is the freefile for open files
Public UserName As String               '# This is the String for the UserName entered
Public RecText As String                '# This is the String for the location of the received file
Public RecSize As String                '# This is the String for the size of the received file
Public RecName As String                '# This is the String for the Name of the received file
'# **********************


'*****************************************************'
'*  Created on June 02,2001 over a period of 5 days  *'
'*  By Skubidu (skubidu42@hotmail.com) and           *'
'*  Sicklid (sicklid_vision@yahoo.com)               *'
'*****************************************************'

'# Visit us at www.sicklidsoft.com
'# Contact me: sicklid_vision@yahoo.com.
'# There should be no major errors, but the possibility exists for minor errors.
'# You can also contact the senior programmer, Skubidu at: skubidu@yahoo.com.
'# I hope you enjoy the program and if you find major errors, please contact us!


Private Sub Btn_Browse_Click()
On Error GoTo Quit
    Dlg_Browser.ShowOpen                            '# This opens the browser window to choose the file to send.
    Txt_File.Text = Dlg_Browser.FileName            '# This sends the name of the FileName chosen to the TextBox.
    Lbl_FileSize.Caption = "Filesize: " & FileLen(Dlg_Browser.FileName) & " kb" '# This sets the size of the file to FileSize.Caption
    Lbl_FileName.Caption = Dlg_Browser.FileTitle    '# This sets the FileName.Caption to the Title of the File being sent.
    RecText = Txt_File.Text                         '# This sets the Textbox value equal to the variable RecText to be sent to the Receiving party.
    RecSize = Lbl_FileSize.Caption                  '# This sets the lbl_FileSize.Caption equal to the variable RecSize to be sent to the Receiving party.
    RecName = Lbl_FileName.Caption                  '# This sets the lbl_FileName.Caption equal to the variable RecName to be sent to the Receiving party.

Quit:

End Sub

Private Sub Btn_Listen_Click()
On Error GoTo ErrorHandler:
        
        
'The following are necessary to reduce the chance of errors
        
        If Winsock_Receive.State <> sckClosed Then          '# If the Winsock is not closed, it closes the winsock.
            Winsock_Receive.Close
        End If
        Winsock_Receive.Protocol = sckTCPProtocol           '# This sets the protocol of the Winsock to TCP.
        
'# Init the Winsock
        If Txt_Port2.Text <> 0 Then                         '# If the port chosen is not zero, then the localport is equal to the selected port.
                Winsock_Receive.LocalPort = Txt_Port2.Text
        Else
                Winsock_Receive.LocalPort = 0               '# In the event Zero is selected, a free port is chosen.
        End If
        Winsock_Receive.Listen                              '# This sets the current local receiving port state to LISTEN or 2.
        
        
'# The following describes what each winsock state means and returns that value to the lbl_status2.Caption.
        If Winsock_Receive.State = 0 Then
                Lbl_Status2.Caption = "Connection closed on port: " & Winsock_Receive.LocalPort
        ElseIf Winsock_Receive.State = 1 Then
                Lbl_Status2.Caption = "Connection in use on port: " & Winsock_Receive.LocalPort
        ElseIf Winsock_Receive.State = 2 Then
                Lbl_Status2.Caption = "Winsock listening on port: " & Winsock_Receive.LocalPort
        ElseIf Winsock_Receive.State = 3 Then
                Lbl_Status2.Caption = "Connection pending on port: " & Winsock_Receive.LocalPort
        ElseIf Winsock_Receive.State = 4 Then
                Lbl_Status2.Caption = "Connection pending on port: " & Winsock_Receive.LocalPort
        ElseIf Winsock_Receive.State = 5 Then
                Lbl_Status2.Caption = "Host resolved on port: " & Winsock_Receive.LocalPort
        ElseIf Winsock_Receive.State = 6 Then
                Lbl_Status2.Caption = "Connecting on port: " & Winsock_Receive.LocalPort
        ElseIf Winsock_Receive.State = 7 Then
                Lbl_Status2.Caption = "Connected on port: " & Winsock_Receive.LocalPort
        ElseIf Winsock_Receive.State = 8 Then
                Lbl_Status2.Caption = "Peer closed connection on port: " & Winsock_Receive.LocalPort
        ElseIf Winsock_Receive.State = 9 Then
                Lbl_Status2.Caption = "Error on port: " & Winsock_Receive.LocalPort
        End If

Exit Sub
        
ErrorHandler:
    MsgBox err.Description, vbCritical
End Sub

Private Sub Btn_Send_Click()
On Error GoTo ErrorHandler:
    
Dim StartTime As Long
        
'The following are necessary to reduce the chance of errors
        If Winsock_Send.State <> sckClosed Then             '# If the Winsock is not closed, it closes the winsock.
            Winsock_Send.Close
        ElseIf Winsock_Send.State <> 0 Then
            Winsock_Send.Close
        End If
        
        Winsock_Send.Protocol = sckTCPProtocol              '# This sets the protocol of the Winsock to TCP.
        Winsock_Send.LocalPort = 0                          '# This sets the localport to zero, or a free port, in order to initialize the port.
        
'# Init the Winsock
        If Txt_Port.Text <> 0 Then                          '# If the port chosen is not zero, then the localport is equal to the selected port.
            Winsock_Send.RemotePort = Txt_Port.Text         '# This sets the winsock_send remoteport equal to the Txt_Port.Text value.  This should be the same as the client Port.
            Winsock_Send.RemoteHost = Txt_RemoteIP.Text     '# This sets the winsock_send remotehost equal to the RemoteIP.Text value.  This should be the same as the client IP.
        Else
            MsgBox "Select a Port first!"
            Exit Sub
        End If
        
'# The following describes what each winsock state means and returns that value to the lbl_status2.Caption.
        If Winsock_Receive.State = 0 Then
                Lbl_Status2.Caption = "Connection closed on port: " & Winsock_Receive.LocalPort
        ElseIf Winsock_Receive.State = 1 Then
                Lbl_Status2.Caption = "Connection in use on port: " & Winsock_Receive.LocalPort
        ElseIf Winsock_Receive.State = 2 Then
                Lbl_Status2.Caption = "Winsock listening on port: " & Winsock_Receive.LocalPort
        ElseIf Winsock_Receive.State = 3 Then
                Lbl_Status2.Caption = "Connection pending on port: " & Winsock_Receive.LocalPort
        ElseIf Winsock_Receive.State = 4 Then
                Lbl_Status2.Caption = "Connection pending on port: " & Winsock_Receive.LocalPort
        ElseIf Winsock_Receive.State = 5 Then
                Lbl_Status2.Caption = "Host resolved on port: " & Winsock_Receive.LocalPort
        ElseIf Winsock_Receive.State = 6 Then
                Lbl_Status2.Caption = "Connecting on port: " & Winsock_Receive.LocalPort
        ElseIf Winsock_Receive.State = 7 Then
                Lbl_Status2.Caption = "Connected on port: " & Winsock_Receive.LocalPort
        ElseIf Winsock_Receive.State = 8 Then
                Lbl_Status2.Caption = "Peer closed connection on port: " & Winsock_Receive.LocalPort
        ElseIf Winsock_Receive.State = 9 Then
                Lbl_Status2.Caption = "Error on port: " & Winsock_Receive.LocalPort
        End If

'# The following describes what each winsock state means and returns that value to the lbl_status.Caption.
        
        If Winsock_Send.State = 0 Then
                Lbl_Status.Caption = "Connection closed on port: " & Winsock_Send.LocalPort
        ElseIf Winsock_Send.State = 1 Then
                Lbl_Status.Caption = "Connection in use on port: " & Winsock_Send.LocalPort
        ElseIf Winsock_Send.State = 2 Then
                Lbl_Status.Caption = "Winsock listening on port: " & Winsock_Send.LocalPort
        ElseIf Winsock_Send.State = 3 Then
                Lbl_Status.Caption = "Connection pending on port: " & Winsock_Send.LocalPort
        ElseIf Winsock_Send.State = 4 Then
                Lbl_Status.Caption = "Connection pending on port: " & Winsock_Send.LocalPort
        ElseIf Winsock_Send.State = 5 Then
                Lbl_Status.Caption = "Host resolved on port: " & Winsock_Send.LocalPort
        ElseIf Winsock_Send.State = 6 Then
                Lbl_Status.Caption = "Connecting on port: " & Winsock_Send.LocalPort
        ElseIf Winsock_Send.State = 7 Then
                Lbl_Status.Caption = "Connected on port: " & Winsock_Send.LocalPort
        ElseIf Winsock_Send.State = 8 Then
                Lbl_Status.Caption = "Peer closed connection on port: " & Winsock_Send.LocalPort
        ElseIf Winsock_Send.State = 9 Then
                Lbl_Status.Caption = "Error on port: " & Winsock_Send.LocalPort
        End If
        
        Winsock_Send.Connect                                '# This connects to the IP and Port selected
        
        StartTime = Timer
          
        Do While Winsock_Send.State <> 7 And Timer - StartTime < 30
            DoEvents                                        '# This makes the program wait until the connections establishes
        Loop                                                '# This establishes a timeout-check to ensure the program ends if hung.
        
        If Timer - StartTime > 30 Then GoTo Timeout         '# If the Timer exceeds 30 Seconds hangtime, then the program Times Out the connection.
       
       
'-----------------------------------------------------
'# This is the send routine.
'# Here the file is opened in binary mode and read in packages and sent to the connected port
'# The package size can be changed. The default packetsize is 4096 bytes.  It can also be 512, 1024, 2048, etc.
'# The smaller the packets, the longer the transfer.
'-----------------------------------------------------
        
            Dim OpenedFileNbr, FileLength, Back
            Dim Temp As String
            Dim packetsize As Long
            Dim lastdata As Boolean

            
            FileLength = FileLen(Txt_File.Text)
            FileBar.Max = FileLength
            FileBar.Value = 0
            
            Winsock_Send.SendData ("FILEINFO|" & FileLength & "|" & Lbl_FileName.Caption & "|")  '# You can add more like filename , description, etc...see next line
                        
'# The following line of code sends the File's Size, Name and Location to the Receiver.
            
            Winsock_Send.SendData (RecText & RecSize & RecName)
            StartTime = Timer
            
                Do While NextPart = False And Timer - StartTime < 30        '# If the next packet is not sent within 30 seconds, the transfer will timeout.
                    DoEvents
                Loop
                
            If Timer - StartTime > 30 Then GoTo Timeout         '# If the Timer exceeds 30 Seconds hangtime, then the program Times Out the connection.
       
            packetsize = 4096                                   '#  This sets the size of the packets to send.

'On an Error, GoTo ErrorHandler
                    
                    lastdata = False                            '#  This is used to make the received file the same size as the original file.

                    NextPart = True                             '#  The NextPart variable is a global boolean variable which determines whether the packet was sent or not.

                    OpenedFileNbr = FreeFile                    '#  This finds a free filenumber to open the original file.
                    
                    Open Txt_File.Text For Binary Access Read As OpenedFileNbr
                        
                    Temp = ""
                        
                        Do Until EOF(OpenedFileNbr)                     '# This will loop the following code until the End Of File.
                        
' The following code adjusts the packetsize at the end so too much data isn't sent.
                            If FileLength - Loc(OpenedFileNbr) <= packetsize Then
                                packetsize = FileLength - Loc(OpenedFileNbr) + 1
                                lastdata = True
                            End If
                            
                            Temp = Space$(packetsize)                   '# This makes the empty temp variable equal to the size of the packet.
                            Get OpenedFileNbr, , Temp                   '# This loads the data into the empty string, Temp.
                            
                            If Winsock_Send.State <> 7 Then Exit Sub    '# This checks whether the connection remains or if the winsock is closed.
                            On Error Resume Next
                            
                            StartTime = Timer
                                Do While NextPart = False And Timer - StartTime < 30       '# If the next packet is not sent within 30 seconds, the transfer will timeout.
                                    DoEvents
                                Loop
                            
                            If Timer - StartTime > 30 Then GoTo Timeout '# If the Timer exceeds 30 Seconds hangtime, then the program Times Out the connection.
                            
                            If Winsock_Send.State = 7 Then              '# This checks whether the connection remains or if the winsock is closed again.
                            
                            If lastdata = True Then
                                Temp = Mid(Temp, 1, Len(Temp) - 1)      '# One additional byte was sent (as was shown above).  It should not be sent.

                            End If
                                FileBar.Value = FileBar.Value + Len(Temp)
                                Lbl_Complete.Caption = "Complete: " & Int(100 / FileLength * FileBar.Value) & " %"
                                DoneBytes = DoneBytes + Len(Temp)
                            
                            Winsock_Send.SendData Temp                  '# This sends the data package
                                NextPart = False                        '# This sets the senddata check boolean variable.
                            Else
                                Exit Sub
                            End If
                    Loop

                            Close #OpenedFileNbr                        '# Once the last package was sent, the files is closed.
                            FileBar.Value = 0
                            
                            Do While NextPart = False                   '# While there is data to be sent, this code will make sure the
                                DoEvents                                '  winsock connection is not closed prematurely or else all data will be lost.
                            Loop
                                Winsock_Send.Close                      '# This action closes the winsock once all data has been sent.
                                NextPart = False
                            Exit Sub

Timeout:
            MsgBox "Timeout"                                            '# This is a short messagebox notifying the user that a timeout has occurred.
                            
Exit Sub
        
ErrorHandler:
    MsgBox err.Description, vbCritical

End Sub

Private Sub cmdClrScreen_Click()
    txtIncoming.Text = ""
End Sub



Private Sub cmdSearchTxt_Click()
   Dim lngPos As Long
  
   m_strSearch = InputBox("Right now this will only find the first instance of a word or phrase." & vbCrLf & "Enter the text to find.", "Find Text")
   lngPos = InStr(1, txtIncoming.Text, m_strSearch, vbTextCompare)
   If lngPos > 0 Then
      txtIncoming.SelStart = lngPos - 1
      txtIncoming.SelLength = Len(m_strSearch)
      txtIncoming.SetFocus
   Else
      MsgBox "Search text was not found.", vbExclamation
      m_strSearch = ""
      txtSend.SetFocus
   End If

End Sub

Private Sub Command1_Click()
    If Command1.Caption = "&Stop Transfer" Then
        MsgBox "Terminating File Transfer", vbOKOnly, "KILLING FILE TRANSFER"   '# If there is a transfer in progress, the button will kill the transfer, closing the port.
    Else
        MsgBox "Closing the Open Port", vbOKOnly, "KILLING THE PORT"            '# If there is no transfer and the port is "listening", this will close the port so no files can be received.
    End If
    
    Close #DownloadingFile
    
    Winsock_Receive.Close                                               '# This closes the Receive port.
    Winsock_Send.Close                                                  '# This closes the Send port.
    
    FileBar.Value = 0                                                   '# This ensures that the filebar for sent files is reset to 0.
    FileBar2.Value = 0                                                  '# This ensures that the filebar for received files is reset to 0.
     
'# The following three lines of code set the caption on the LISTEN port to "Closing Port" and resets the sent and receive percentages.
    Lbl_Status2.Caption = "Closing Port: " & Winsock_Receive.LocalPort
    Lbl_Complete.Caption = "Complete: " & Int(100 / FileBar2.Max * FileBar2.Value) & " %"
    Lbl_Complete2.Caption = "Complete: " & Int(100 / FileBar2.Max * FileBar2.Value) & " %"
    
End Sub

Private Sub Command10_Click()
sckClient.Close
End Sub

Private Sub Command11_Click()
sckAccept.Close
End Sub

Private Sub Command12_Click()
MsgBox wData.State
End Sub


Private Sub Command2_Click()

    MsgBox "Thank You For Using SicklidSoft's SICKLIDSHARE!!", vbOKOnly, "GOODBYE!!!"   '# This is just a goobye textbox when the user wishes to close the program.
    End

End Sub

Private Sub Command3_Click()
    OpenUrl ("http://www.sicklidsoft.com")
End Sub

Private Sub Command6_Click()
    OpenUrl ("http://www.sicklidsoft.com")
End Sub

Private Sub Command7_Click()
    OpenUrl ("http://www.sicklidsoft.com")
End Sub

Private Sub Command8_Click()
MsgBox sckClient.State
End Sub

Private Sub Command9_Click()
MsgBox sckAccept.State
End Sub

Private Sub Dir1_Change()

       ' Obtain the file names from the new directory and populate
       ' the ListBox.
       Dim strCurrentPath As String
       Dim strFileName As String
       EraseTXTBoxes
       File12.Pattern = "*.*"
       File12.Pattern = "*.mp3"
       File12.Path = Dir1.Path
       If Right(Dir1.Path, 1) = "\" Then
           strCurrentPath = Dir1.Path
       Else
           strCurrentPath = Dir1.Path & "\"
       End If

       ' Clear the Listbox.

       ' Populate the Listbox with the file names.
       'strFileName = Dir(strCurrentPath)

       Do While strFileName <> ""
           List1.AddItem strFileName
           strFileName = Dir
       Loop

End Sub

      Private Sub List1_DblClick()
       ' Retrieve the text from the currently selected item.
       MsgBox List1.List(List1.ListIndex)
      End Sub

Private Sub Drive1_Change()
 On Error GoTo Error_Line 'Prevent Error (If there is)
 Dir1.Path = Drive1.List(Drive1.ListIndex)
 Exit Sub 'prevent use in the Error line (case it hasn't been called)

Error_Line:
 If Drive1.Drive = "a:" Then 'if drive is Fluppy
  MsgBox "No formatted Disket on drive A:"
  Drive1.Drive = "c:"
 Else 'if drive isn't fluppy (Cdrom, Etc)
  MsgBox "No Disk was found on drive " & Drive1.Drive
  Drive1.Drive = "c:"
 End If
End Sub

Private Sub Form_Load()
'# **************************************
'# THIS PORTION OF THE FORM_LOAD IS FOR THE MP3 PLAYER:
Dim X As Integer
Dim iLower As Integer
Dim iUpper As Integer

Drive1.Drive = "C:"
Dir1.Path = "C:\"
File12.Path = Dir1.Path
GenresTypes = Array("Blues", "Classic Rock", "Country", _
    "Dance", "Disco", "Funk", "Grunge", "Hip -Hop", _
    "Jazz", "Metal", "New Age", "Oldies", "Other", _
    "Pop", "R&b", "Rap", "Reggae", "Rock", "Techno", _
    "Industrial", "Alternative", "Ska", "Death Metal", _
    "Pranks", "Soundtrack", "Euro -Techno", "Ambient", _
    "Trip -Hop", "Vocal", "Jazz Funk", "Fusion", _
    "Trance", "Classical", "Instrumental")

iLower = LBound(GenresTypes)
iUpper = UBound(GenresTypes)
For X = iLower To iUpper
    Combo1.AddItem GenresTypes(X)
Next X

'# **************************************
'# THIS PORTION OF THE FORM_LOAD IS FOR THE PING PROG:
   With ListView1
      .View = lvwReport
      .ColumnHeaders.Add , , "Index"
      .ColumnHeaders.Add , , "IP Address"
      .ColumnHeaders.Add , , "Subnet Mask"
      .ColumnHeaders.Add , , "Broadcast Addr"
      .ColumnHeaders.Add , , "Reassembly"
      .ColumnHeaders.Add , , "unused1"
      .ColumnHeaders.Add , , "unused2"
   End With
'# **************************************

    Peeps = 2
    Timer3.Interval = 1

    Command8.Visible = False
    Command9.Visible = False
    Command10.Visible = False
    Command11.Visible = False

'Set the disconnect button to be on top of the connect button:
    cmdDisconnect.Left = cmdConnect2.Left
    cmdDisconnect.Top = cmdConnect2.Top
        
    Btn_Listen_Click                                                    '# At startup, this runs the Listen procedure which sets the port to LISTEN at port 420.
    Txt_CurrentIP.Text = Winsock_Receive.LocalIP                        '# This sets the txtlocal.Text value equal to the current ip to which the machine is attached. (File Sharing Value)
    txtLocal.Text = Winsock_Receive.LocalIP                             '# This sets the receivers local ip equal to the current ip to which the machine is attached. (Chat Value)
    txtNick.Text = UserName                                             '# This sets the User's Nickname for chatting equal to the UserName entered at logon.
End Sub

Private Sub mnuFileSharing_Click()
SicklidServer.TabSheet.Tab = 0
End Sub

Private Sub mnuPrivateChat_Click()
SicklidServer.TabSheet.Tab = 1
End Sub

Private Sub mnuMisc_Click()
SicklidServer.TabSheet.Tab = 2
End Sub

Private Sub mnuChatroom_Click()
SicklidServer.TabSheet.Tab = 3
End Sub

Private Sub mnuMP3Player_Click()
SicklidServer.TabSheet.Tab = 4
End Sub

Private Sub ShockwaveFlash1_OnReadyStateChange(NewState As Long)
    
    OpenUrl ("http://www.sicklidsoft.com")

End Sub



Private Sub Timer1_Timer()
'#  This sets the average Kilobytes per second transfer rate and displays it for user reference.
'    If wData.State = 8 And Lbl_Status = "Connection Closed..." Then
        
    Lbl_Averages.Caption = "Average: " & Format(DoneBytes / 1000, "###0.0") & " / KBps"
    DoneBytes = 0

'# The following describes what each winsock state means and returns that value to the lbl_status2.Caption.
  
        If Winsock_Receive.State = 0 Then
                Lbl_Status2.Caption = "Connection closed on port: " & Winsock_Receive.LocalPort
                Command1.Caption = "&Stop Transfer"
        ElseIf Winsock_Receive.State = 1 Then
                Lbl_Status2.Caption = "Connection in use on port: " & Winsock_Receive.LocalPort
        ElseIf Winsock_Receive.State = 2 Then
                Lbl_Status2.Caption = "Winsock listening on port: " & Winsock_Receive.LocalPort
                Command1.Caption = "&Close Port"
        ElseIf Winsock_Receive.State = 3 Then
                Lbl_Status2.Caption = "Connection pending on port: " & Winsock_Receive.LocalPort
        ElseIf Winsock_Receive.State = 4 Then
                Lbl_Status2.Caption = "Connection pending on port: " & Winsock_Receive.LocalPort
        ElseIf Winsock_Receive.State = 5 Then
                Lbl_Status2.Caption = "Host resolved on port: " & Winsock_Receive.LocalPort
        ElseIf Winsock_Receive.State = 6 Then
                Lbl_Status2.Caption = "Connecting on port: " & Winsock_Receive.LocalPort
        ElseIf Winsock_Receive.State = 7 Then
                Lbl_Status2.Caption = "Connected on port: " & Winsock_Receive.LocalPort
        ElseIf Winsock_Receive.State = 8 Then
                Lbl_Status2.Caption = "Peer closed connection on port: " & Winsock_Receive.LocalPort
        ElseIf Winsock_Receive.State = 9 Then
                Lbl_Status2.Caption = "Error on port: " & Winsock_Receive.LocalPort
        End If
        
'# The following describes what each winsock state means and returns that value to the lbl_status.Caption.
        
        
        If Winsock_Send.State = 0 Then
                Lbl_Status.Caption = "Connection closed on port: " & Winsock_Send.LocalPort
        ElseIf Winsock_Send.State = 1 Then
                Lbl_Status.Caption = "Connection in use on port: " & Winsock_Send.LocalPort
        ElseIf Winsock_Send.State = 2 Then
                Lbl_Status.Caption = "Winsock listening on port: " & Winsock_Send.LocalPort
        ElseIf Winsock_Send.State = 3 Then
                Lbl_Status.Caption = "Connection pending on port: " & Winsock_Send.LocalPort
        ElseIf Winsock_Send.State = 4 Then
                Lbl_Status.Caption = "Connection pending on port: " & Winsock_Send.LocalPort
        ElseIf Winsock_Send.State = 5 Then
                Lbl_Status.Caption = "Host resolved on port: " & Winsock_Send.LocalPort
        ElseIf Winsock_Send.State = 6 Then
                Lbl_Status.Caption = "Connecting on port: " & Winsock_Send.LocalPort
        ElseIf Winsock_Send.State = 7 Then
                Lbl_Status.Caption = "Connected on port: " & Winsock_Send.LocalPort
        ElseIf Winsock_Send.State = 8 Then
                Lbl_Status.Caption = "Peer closed connection on port: " & Winsock_Send.LocalPort
        ElseIf Winsock_Send.State = 9 Then
                Lbl_Status.Caption = "Error on port: " & Winsock_Send.LocalPort
        End If

End Sub

Private Sub Timer2_Timer()
'#  This sets the average Kilobytes per second transfer rate and displays it for user reference.
    
    Lbl_Averages2.Caption = "Average: " & Format(DoneRec / 1000, "###0.0") & " / KBps"
    DoneRec = 0
    
'# The following describes what each winsock state means and returns that value to the lbl_status2.Caption.
        
        If Winsock_Receive.State = 0 Then
                Lbl_Status2.Caption = "Connection closed on port: " & Winsock_Receive.LocalPort
        ElseIf Winsock_Receive.State = 1 Then
                Lbl_Status2.Caption = "Connection in use on port: " & Winsock_Receive.LocalPort
        ElseIf Winsock_Receive.State = 2 Then
                Lbl_Status2.Caption = "Winsock listening on port: " & Winsock_Receive.LocalPort
        ElseIf Winsock_Receive.State = 3 Then
                Lbl_Status2.Caption = "Connection pending on port: " & Winsock_Receive.LocalPort
        ElseIf Winsock_Receive.State = 4 Then
                Lbl_Status2.Caption = "Connection pending on port: " & Winsock_Receive.LocalPort
        ElseIf Winsock_Receive.State = 5 Then
                Lbl_Status2.Caption = "Host resolved on port: " & Winsock_Receive.LocalPort
        ElseIf Winsock_Receive.State = 6 Then
                Lbl_Status2.Caption = "Connecting on port: " & Winsock_Receive.LocalPort
        ElseIf Winsock_Receive.State = 7 Then
                Lbl_Status2.Caption = "Connected on port: " & Winsock_Receive.LocalPort
        ElseIf Winsock_Receive.State = 8 Then
                Lbl_Status2.Caption = "Peer closed connection on port: " & Winsock_Receive.LocalPort
        ElseIf Winsock_Receive.State = 9 Then
                Lbl_Status2.Caption = "Error on port: " & Winsock_Receive.LocalPort
        End If

'# The following describes what each winsock state means and returns that value to the lbl_status2.Caption.
   
        If Winsock_Send.State = 0 Then
                Lbl_Status.Caption = "Connection closed on port: " & Winsock_Send.LocalPort
        ElseIf Winsock_Send.State = 1 Then
                Lbl_Status.Caption = "Connection in use on port: " & Winsock_Send.LocalPort
        ElseIf Winsock_Send.State = 2 Then
                Lbl_Status.Caption = "Winsock listening on port: " & Winsock_Send.LocalPort
        ElseIf Winsock_Send.State = 3 Then
                Lbl_Status.Caption = "Connection pending on port: " & Winsock_Send.LocalPort
        ElseIf Winsock_Send.State = 4 Then
                Lbl_Status.Caption = "Connection pending on port: " & Winsock_Send.LocalPort
        ElseIf Winsock_Send.State = 5 Then
                Lbl_Status.Caption = "Host resolved on port: " & Winsock_Send.LocalPort
        ElseIf Winsock_Send.State = 6 Then
                Lbl_Status.Caption = "Connecting on port: " & Winsock_Send.LocalPort
        ElseIf Winsock_Send.State = 7 Then
                Lbl_Status.Caption = "Connected on port: " & Winsock_Send.LocalPort
        ElseIf Winsock_Send.State = 8 Then
                Lbl_Status.Caption = "Peer closed connection on port: " & Winsock_Send.LocalPort
        ElseIf Winsock_Send.State = 9 Then
                Lbl_Status.Caption = "Error on port: " & Winsock_Send.LocalPort
        End If
        
End Sub


Private Sub txtRemote_Change()
If txtRemote.Text <> "" Then
    cmdConnect2.Enabled = True
Else
    cmdConnect2.Enabled = False
End If
End Sub

Private Sub txtSend_Change()
If txtSend.Text <> "" Then
cmdSend.Enabled = True
Else
cmdSend.Enabled = False
End If
End Sub

Private Sub Winsock_Receive_Close()
    Close #DownloadingFile
    FileBar.Value = 0
    FileBar2.Value = 0
    Winsock_Receive.Close
    Winsock_Send.Close
    If Winsock_Receive.State = 0 Then
        Lbl_Status2.Caption = "Port: " & Winsock_Receive.LocalPort & " is closed!"
    End If
    
'************************************************************************'
'This part will reset the major variable to their default state.         '
'MUST have NextPart=False, otherwise all subsequent file transfer will be'
'EXACTLY one packet short. :)                                            '
'************************************************************************'
    
    DoneBytes = 0
    NextPart = False
    DoneRec = 0
    DownloadingFile = 0
    
    Btn_Listen_Click

End Sub

Private Sub Winsock_Receive_ConnectionRequest(ByVal requestID As Long)

    '# The following will close the Winsock if not already closed.accept the connections
    If Winsock_Receive.State <> sckClosed Then
        Winsock_Receive.Close
    End If
    Winsock_Receive.Accept requestID
    
    
    '# We use the close event to close the file afterwards
    
        
End Sub

Private Sub Winsock_Receive_DataArrival(ByVal bytesTotal As Long)
    Dim StrData As String
    Dim RecText2 As String
    Dim RecSize2 As String
    Dim RecName2 As String
    Dim lNewValue As Long
    Dim Info As String
    Dim Glob_FileName As String
    
    StrData = ""                                    '# This resets the variable to null.
                                                    
    RecText2 = ""                                   '# This resets the variable to null.
    RecSize2 = ""                                   '# This resets the variable to null.
    RecName2 = ""                                   '# This resets the variable to null.
    
    Winsock_Receive.GetData StrData, vbString       '# This gets the data sent by the sender.
    Winsock_Receive.GetData RecText2, vbString      '# This gets the data sent by the sender.
    Winsock_Receive.GetData RecSize2, vbString      '# This gets the data sent by the sender.
    Winsock_Receive.GetData RecName2, vbString      '# This gets the data sent by the sender.
    
'# This allows some file info to be sent before we receive the first package
    Info = Left(StrData, 8)
    If Info = "FILEINFO" Then                       '# If there is file information remaining to be read in the data, it loops:
        FileBar2.Max = GetField(StrData, 2)         '# This sets the Receiving Filebar equal to the second data segment sent.
        Glob_FileName = GetField(StrData, 3)        '# This sets the Receiving Filebar equal to the third data segment sent.
        RecText2 = GetField(StrData, 4) '# This sets the Receiving Filebar equal to the fourth data portion segment sent.
        RecName2 = GetField(StrData, 5) '# This sets the Receiving Filebar equal to the fifth data portion segment sent.
        RecSize2 = GetField(StrData, 6) '# This sets the Receiving Filebar equal to the sixth data portion segment sent.
        
        Txt_File2.Text = App.Path & "\" & Glob_FileName   '# This sets the Txt_File2.Text valye equal to the path location and the name of the file.
        DownloadingFile = FreeFile
        
        Open App.Path & "\Downloads\" & Glob_FileName For Binary Access Write As #DownloadingFile
    
        Lbl_FileSize2.Caption = "Filesize: " & RecSize2 & " kb" '# This allows the Filesize sent by the sender to be received and displayed by the receiving party.
        Lbl_FileName2.Caption = Glob_FileName                   '# This allows the Filename sent by the sender to be received and displayed by the receiving party.

        Exit Sub
    End If

    FileBar2.Value = FileBar2.Value + bytesTotal
    DoneRec = DoneRec + bytesTotal
       
    Lbl_Complete2.Caption = "Complete: " & Int(100 / FileBar2.Max * FileBar2.Value) & " %"
    
    Put #DownloadingFile, , StrData
        DoEvents
    
    Debug.Print Len(StrData)
    
End Sub

Public Function GetField(Field As String, FieldPos As Long) As String

'# This function extracts the elements from the data string sent by the sender.
'# So long as there is a byte left to read in the string, this function will loop.

Dim FieldCounter As Long
Dim IPPositionStart As Long
Dim IPPositionEnde As Long
Dim TempPosition As Long
Dim OpenedID As String
    
    TempPosition = 1
    
    For FieldCounter = 1 To FieldPos - 1 Step 1
        IPPositionStart = InStr(TempPosition, Field, "|", vbTextCompare)
        TempPosition = IPPositionStart + 1
    Next FieldCounter
    IPPositionStart = IPPositionStart + 1
    IPPositionEnde = InStr(IPPositionStart, Field, "|", vbTextCompare)
On Error Resume Next
    If IPPositionEnde >= IPPositionStart Then
        GetField = Mid(Field, IPPositionStart, IPPositionEnde - IPPositionStart)
    End If

End Function

Private Sub Winsock_Send_Connect()
        
        If Winsock_Send.State = 0 Then
                Lbl_Status.Caption = "Connection closed on port: " & Winsock_Receive.LocalPort
        ElseIf Winsock_Send.State = 1 Then
                Lbl_Status.Caption = "Connection in use on port: " & Winsock_Receive.LocalPort
        ElseIf Winsock_Send.State = 2 Then
                Lbl_Status.Caption = "Winsock listening on port: " & Winsock_Receive.LocalPort
        ElseIf Winsock_Send.State = 3 Then
                Lbl_Status.Caption = "Connection pending on port: " & Winsock_Receive.LocalPort
        ElseIf Winsock_Send.State = 4 Then
                Lbl_Status.Caption = "Connection pending on port: " & Winsock_Receive.LocalPort
        ElseIf Winsock_Send.State = 5 Then
                Lbl_Status.Caption = "Host resolved on port: " & Winsock_Receive.LocalPort
        ElseIf Winsock_Send.State = 6 Then
                Lbl_Status.Caption = "Connecting on port: " & Winsock_Receive.LocalPort
        ElseIf Winsock_Send.State = 7 Then
                Lbl_Status.Caption = "Connected on port: " & Winsock_Receive.LocalPort
        ElseIf Winsock_Send.State = 8 Then
                Lbl_Status.Caption = "Peer closed connection on port: " & Winsock_Receive.LocalPort
        ElseIf Winsock_Send.State = 9 Then
                Lbl_Status.Caption = "Error on port: " & Winsock_Receive.LocalPort
        End If
        
End Sub

Private Sub Winsock_Send_SendComplete()
    NextPart = True                             '# If this value were False, then no packets would be sent.
End Sub

Private Sub cmdConnect2_Click()

If frmConnect.IsAlpha(txtRemote) Then
    MsgBox "Enter a valid IP address, no letters!!", vbExclamation, "Error"
Else
    On Error Resume Next

    wData.Close                                 '# This closes the current connection for a new connection

    wData.Connect txtRemote.Text, txtPort.Text  '# This connects to the remote computer
    wData.LocalPort = txtPort.Text
    lblStatus.Caption = "Connecting..."         '# This message notifies the user that a connection is being attempted
    cmdConnect2.Visible = False                  '# This shows the Disconnect button since we are Connecting
    cmdDisconnect.Visible = True
    'MsgBox "Client " & sckClient.State
    'MsgBox "Accept " & sckAccept.State
    If err Then
        If err.Number = 40020 Then
            lblStatus.Caption = "Your Remote Host May Not Be Entered Correctly" '# If there was an error, inform the user:
            'txtRemote.sel
            'MsgBox "Client " & sckClient.State
            'MsgBox "Accept " & sckAccept.State
            txtRemote.SetFocus
            cmdConnect2.Visible = True           '# This resets the disconnect button to connect.
            cmdDisconnect.Visible = False
            NextPart = False
            Exit Sub
        End If
        
    lblStatus.Caption = "Error Number: " & err.Number & " " & err.Description    '# If there was an error, inform the user:

    cmdConnect2.Visible = True           '# This resets the disconnect button to connect.
    cmdDisconnect.Visible = False
    NextPart = False
    
    End If
End If
End Sub

Private Sub cmdDisconnect_Click()

'# This disconnects the current connection
    
    wData_Close
    wData.Close

End Sub

Private Sub cmdListen_Click()
    On Error Resume Next
    
Select Case cmdListen.Caption


    Case "&Listen"
        '# This sets the port to listen to:
        wData.LocalPort = txtPort.Text
    
        '# This starts the listening on the port for a connection:
        wData.Listen
        'MsgBox wData.State
        '# This notifies the user that the program is listening for a connection:
        lblStatus.Caption = "Listening..."
        cmdListen.Caption = "Close Port"
        cmdConnect2.Visible = False
        cmdDisconnect.Visible = False
        '# This notifies the user if there was an error and resets the connection:
        'If err Then
            'lblStatus.Caption = err.Description
            'Command1_Click
        'End If

    Case "Close Port"
        'MsgBox wData.State
        cmdListen.Caption = "&Listen"
        cmdConnect2.Visible = True

        wData2_Close
        wData.Close
        'MsgBox wData.State
End Select

End Sub


'****************************************************************************'
'*  Here begins the code for the chat program.  Currently this chat program *'
'*  only allows for one on one connections.  This will be remedied it later *'
'*  versions of this software.                                              *'
'****************************************************************************'

Private Sub cmdSend2_Click()
    
    Dim SendStr As String
    On Error Resume Next

'This puts the user's nickname and the message to be sent into a variable:
    SendStr = txtNick & ":" & vbTab & txtText.Text

'This sends the message to the recipient:
    wData.SendData SendStr

'This puts the current selection to the end:
    txtData.SelStart = Len(txtData.Text)

'This sets the color of the nickname to blue:
    txtData.SelColor = vbBlue

'This sets the text to be attached to our nickname:
    txtData.SelText = txtNick & ":" & vbTab

'This puts the current selection to the end:
    txtData.SelStart = Len(txtData.Text)

'This changes the nickname color back to black:
    txtData.SelColor = vbBlack

'This sets the text to the message we sent:
    txtData.SelText = txtText.Text & vbCrLf
    txtText.Text = ""
    
'This notifies the user if there has been an error:
    If err Then lblStatus.Caption = err.Description

End Sub

Private Sub txtText_GotFocus()

    cmdSend2.Default = True      '# This sets the Send button to activate when we press the Enter key. Nice huh?

End Sub

Private Sub txtText_Change()
If txtText.Text <> "" Then
cmdSend2.Enabled = True
Else
cmdSend2.Enabled = False
End If
End Sub

Private Sub txtText_LostFocus()

    cmdSend2.Default = False     '# This stops the Send button from being the default button.

End Sub

Private Sub wData2_Close()
'# This sub procedure closes the data and informs the user that the connection has closed:
    If wData.State = 7 And cmdListen.Caption = "&Listen" Then
        lblStatus.Caption = "Connection Closed By Server"
    Else
        lblStatus.Caption = "Connection Closed By Client"
    End If
    cmdConnect2.Visible = True           '# This resets the disconnect button to connect.
    cmdDisconnect.Visible = False
    NextPart = False
End Sub


Private Sub wData_Close()
'# This sub procedure closes the data and informs the user that the connection has closed:
    If wData.State = 8 And cmdListen.Caption = "&Listen" Then
        lblStatus.Caption = "Connection Closed By Server"
    Else
        lblStatus.Caption = "Connection Closed By Client"
    End If
        
    cmdListen.Caption = "&Listen"
    wData.Close
    cmdConnect2.Visible = True           '# This resets the disconnect button to connect.
    cmdDisconnect.Visible = False
    cmdListen.Visible = True
    NextPart = False
End Sub

Private Sub wData_Connect()
'# This informs the user that the connection has been established on the initiator's side:
    lblStatus.Caption = "Connected!"

    cmdConnect2.Visible = False          '# This resets the connect button to disconnect.
    cmdDisconnect.Visible = True
    cmdListen.Visible = False
End Sub


Private Sub wData_ConnectionRequest(ByVal requestID As Long)

'# This closes the current connection:
    wData.Close

'This accepts the connection request:
    wData.Accept requestID

'# This informs the user that the connection has been established on the recipients side:
    lblStatus.Caption = "Connection Accepted!"

'Show the disconnect button since we are connecting:
    If cmdListen.Caption = "&Listen" Then
        cmdConnect2.Visible = False
        cmdDisconnect.Visible = True
    End If
End Sub


Private Sub wData_DataArrival(ByVal bytesTotal As Long)
    
    Dim nData As String
    On Error Resume Next

'# This retrieves the data message being sent to us:
    wData.GetData nData

'# This sets the selection to the end:
    txtData.SelStart = Len(txtData.Text)

'# This sets the color of the sender's nickname to red:
    txtData.SelColor = vbRed

'# This puts the nickname of the other person into the message:
    txtData.SelText = Left(nData, InStr(1, nData, ":"))

'# This sets the selection to the end:
    txtData.SelStart = Len(txtData.Text)

'# This changes the color of the nickname back to black:
    txtData.SelColor = vbBlack

'# This sets the text in the messagebox to the message we received:
    txtData.SelText = Mid(nData, InStr(1, nData, ":") + 1) & vbCrLf

'# This notifies the user that there has been an error:
    If err Then lblStatus.Caption = err.Description

End Sub


Private Sub wData_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    lblStatus.Caption = Description     '# This notifies the user if there is an error and sets the error equal to the description of the error.

End Sub

'# **********************************************************
'# PORT SNIFFER PORTION OF PROGRAM



Private Sub Command5_Click()
If frmConnect.IsAlpha(Text1) Then
    MsgBox "Enter a valid IP address, no letters!!", vbExclamation, "Error"
Else


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
Text99.Text = "Starting..."
End If
Command4.Enabled = True
Paused = 0
If mnuLogtoFile.Checked = True Then
Open "PortLog.txt" For Output As #1
Close #1
CopyFileText
Open "PortLog.txt" For Output As #2
Write #2, FileText & vbCrLf & vbCrLf & "The remote IP at " & Winsock1.RemoteHost & " has the following open ports: " & vbCrLf & vbCrLf
Close #2
End If
mnuLogtoFile.Enabled = False
End If
End Sub

Private Sub Command4_Click()
Winsock1.Close
Text1.Enabled = True
Text2.Enabled = True
Text2.Text = PortToScan - 1
Text3.Enabled = True
Text5.Enabled = True
Command5.Enabled = True
Text99.Text = "Paused"
Timer3.Enabled = False
Command4.Enabled = True
Command5.Enabled = True
Paused = 1
End Sub

Private Sub mnuAbout_Click()
MsgBox "Hi " & UserName & "!!" & vbCrLf & vbCrLf & _
        "This program was created by Sicklid and Skubidu.  The File Sharing Tab will allow you " & vbCrLf & _
        "to do many things: Chat ONE-on-ONE; Host or Join a Chatroom holding as many people as " & vbCrLf & _
        "You wish; Ping an IP address;  Sniff a Remote IP for open ports and finally listen to " & vbCrLf & _
        "your favorite music while you do these things!  This was compiled using unique code & " & vbCrLf & _
        "code that was not working as well as code that was borrowed from others. " & vbCrLf & vbCrLf & _
        "Major bugs should be gone, but if you find any, please send an email to these emails: " & vbCrLf & _
        "Sicklid at sicklid_vision@yahoo.com or Skubidu at skubidu@yahoo.com" & vbCrLf & _
        "Note: If you see what might be your code in this program and you did not intend it to " & vbCrLf & _
        "be freeware, let us know and we'll give you the appropriate credit! We aren't pirates " & vbCrLf & _
        "so we like to give credit where credit is due." & vbCrLf & vbCrLf & _
        "IMPORTANT!!! YOU MUST HAVE A FOLDER NAMED 'DOWNLOADS' IN THE APPLICATION'S HOME DIR " & vbCrLf & _
        "IF YOU WANT TO SHARE FILES!" & vbCrLf & vbCrLf & _
        "Thank you for downloading this program and we hope it is useful to you!" & vbCrLf & _
        vbCrLf, vbOKOnly, "About SicklidShare version " & App.Major & "." & App.Minor & App.Revision


MsgBox "One more thing, " & UserName & "!" & vbCrLf & _
        "Some of the newer features of this software include a nifty little PortSniffer that allows " & vbCrLf & _
        "you to search any IP address and find the ports that are open on the remote machine.   The " & vbCrLf & _
        "Log to File Function in the Menubar will only Log the open ports for each IP address  that " & vbCrLf & _
        "you select.  PortLog.txt is the log file & is located in the application's home directory. " & vbCrLf & _
        "Another neat, new added feature is the Ping Utility that will allow you to ping  a  remote " & vbCrLf & _
        "address returning information on the IP address. You will also be able to list information " & vbCrLf & _
        "on the Host IP address.  The latest feature is the MP3 Player that will allow you to chat, " & vbCrLf & _
        "share files and ping the hell out of people while listening to your favorite MP3s!   Again " & vbCrLf & _
        "should you find any bugs or if you have any ideas or suggestions, please send us an email! "
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuLogToFile_Click()
If mnuLogtoFile.Checked = True Then
mnuLogtoFile.Checked = False
Else
mnuLogtoFile.Checked = True
End If
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
Text99.Text = PortToScan
TimerThing = TimerThing + 1

If TimerThing = 10 Then
 If TimeOut1 = 1 Then
  If Winsock1.State = sckConnected Then
  Winsock1.Close
  List1.AddItem Port1
   If mnuLogtoFile.Checked = True Then
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
  Text99.Text = "Done"
  Command4.Enabled = False
  Timer3.Enabled = False
  mnuLogtoFile.Enabled = True
  End If
 End If
End If

'

If TimerThing = 1 Then
 If TimeOut2 = 1 Then
  If Winsock2.State = sckConnected Then
  Winsock2.Close
  List1.AddItem Port2
   If mnuLogtoFile.Checked = True Then
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
  Text99.Text = "Done"
  Command4.Enabled = False
  Timer3.Enabled = False
  mnuLogtoFile.Enabled = True
  End If
 End If
End If

'

If TimerThing = 2 Then
 If TimeOut3 = 1 Then
  If Winsock3.State = sckConnected Then
  Winsock3.Close
  List1.AddItem Port3
   If mnuLogtoFile.Checked = True Then
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
  Text99.Text = "Done"
  Command4.Enabled = False
  Timer3.Enabled = False
  mnuLogtoFile.Enabled = True
  End If
 End If
End If

'

If TimerThing = 3 Then
 If TimeOut4 = 1 Then
  If Winsock4.State = sckConnected Then
  Winsock4.Close
  List1.AddItem Port4
   If mnuLogtoFile.Checked = True Then
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
  Text99.Text = "Done"
  Command4.Enabled = False
  Timer3.Enabled = False
  mnuLogtoFile.Enabled = True
  End If
 End If
End If

'

If TimerThing = 4 Then
 If TimeOut5 = 1 Then
  If Winsock5.State = sckConnected Then
  Winsock5.Close
  List1.AddItem Port5
   If mnuLogtoFile.Checked = True Then
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
  Text99.Text = "Done"
  Command4.Enabled = False
  Timer3.Enabled = False
  mnuLogtoFile.Enabled = True
  End If
 End If
End If

'

If TimerThing = 5 Then
 If TimeOut6 = 1 Then
  If Winsock6.State = sckConnected Then
  Winsock6.Close
  List1.AddItem Port6
   If mnuLogtoFile.Checked = True Then
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
  Text99.Text = "Done"
  Command4.Enabled = False
  Timer3.Enabled = False
  mnuLogtoFile.Enabled = True
  End If
 End If
End If

'

If TimerThing = 6 Then
 If TimeOut7 = 1 Then
  If Winsock7.State = sckConnected Then
  Winsock7.Close
  List1.AddItem Port7
   If mnuLogtoFile.Checked = True Then
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
  Text99.Text = "Done"
  Command4.Enabled = False
  Timer3.Enabled = False
  mnuLogtoFile.Enabled = True
  End If
 End If
End If

'

If TimerThing = 7 Then
 If TimeOut8 = 1 Then
  If Winsock8.State = sckConnected Then
  Winsock8.Close
  List1.AddItem Port8
   If mnuLogtoFile.Checked = True Then
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
  Text99.Text = "Done"
  Command4.Enabled = False
  Timer3.Enabled = False
  mnuLogtoFile.Enabled = True
  End If
 End If
End If

'

If TimerThing = 8 Then
 If TimeOut9 = 1 Then
  If Winsock9.State = sckConnected Then
  Winsock9.Close
  List1.AddItem Port9
   If mnuLogtoFile.Checked = True Then
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
  Text99.Text = "Done"
  Command4.Enabled = False
  Timer3.Enabled = False
  mnuLogtoFile.Enabled = True
  End If
 End If
End If

'

If TimerThing = 9 Then
 If TimeOut10 = 1 Then
  If Winsock10.State = sckConnected Then
  Winsock10.Close
  List1.AddItem Port10
   If mnuLogtoFile.Checked = True Then
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
  Text99.Text = "Done"
  Command4.Enabled = False
  Timer3.Enabled = False
  mnuLogtoFile.Enabled = True
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

'# ************************************************************
'# ********************CHAT ROOM CODE**************************
'# ************************************************************


Private Sub cmdConnect_Click()
Select Case cmdConnect.Caption

    Case "Connect"
        lblStatus2.Caption = "Connecting..."         '# This message notifies the user that a connection is being attempted
        If err Then lblStatus2.Caption = err.Description    '# If there was an error, inform the user:

        frmConnect.Show

    Case "Cancel"
        cmdConnect.Caption = "Connect"
        txtIncoming.Text = ""
        cmdHost.Enabled = True
        sckClient.Close

    Case "Disconnect"
        lblStatus2.Caption = "Disconnected..."
        sckClient.SendData "X" & Me.Tag
        txtIncoming.Text = ""
        Me.Caption = "Not Connected"
        cmdHost.Enabled = True
        cmdConnect.Caption = "Connect"
        cmdClrScreen.Enabled = False
        cmdSend.Enabled = False
        lstPeople.Clear
        
        If err Then lblStatus2.Caption = err.Description    '# If there was an error, inform the user:
        
        If sckClient.State = 0 Then
            txtIncoming.Text = ""
            lblStatus2.Caption = "Not Connected..."
            Me.Caption = "Not Connected"
            cmdHost.Enabled = True
            cmdConnect.Caption = "Connect"
            cmdClrScreen.Enabled = False
            cmdSend.Enabled = False
            lstPeople.Clear
            Exit Sub
        End If
        
        If sckClient.State = 8 Then
            MsgBox "Peer Closed Connection", vbOKOnly, "CONNECTION CLOSED"
            Me.Caption = "Not Connected"
            cmdHost.Enabled = True
            cmdConnect.Caption = "Connect"
            cmdClrScreen.Enabled = False
            cmdSend.Enabled = False
            lstPeople.Clear
            txtIncoming.Text = ""
            Exit Sub
        End If
    
End Select
End Sub

Private Sub cmdHost_Click()
        
Select Case cmdHost.Caption
    
    Case "Host a Room"
        txtIncoming.Text = ""
        frmHost.Show
    
    Case "Close Room"
        If MsgBox("Are you sure?  This will disconnect anyone who is in your room!", vbYesNo, "Are you sure?") = vbYes Then



If Len(Me.Tag) < 7 Then tabs = vbTab & vbTab Else: tabs = vbTab
If sckClient.State > 0 Then
    sckClient.SendData "[" & Me.Tag & ":" & tabs & "CLOSING" & vbCrLf
    txtIncoming.Text = ""
    'txtIncoming.Text = txtIncoming.Text & Me.Tag & ":" & tabs & "CLOSING THE ROOM.  HIT DISCONNECT" & vbCrLf
    sckClient.SendData "X" & Me.Tag
    
    sendtoall "[" & Me.Tag & ":" & tabs & "Sorry, Gotta Close This Room." & vbCrLf & "Click the Disconnect Button" & vbCrLf

Else


            'If sckAccept.State = 0 Then
            '    Exit Sub
            'End If
            
End If
sckClient_Close
txtSend.Text = ""
            
            If sckAccept.State = 2 Then
                sckAccept.Close
            End If
            txtSend.Text = ""
            sckAccept.Close
            
            cmdHost.Caption = "Host a Room"
            cmdSend.Enabled = False
            cmdConnect.Enabled = True
            cmdClrScreen.Enabled = False
            Me.Caption = "Not Connected"
            lstPeople.Clear
            txtIncoming.Text = ""
            End If

End Select
End Sub


Private Sub cmdSend_Click()
If sckClient.State = 0 And sckAccept.State = 0 Then
    txtSend.Text = ""
    Exit Sub
End If

'This notifies the user if there has been an error:
    If err Then lblStatus2.Caption = err.Description

If Len(Me.Tag) < 7 Then tabs = vbTab & vbTab Else: tabs = vbTab
If sckClient.State > 0 Then
    sckClient.SendData "[" & Me.Tag & ":" & tabs & txtSend.Text & vbCrLf
Else
    'If sckAccept.State = 2 Then txtIncoming.SelColor = vbRed
    'If sckAccept.State = 0 Then txtIncoming.SelColor = vbBlue
    'txtIncoming.SelStart = Len(txtSend.Text)
    'If sckAccept.State = 2 Then txtIncoming.SelColor = vbRed
    'If sckAccept.State = 0 Then txtIncoming.SelColor = vbBlue
    'txtIncoming.SelText = Me.Tag & ":" & vbTab
    'txtIncoming.SelStart = Len(txtSend.Text)
    'txtIncoming.SelColor = vbBlack

    txtIncoming.Text = txtIncoming.Text & Me.Tag & ":" & tabs & txtSend.Text & vbCrLf
    sendtoall "[" & Me.Tag & ":" & tabs & txtSend.Text & vbCrLf
    
End If
txtSend.Text = ""
txtSend.SetFocus
End Sub


Private Sub Form_Unload(Cancel As Integer)

End
End Sub

Private Sub Close_Conn()
    sendtoall "[" & Me.Tag & ":" & tabs & "HAS CLOSED THE CHATROOM" & vbCrLf
    sckAccept.Close
    sckClient.Close
    For n = 0 To sckServer.Count - 1
    If sckServer(n).State <> 0 Then
        sckServer(n).Close
        Exit Sub
    End If
Next n

End Sub

Private Sub sckAccept_ConnectionRequest(ByVal requestID As Long)



If Peeps <= frmHost.txtMaxPeeps.Text Then
Peeps = Peeps + 1
For n = 0 To sckServer.Count - 1
    If sckServer(n).State = 0 Then
        sckServer(n).Accept requestID
        Exit Sub
    ElseIf sckServer(n).State = 8 Then
        sckServer(n).Close
        sckServer(n).Accept requestID
        Exit Sub
    End If
Next n
'they were all busy if we get here
X = sckServer.Count
Load sckServer(X)
sckServer(X).Accept requestID

End If

End Sub

Private Sub sckClient_Connect()
'# This informs the user that the connection has been established on the initiator's side:
    lblStatus2.Caption = "Connected!"
    lblChatroom.Caption = "Connected to the: " & sckClient.Tag & " Chat Room & Your host is: " & Mid(dat$, 2)

    
    
    
txtIncoming.Text = txtIncoming.Text & vbCrLf & "Status:" & vbTab & vbTab & "...Getting Info From Server..." & vbCrLf
cmdConnect.Caption = "Disconnect"
cmdHost.Enabled = False
cmdClrScreen.Enabled = True
sckClient.SendData "_" & Me.Tag
End Sub

Private Sub sckClient_DataArrival(ByVal bytesTotal As Long)
'This notifies the user if there has been an error:
    If err Then lblStatus2.Caption = err.Description
    
sckClient.GetData dat$
commnd$ = Left(dat$, 1)
If commnd$ = "P" Then              'Person coming in
    lstPeople.AddItem Mid(dat$, 2)
    sckClient.SendData "."
    txtIncoming.Text = txtIncoming.Text & "Status:" & vbTab & vbTab & "--- " & Mid(dat$, 2) & " has entered the room ---" & vbCrLf
ElseIf commnd$ = "N" Then          'Someone left
    For n = 0 To lstPeople.ListCount - 1
        If lstPeople.List(n) = Mid(dat$, 2) Then
            lstPeople.RemoveItem n
            sckClient.SendData "."
            txtIncoming.Text = txtIncoming.Text & "Status:" & vbTab & vbTab & "--- " & Mid(dat$, 2) & " has left the room ---" & vbCrLf
            Exit Sub
        End If
    Next n
ElseIf commnd$ = "R" Then          'Room name
    sckClient.Tag = Mid(dat$, 2)
    sckClient.SendData "."
ElseIf commnd$ = "H" Then          'Your host
    lblChatroom.Caption = "Connected to the: " & sckClient.Tag & " Chat Room & Your host is: " & Mid(dat$, 2)
    Me.Caption = "Connected to the: " & sckClient.Tag & " Chat Room & Your host is: " & Mid(dat$, 2)
    sckClient.SendData "."
ElseIf commnd$ = "K" Then          'Ready for data
    cmdSend.Enabled = True
    txtIncoming.Text = txtIncoming.Text & "Status:" & vbTab & vbTab & "...Connected." & vbCrLf
    sckClient.SendData "."
ElseIf commnd$ = "[" Then          'Someone said...
    txtIncoming.SelColor = vbRed
    txtIncoming.Text = txtIncoming.Text & Mid(dat$, 2)
    txtIncoming.SelColor = vbBlack
ElseIf commnd$ = "*" Then
    sckClient.Close
    txtIncoming.Text = ""
End If
End Sub

Private Sub sckClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
If Number = 10053 Then
    MsgBox "This Host Has Maxed Out Their Connections!" & vbCrLf & "Try Again Later!", vbCritical & vbOKOnly, "SERVER MAXED OUT!"
'# This sub procedure closes the data and informs the user that the connection has closed:
    n = 0
    sckServer(n).Close
    lblChatroom.Caption = "Not Connected or Hosting a Chatroom"
    lblStatus2.Caption = "Connection Closed"
    sckClient.Close
    sckAccept.Close
'    lblStatus2.Caption = "You have ..."
    txtIncoming.Text = ""
    Me.Caption = "Not Connected"
    cmdHost.Enabled = True
    cmdConnect.Caption = "Connect"
    cmdClrScreen.Enabled = False
    cmdSend.Enabled = False
    lstPeople.Clear
    lblChatroom.Caption = "Not Connected or Hosting a Chatroom"
    
    disconnect
    Exit Sub
End If

MsgBox "Error number " & Number & ":" & vbCrLf & vbCrLf & Description
'# This sub procedure closes the data and informs the user that the connection has closed:
    n = 0
    sckServer(n).Close
    lblChatroom.Caption = "Not Connected or Hosting a Chatroom"
    lblStatus2.Caption = "Connection Closed"
    sckClient.Close
    sckAccept.Close
'    lblStatus2.Caption = "You have ..."
    txtIncoming.Text = ""
    Me.Caption = "Not Connected"
    cmdHost.Enabled = True
    cmdConnect.Caption = "Connect"
    cmdClrScreen.Enabled = False
    cmdSend.Enabled = False
    lstPeople.Clear
    lblChatroom.Caption = "Not Connected or Hosting a Chatroom"

disconnect

End Sub

Public Sub disconnect()
    Me.Caption = "Not Connected"
    cmdHost.Enabled = True
    cmdConnect.Caption = "Connect"
    cmdClrScreen.Enabled = False
    cmdSend.Enabled = False
    lstPeople.Clear
    txtIncoming.Text = ""

If sckClient.State = 0 Then
    Exit Sub
Else
    sckClient.Close
End If

End Sub

Public Sub sendtoall(dats As String)
For n = 0 To sckServer.Count - 1
    If sckServer(n).State = 7 Then
        sckServer(n).SendData dats
    End If
Next n
End Sub

Private Sub sckServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)

'This notifies the user if there has been an error:
    If err Then lblStatus2.Caption = err.Description
    
If sckAccept.State = 2 Then
sckServer(Index).GetData da$
lastdata$ = da$
commnd$ = Left(da$, 1)
Debug.Print da$
If commnd$ = "." Then 'okgo
    Pause 0.5
    okgo = True
ElseIf commnd$ = "_" Then      'Person Loggin in
    Pause 0.5
    sckServer(n).Tag = "N" & Mid(da$, 2)
    txtIncoming.Text = txtIncoming.Text & "Status:" & vbTab & vbTab & "--- " & Mid(da$, 2) & " has entered the room ---" & vbCrLf
    lstPeople.AddItem Mid(da$, 2)
    sendtoall "P" & Mid(da$, 2)
    If waitforok(n, 6) = False Then
    Pause 0.5
222     sendtoall "N" & Mid(da$, 2)
        For X = 0 To lstPeople.ListCount - 1
            If lstPeople.List(X) = Mid(da$, 2) Then lstPeople.RemoveItem X
        Next X
        txtIncoming.Text = txtIncoming.Text & "Status:" & vbTab & vbTab & "--- " & Mid(da$, 2) & " has left the room ---" & vbCrLf
        Exit Sub
    End If
    
'# Removed txtIncoming.text line here

    sckServer(Index).SendData "R" & sckAccept.Tag
    If waitforok(n, 6) = False Then GoTo 222
    sckServer(Index).SendData "H" & Me.Tag
    If waitforok(n, 6) = False Then GoTo 222
    For n = 0 To lstPeople.ListCount - 2
        If lstPeople.List(n) = Mid(da$, 2) Then GoTo 11
        sckServer(Index).SendData "P" & lstPeople.List(n)
        If waitforok(n, 6) = False Then GoTo 222
11  Next n
    sckServer(Index).SendData "K"
ElseIf commnd$ = "[" Then  'Someones talkin
    txtIncoming.SelColor = vbBlue
    txtIncoming.Text = txtIncoming.Text & Mid(da$, 2)
    sendtoall da$
    txtIncoming.SelColor = vbBlack

ElseIf commnd$ = "X" Then
    sckServer(Index).SendData "*"
    Pause 0.5
    sckServer(Index).Close
    GoTo 222
End If
End If

If sckAccept.State <> 2 Then
    
If Len(Me.Tag) < 7 Then tabs = vbTab & vbTab Else: tabs = vbTab
If sckClient.State > 0 Then
    sckClient.SendData "[" & Me.Tag & ":" & tabs & "CLOSING" & vbCrLf
Else
    'txtIncoming.Text = txtIncoming.Text & Me.Tag & ":" & tabs & "CLOSING THE ROOM.  HIT DISCONNECT" & vbCrLf
    sendtoall "[" & "" & tabs & ""
    'sendtoall "[" & Me.Tag & " closed this room: " & tabs & vbCrLf & "Don't keep typing text, silly!" & vbCrLf & "The room is CLOSED" & vbCrLf & "Click the Disconnect Button" & vbCrLf
End If
txtSend.Text = ""
End If

End Sub

Public Function waitforok(server As Integer, Timeout As Long) As Boolean
Dim StartTime As Long
StartTime = Timer
X = StartTime + Timeout
Do Until Timer > X
DoEvents
If okgo Then
    okgo = False
    waitforok = True
    Exit Function
End If
Loop
sckServer(server - 1).Close
waitforok = False
End Function

Private Sub Timer5_Timer()

    If sckClient.State = 8 Then
        'MsgBox "You're OUTTA HERE!", vbOKOnly, "GOODBYE"
        disconnect
    End If
    If sckClient.State = 0 And cmdConnect.Caption = "Disconnect" Then
        Me.Caption = "Not Connected"
        sckClient.SendData "X" & Me.Tag
        cmdHost.Enabled = True
        cmdConnect.Caption = "Connect"
        cmdClrScreen.Enabled = False
        cmdSend.Enabled = False
        lstPeople.Clear
        txtIncoming.Text = ""
        sckClient_Close
    End If

End Sub

Private Sub txtIncoming_Change()
txtIncoming.SelStart = Len(txtIncoming.Text)
End Sub

Private Sub sckClient_Close()
'# This sub procedure closes the data and informs the user that the connection has closed:
    n = 0
    sckServer(n).Close
    lblChatroom.Caption = "Not Connected or Hosting a Chatroom"
    lblStatus2.Caption = "Connection Closed"
    sckClient.Close
    sckAccept.Close
'    lblStatus2.Caption = "You have ..."
    txtIncoming.Text = ""
    Me.Caption = "Not Connected"
    cmdHost.Enabled = True
    cmdConnect.Caption = "Connect"
    cmdClrScreen.Enabled = False
    cmdSend.Enabled = False
    lstPeople.Clear
    lblChatroom.Caption = "Not Connected or Hosting a Chatroom"

End Sub

Private Sub sckClient_ConnectionRequest(ByVal requestID As Long)

'# This closes the current connection:
    sckClient.Close

'This accepts the connection request:
    sckClient.Accept requestID

'# This informs the user that the connection has been established on the recipients side:
    lblStatus2.Caption = "Connection Accepted!"

End Sub


'# *********************************************
'# *************IP PINGER PROGRAM***************
'# *********************************************

Public Function GetInetStrFromPtr(ByVal Address As Long) As String
  
   GetInetStrFromPtr = GetStrFromPtrA(inet_ntoa(Address))

End Function

Public Function GetStrFromPtrA(ByVal lpszA As Long) As String

   GetStrFromPtrA = String$(lstrlenA(ByVal lpszA), 0)
   Call lstrcpyA(ByVal GetStrFromPtrA, ByVal lpszA)
   
End Function

Private Sub Command14_Click()
   
   Dim ECHO As ICMP_ECHO_REPLY
   Dim pos As Long
   Dim success As Long
   
   If SocketsInitialize2() Then

     'ping the ip passing the address, text
     'to send, and the ECHO structure.
      success = Ping((Text7.Text), (Text6.Text), ECHO)
      
     'display the results
      Text4(0).Text = GetStatusCode(success)
      If success = 0 Then
      Text4(1).Text = Text7.Text
      Else
        Text4(1).Text = ""
      End If
      'Text4(1).Text = ECHO.Address & " Dec Add"
      Text4(2).Text = ECHO.RoundTripTime & " ms"
      Text4(3).Text = ECHO.DataSize & " bytes"
      
      If Left$(ECHO.Data, 1) <> Chr$(0) Then
         pos = InStr(ECHO.Data, Chr$(0))
         Text4(4).Text = Left$(ECHO.Data, pos - 1)
      End If
   
      Text4(5).Text = ECHO.DataPointer
      
      SocketsCleanup
      
   Else
   
        MsgBox "Windows Sockets for 32 bit Windows " & _
               "environments is not successfully responding."
   
   End If
   
End Sub

Private Sub Command13_Click()

   Dim IPAddrRow As MIB_IPADDRROW
   Dim buff() As Byte
   Dim cbRequired As Long
   Dim nStructSize As Long
   Dim nRows As Long
   Dim cnt As Long
   Dim itmx As ListItem
   
   Call GetIpAddrTable(ByVal 0&, cbRequired, 1)
   
   If cbRequired > 0 Then
    
      ReDim buff(0 To cbRequired - 1) As Byte
      
      If GetIpAddrTable(buff(0), cbRequired, 1) = ERROR_SUCCESS Then
      
        'saves using LenB in the CopyMemory calls below
         nStructSize = LenB(IPAddrRow)
        'first 4 bytes is a long indicating the
        'number of entries in the table
         CopyMemory nRows, buff(0), 4
      
         For cnt = 1 To nRows
         
           'moving past the four bytes obtained
           'above, get one chunk of data and cast
           'into an IPAddrRow type
            CopyMemory IPAddrRow, buff(4 + (cnt - 1) * nStructSize), nStructSize
            
           'pass the results to the listview
            With IPAddrRow
                Set itmx = ListView1.ListItems.Add(, , GetInetStrFromPtr(.dwIndex))
                itmx.SubItems(1) = GetInetStrFromPtr(.dwAddr)
                itmx.SubItems(2) = GetInetStrFromPtr(.dwMask)
                itmx.SubItems(3) = GetInetStrFromPtr(.dwBCastAddr)
                itmx.SubItems(4) = GetInetStrFromPtr(.dwReasmSize)
                itmx.SubItems(5) = (.unused1)
                itmx.SubItems(6) = (.unused2)
            End With
          Next cnt
      End If
   End If

End Sub

'# *****************************************************************
'# **************MP3 PLAYER CODE
'# *****************************************************************

'# This portion of the code deciphers the tag information
'# The most popular tag encryption appears to be ID3,
'# so we'll base our code on that format. This standard stores the
'# tag information in the last 128 bytes of the file. Table A shows
'# the exact length and order of each piece of data.
'# To read this information, you first open the MP3 file and grab the
'# last 128 bytes. With ID3, the first three slots hold the string TAG
'# if the file actually contains information. If the file does contain
'# tag information, we'll store the last 128 bytes in a custom Type variable.
'# After that, our code can cycle through the MP3 file, extracting information
'# as it goes. Listing A shows the code that extracts this information as well
'# as creates several important variables.
'# Notice that the code has to handle the genre character a little differently.
'# That's because ID3 stores this data as a single ASCII character. To match up
'# the actual number with its corresponding combobox description, the procedure
'# converts the ASCII to a number, and then looks up that number in the combobox.

Private Sub File12_Click()
Dim Temp As String
Dim myfile As String
On Error Resume Next
EraseTXTBoxes


If Right(Dir1.Path, 1) = "\" Then
    FileName = Dir1.Path & File12.FileName
Else
    FileName = Dir1.Path & "\" & File12.FileName
    myfile = FileName
End If
    
Open FileName For Binary As #1
With CurrentTag
    Get #1, FileLen(FileName) - 127, .Tag
    If Not .Tag = "TAG" Then
        ' Label8.Caption = "No tag"
        Close #1
        Exit Sub
    End If
    Get #1, , .Songname
    Get #1, , .artist
    Get #1, , .album
    Get #1, , .year
    Get #1, , .comment
    Get #1, , .genre
    Close #1

    txtTitle = RTrim(.Songname)
    txtArtist = RTrim(.artist)
    txtAlbum = RTrim(.album)
    txtYear = RTrim(.year)
    txtComment = RTrim(.comment)
    
    Temp = RTrim(.genre)
    txtGenreCode = Asc(Temp)
    Combo1.ListIndex = CInt(txtGenreCode) - 1
End With
End Sub

'# This portion of the code writes the tag information back to the MP3 File
'# To write the tag information back to the file, you use a similar technique.
'# Again, our code will take advantage of the Open command. This time, however,
'# it will use the Put command. Listing B shows the procedure we attached to the
'# command button cmdWriteTag's Click() event.

Private Sub cmdWriteTag_Click()
If FileOpen Then
    MsgBox "You can't save to an open file", _
        vbCritical, "MP3 Tag Save Error"
    Exit Sub
Else:
    If txtTitle.Text = "" Then
        MsgBox "You have to pick a valid MP3 File First!", vbCritical, "MP3 Tag Save Error"
        Exit Sub
    End If
End If
    
If Right(Dir1.Path, 1) = "\" Then
    FileName = Dir1.Path & File12.FileName
Else
    FileName = Dir1.Path & "\" & File12.FileName
End If

With CurrentTag
    .Tag = "TAG"
    .Songname = txtTitle
    .artist = txtArtist
    .album = txtAlbum
    .year = txtYear
    .comment = txtComment
    .genre = Chr(Combo1.ListIndex + 1)
    
    Open FileName For Binary Access Write As #1
        Seek #1, FileLen(FileName) - 127
        Put #1, , .Tag
        Put #1, , .Songname
        Put #1, , .artist
        Put #1, , .album
        Put #1, , .year
        Put #1, , .comment
        Put #1, , .genre
    Close #1
End With
End Sub

'# This portion of the code will play the MP3 File.
'# Up to this point, we have yet to explain how to play an MP3 file.
'# Believe it or not, playing the file is actually the easiest part
'# of this process. To do so, you set the Media Player's FileName
'# property to the file you want to play. That's all there is to it!
'# From this point, you use the media player's built-in buttons to
'# control playing. Listing C shows the code we created for our application.
'# Notice that after setting the FileName property, we chose not to
'# have it start playing automatically, so we set the AutoStart property to False.
'# The code in Listing A stored the currently selected file location in the
'# FileName variable

Private Sub cmdOpen_Click()
With MediaPlayer1
    If Not FileOpen Then
       MediaPlayer1.FileName = FileName
        .AutoStart = False
        cmdOpen.Caption = "Close MP3"
        If txtTitle.Text = "" Then cmdOpen.Caption = "Open MP3"
    Else
        cmdOpen.Caption = "Open MP3"
        FileOpen = False
    End If
End With

End Sub

'# This portion of code will determine the length of the MP3 file.
'# Now that the code has opened the file, we can determine its duration.
'# Whenever a file's open state changes, Media Player triggers an
'# OpenStateChange() event. We'll use this event to calculate the song's duration.
'# Listing D shows this event. As you can see, this event also sets the
'# FileOpen variable.

Private Sub MediaPlayer1_OpenStateChange(ByVal _
    OldState As Long, ByVal NewState As Long)

Min = MediaPlayer1.Duration \ 60
Sec = MediaPlayer1.Duration - (Min * 60)
lblTotalTime = "Total Time: " & Format(Min, "0#") _
    & ":" & Format(Sec, "0#") 'format time to 00:00
    
FileOpen = CBool(NewState)
MediaPlayer1.AutoStart = False
End Sub

'# This portion of code will keep track of play time.
'# As one of our final touchups to the application, we can provide the current
'# running time. Earlier, we placed two labels and a timer control on the form
'# for just such a purpose. Listing E contains the code for this feature.

Private Sub Timer6_Timer()
Min = MediaPlayer1.CurrentPosition \ 60
Sec = MediaPlayer1.CurrentPosition - (Min * 60)
If Min > 0 Or Sec > 0 Then
    lblElapsedTime = "Elapsed Time: " & Format(Min, "0#") _
        & ":" & Format(Sec, "0#")
Else
    lblElapsedTime = "Elapsed Time: 00:00"
End If
End Sub

'# This portion of code will clear the information.
'# As our last step, we need to add a procedure that clears the tag
'# information from the textboxes, and code to fill the GenresTypes array.
'# Listing F contains both procedures. To save space, we shortened the full
'# list of genre types. For an expanded list, check out this month's download file.

Private Sub EraseTXTBoxes()
'Label8.Caption = ""
txtTitle = ""
txtArtist = ""
txtAlbum = ""
txtYear = ""
txtComment = ""
txtGenreCode = ""
Combo1.ListIndex = -1
End Sub



'# ONE LAST NOTE OF IMPORT!
'# When you use a For...Next loop to iterate through an array,
'# you may be inclined to hard code the starting and ending counter values.
'# For instance, in the loop in Listing F, we could have used
'#
'# For x = 0 To 33
'#
'# However, because you may want to expand the genre list at a later time,
'# it's best to use the LBound() and UBound() functions to delimit the counter's
'# boundaries. That way, no matter how many times you add items to the array,
'# you won't need to modify the For...Next loop at all.
'#
'# Also, keep in mind that we didn't need to subtract 1 from the UBound() value
'# because the function returns the array's largest available subscript not the
'# number of items in the array.
