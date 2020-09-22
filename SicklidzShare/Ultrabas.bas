Attribute VB_Name = "modUltraBas"
Public tabs As String
Public X As Long
Public dat As String
Public commnd As String
Public da As String
Public tdata As String
Public mnd As String
Public m_strSearch As String


'The Ultra Bas File (Declarations)
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Sub ReleaseCapture Lib "user32" ()
Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Declare Function findwindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function Getwindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function getwindowtext Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Integer
Public Declare Function sendmessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long
Public Declare Function setwindowtext Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal Cx As Long, ByVal Cy As Long, ByVal wFlags As Long) As Long
Public Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function SendDriverMessage Lib "winmm.dll" (ByVal hDriver As Long, ByVal Message As Long, ByVal lParam1 As Long, ByVal lParam2 As Long) As Long
Public Declare Function SetDoubleClickTime Lib "user32" (ByVal wCount As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Sub FatalAppExit Lib "kernel32" Alias "FatalAppExitA" (ByVal uAction As Long, ByVal lpMessageText As String)
Public Declare Sub FatalExit Lib "kernel32" (ByVal code As Long)
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function HideCaret Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
'Public Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long

      Public Type POINTAPI
         X As Long
         Y As Long
      End Type

Public Slippery As Boolean
Public Closee As Boolean
Public Const EM_UNDO = &HC7
Public tdOptions$, Options$
Public PercentComplete, Dire, n%
Public Const MAX_PATH = 206
Private Const GWW_HINSTANCE = (-6)
Private Const GWW_ID = (-12)
Private Const GWL_STYLE = (-16)
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_QUIT = &H12
Public Const WM_PASTE = &H302
Public Const WM_MENUSELECT = &H11F
Public Const WM_SETTEXT = &HC
Public Const WM_ENABLE = &HA
Public Const WM_GETFONT = &H31
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_INITMENU = &H116
Public Const WM_KEYLAST = &H108

Private Const GW_CHILD = 5
Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDLAST = 1
Private Const GW_HWNDNEXT = 2
Private Const GW_HWNDPREV = 3
Private Const GW_MAX = 5
Private Const GW_OWNER = 4
Public Const clsAol4combo = "_AOL_Combobox"
Public Const clsAol4child = "AOL Child"
Public Const clsMsgbox = "#32770"
Public Const clsMsgboxMsg = "Static"

Public Function MousePos(X As Boolean) As Long
    Dim CursorPos As POINTAPI
    GetCursorPos CursorPos
    If X = True Then
        MousePos = CursorPos.X
    Else
        MousePos = CursorPos.Y
    End If
End Function


Public Function SetMouse(X, Y)
    SetCursorPos X, Y
End Function
Public Sub Slip()
    SlipAmmount = 1.1 'Set this value as 2 to make it less slippery
    Dim X As Long
    Dim Y As Long
    Dim Pt As POINTAPI
    Do
        GetCursorPos Pt
        Dim Tx, Ty
        Tx = (Pt.X - X) / SlipAmmount
        Ty = (Pt.Y - Y) / SlipAmmount
        If Pt.X >= Screen.Width / Screen.TwipsPerPixelX - Screen.TwipsPerPixelX Then Tx = 0 - Tx
        If Pt.X <= 0 + Screen.TwipsPerPixelX Then Tx = 0 - Tx
        If Pt.Y >= Screen.Height / Screen.TwipsPerPixelY - Screen.TwipsPerPixelY Then Ty = 0 - Ty
        If Pt.Y <= 0 + Screen.TwipsPerPixelY Then Ty = 0 - Ty
        SetCursorPos Pt.X + Tx, Pt.Y + Ty
        X = Pt.X
        Y = Pt.Y
There:
        Sleep 2
        DoEvents
        Loop Until Closee = True


    End Sub

Public Sub Stayontop(Frm As Form)
Dim X As Integer
With Frm
t% = .Top
l% = .Left
h% = .Height
w% = .Width
X% = SetWindowPos(Frm.hwnd, -1, 0, 0, 0, 0, 0)
.Top = t%
.Left = l%
.Height = h%
.Width = w%
End With
End Sub
Public Sub Pause(seconds As Long)
Dim X As Long
X = Timer + seconds
Do Until Timer > X
DoEvents
Loop
End Sub
Public Sub GetoffTop(Frm As Form)
Dim X As Integer
With Frm
t% = .Top
l% = .Left
h% = .Height
w% = .Width
X% = SetWindowPos(Frm.hwnd, -2, 0, 0, 0, 0, 0)
.Top = t%
.Left = l%
.Height = h%
.Width = w%
End With
End Sub
Public Sub TypeEffect(masg As String, txt As TextBox)
Char% = 1
    HideCaret txt.hwnd
    Do While Not Char% > Len(masg)
        NoFreeze% = DoEvents
        ch$ = Mid(masg, Char%, 1)
        txt.Text = txt.Text & ch$
        Char% = Char% + 1
    Loop
End Sub
Sub Playsound(snd As String)
Dim X As Integer
X% = sndPlaySound(snd, 0)
End Sub
Public Function GetChildArray(ChildArray() As Long, ByVal hParent As Long) As Long

Dim hChild As Long         ' Handle of the Child
Dim Temp(256) As Long   ' Temporary array

   Dim I As Integer        ' Index counter
   If hParent = 0 Then
      GoTo Return_False
   End If
   hChild = Getwindow(hParent, GW_CHILD)
   While hChild
      Temp(I) = hChild
      hChild = Getwindow(hChild, GW_HWNDNEXT)
      I = I + 1
   Wend

   If I = 0 Then GoTo Return_False
      ReDim ChildArray(I - 1) As Long

   For I = 0 To I - 1
      ChildArray(I) = Temp(I)
   Next I
   GetChildArray = I
   
   Exit Function
Return_False:
   GetChildArray = 0
   Exit Function
End Function


Public Function GetChildCount(ByVal hwnd As Long) As Long
      
      Dim hChild As Long      ' Handle of the Child
   Dim I As Integer        ' Index counter
   
   If hwnd = 0 Then
      GoTo Return_False
   End If
   hChild = Getwindow(hwnd, GW_CHILD)
   While hChild
      hChild = Getwindow(hChild, GW_HWNDNEXT)
      I = I + 1
   Wend
   GetChildCount = I
   
   Exit Function
Return_False:
   GetChildCount = 0
   Exit Function
End Function


Public Function GetFileName(ByVal hwnd As Long) As String

      Dim sModuleFileName As String * 100
      Dim hInstance, ret As Long
  
   hInstance = GetWindowWord(hwnd, GWW_HINSTANCE)
   ret = GetModuleFileName(hInstance, sModuleFileName, 100)
   GetFileName = Left(sModuleFileName, ret)
End Function

Public Function GetParentWindow(ByVal hwnd As Long) As Long
   GetParentWindow = GetParent(hwnd)
End Function



Function GetTopChild(ByVal hwnd As Long) As Long
   GetTopChild = Getwindow(hwnd, GW_CHILD)
End Function



Public Function hWndOver() As Long
      Dim pt32 As POINTAPI
      Dim ptx As Long
      Dim pty As Long
   
      Call GetCursorPos(pt32)               ' Get cursor position
      ptx = pt32.X
      pty = pt32.Y
      hWndOver = WindowFromPointXY(ptx, pty)    ' Get window cursor is over
End Function




Public Function GetWindowClassName(ByVal hWindow As Long) As String
      Dim sClassName As String * 100
      Dim ret As Long
   ret = GetClassName(hWindow, sClassName, 100)
   GetWindowClassName = Trim$(Left(sClassName, ret))
End Function
               



Public Function findchildbyclass(ByVal hParent As Long, ByVal sClassName As String, Optional ByVal nIndex) As Long
   Dim hChild As Long
   Dim I As Integer

   If IsMissing(nIndex) Then
      nIndex = 1
   ElseIf nIndex < 1 Then
      Exit Function
   End If
   hChild = Getwindow(hParent, GW_CHILD)
   While I < nIndex And hChild
      If GetWindowClassName(hChild) = sClassName Then
         I = I + 1   ' Increase counter
      End If
      
      If I < nIndex Then
         hChild = Getwindow(hChild, GW_HWNDNEXT)
      End If
   Wend
   findchildbyclass = hChild
   Exit Function
End Function


Public Function findchildbytitle(ByVal hParent As Long, ByVal sWindowText As String, Optional ByVal nIndex) As Long
   Dim hChild As Long
   Dim I As Integer              ' Index counte

   If IsMissing(nIndex) Then
      nIndex = 1
   ElseIf nIndex < 1 Then
      Exit Function
   End If
   hChild = Getwindow(hParent, GW_CHILD)
   While I < nIndex And hChild
      If WindowText(hChild) = sWindowText Then
         I = I + 1
      End If
      If I < nIndex Then
         hChild = Getwindow(hChild, GW_HWNDNEXT)
      End If
   Wend
   findchildbytitle = hChild
   Exit Function
End Function

Public Function WindowText(ByVal hwnd As Long) As String
   ' Declare Return Variable
   Dim ret As Long
   Dim sWindowText As String * 100
   ' Get the Window's text
   ret = getwindowtext(hwnd, sWindowText, 100)
   WindowText = Left(sWindowText, ret)
End Function

Public Function HasClassName(ByVal hwnd As Long) As Boolean
   HasClassName = Len(GetWindowClassName(hwnd))
End Function

Public Function FileExist(fan As String) As Boolean
On Error GoTo err
    X = FileLen(fan)
    If X > 0 Then FileExist = True
    Exit Function
err:
    FileExist = False
End Function
Public Function Scrambletext(thetext)
'sees if there's a space in the text to be scrambled,
'if found space, continues, if not, adds it
findlastspace = Mid(thetext, Len(thetext), 1)
If Not findlastspace = " " Then
thetext = thetext & " "
Else
thetext = thetext
End If
'Scrambles the text
For scrambling = 1 To Len(thetext)
thechar$ = Mid(thetext, scrambling, 1)
Char$ = Char$ & thechar$
If thechar$ = " " Then
'takes out " " space from the text left of the space
chars$ = Mid(Char$, 1, Len(Char$) - 1)
'gets first character
firstchar$ = Mid(chars$, 1, 1)
'gets last character (if not, makes first character only)
On Error GoTo cityz
lastchar$ = Mid(chars$, Len(chars$), 1)
'finds what is inbetween the last and first character
midchar$ = Mid(chars$, 2, Len(chars$) - 2)
'reverses the text found in between the last and first
'character
For SpeedBack = Len(midchar$) To 1 Step -1
backchar$ = backchar$ & Mid$(midchar$, SpeedBack, 1)
Next SpeedBack
GoTo sniffe
'adds the scrambled text to the full scrambled element
cityz:
scrambled$ = scrambled$ & firstchar$ & " "
GoTo sniffs
sniffe:
scrambled$ = scrambled$ & lastchar$ & firstchar$ & backchar$ & " "
'clears character and reversed buffers
sniffs:
Char$ = ""
backchar$ = ""
End If
Next scrambling
'Makes function return value the scrambled text
Scrambletext = scrambled$
Exit Function
End Function
Public Function StringParse(str As String, whichpart As Integer, Delimeter As String)
whered = InStr(1, str, Delimeter)
firstpart$ = Mid(str, 1, whered - 1)
If whichpart = 1 Then StringParse = firstpart$: Exit Function
Number = 1
Do
DoEvents
Number = Number + 1
Char$ = Mid(str, whered + 1 + Count, 1)
If Char$ = Delimeter Then
    If Number = whichpart Then
        StringParse = lastString$
        Exit Function
    Else
        lastString$ = ""
        GoTo 1
    End If
End If
lastString$ = lastString$ & Char$
Count = Count + 1
1
Loop
End Function

Public Sub ClickWindow(hwndd As Long, rightclickd As Boolean)
If rightclickd Then
sendmessage hwndd, WM_RBUTTONDOWN, 1, 0&
sendmessage hwndd, WM_RBUTTONUP, 1, 0&
Else
sendmessage hwndd, WM_LBUTTONDOWN, 1, 0&
sendmessage hwndd, WM_LBUTTONUP, 1, 0&
End If
End Sub

Public Function DesktophWnd() As Long
DesktophWnd = GetDesktopWindow
End Function
