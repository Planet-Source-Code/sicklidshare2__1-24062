Attribute VB_Name = "XMessagerIPbas"
Private Const IP_SUCCESS As Long = 0
Private Const IP_STATUS_BASE As Long = 11000
Private Const IP_BUF_TOO_SMALL As Long = (11000 + 1)
Private Const IP_DEST_NET_UNREACHABLE As Long = (11000 + 2)
Private Const IP_DEST_HOST_UNREACHABLE As Long = (11000 + 3)
Private Const IP_DEST_PROT_UNREACHABLE As Long = (11000 + 4)
Private Const IP_DEST_PORT_UNREACHABLE As Long = (11000 + 5)
Private Const IP_NO_RESOURCES As Long = (11000 + 6)
Private Const IP_BAD_OPTION As Long = (11000 + 7)
Private Const IP_HW_ERROR As Long = (11000 + 8)
Private Const IP_PACKET_TOO_BIG As Long = (11000 + 9)
Private Const IP_REQ_TIMED_OUT As Long = (11000 + 10)
Private Const IP_BAD_REQ As Long = (11000 + 11)
Private Const IP_BAD_ROUTE As Long = (11000 + 12)
Private Const IP_TTL_EXPIRED_TRANSIT As Long = (11000 + 13)
Private Const IP_TTL_EXPIRED_REASSEM As Long = (11000 + 14)
Private Const IP_PARAM_PROBLEM As Long = (11000 + 15)
Private Const IP_SOURCE_QUENCH As Long = (11000 + 16)
Private Const IP_OPTION_TOO_BIG As Long = (11000 + 17)
Private Const IP_BAD_DESTINATION As Long = (11000 + 18)
Private Const IP_ADDR_DELETED As Long = (11000 + 19)
Private Const IP_SPEC_MTU_CHANGE As Long = (11000 + 20)
Private Const IP_MTU_CHANGE As Long = (11000 + 21)
Private Const IP_UNLOAD As Long = (11000 + 22)
Private Const IP_ADDR_ADDED As Long = (11000 + 23)
Private Const IP_GENERAL_FAILURE As Long = (11000 + 50)
Private Const MAX_IP_STATUS As Long = (11000 + 50)
Private Const IP_PENDING As Long = (11000 + 255)
Private Const PING_TIMEOUT As Long = 500
Private Const INADDR_NONE As Long = &HFFFFFFFF
Private Const MAX_WSADescription As Long = 256
Private Const MAX_WSASYSStatus As Long = 128


Public Const WS_VERSION_REQD = &H101
Public Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD = 1
Public Const SOCKET_ERROR = -1
Public Const WSADescription_Len = 256
Public Const WSASYS_Status_Len = 128



Public Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLength As Integer
    hAddrList As Long
End Type

Public Type WSADATA
    wversion As Integer
    wHighVersion As Integer
    szDescription(0 To WSADescription_Len) As Byte
    szSystemStatus(0 To WSASYS_Status_Len) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpszVendorInfo As Long
End Type

Private Type ICMP_OPTIONS
    Ttl             As Byte
    Tos             As Byte
    Flags           As Byte
    OptionsSize     As Byte
    OptionsData     As Long
End Type

Public Type ICMP_ECHO_REPLY
    Address         As Long
    status          As Long
    RoundTripTime   As Long
    DataSize        As Long 'formerly integer
   'Reserved        As Integer
    DataPointer     As Long
    Options         As ICMP_OPTIONS
    Data            As String * 250
End Type

Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long

Private Declare Function IcmpCloseHandle Lib "icmp.dll" _
   (ByVal IcmpHandle As Long) As Long
   
Private Declare Function IcmpSendEcho Lib "icmp.dll" _
   (ByVal IcmpHandle As Long, _
    ByVal DestinationAddress As Long, _
    ByVal RequestData As String, _
    ByVal RequestSize As Long, _
    ByVal RequestOptions As Long, _
    ReplyBuffer As ICMP_ECHO_REPLY, _
    ByVal ReplySize As Long, _
    ByVal Timeout As Long) As Long

Public Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Public Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired&, lpWSAData As WSADATA) As Long
Public Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Public Declare Function gethostname Lib "WSOCK32.DLL" (ByVal hostname$, ByVal HostLen As Long) As Long
Public Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal hostname$) As Long
Private Declare Function gethostbyname2 Lib "WSOCK32.DLL" (ByVal szHost As String) As Long
Public Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource&, ByVal cbCopy&)

Private Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (xDest As Any, _
   xSource As Any, _
   ByVal nbytes As Long)
   
Private Declare Function inet_conv Lib "WSOCK32.DLL" _
    (ByVal s As String) As Integer

Private Declare Function inet_addr Lib "WSOCK32.DLL" _
   (ByVal s As String) As Long

' *********************************************************
' * Code was not originally written by Master Yoda.       *
' * I didn't spend a lot of time to fixed up this code so *
' * feel free to do so if nessesary.                      *
' *********************************************************
Function hibyte(ByVal wParam As Integer)
    hibyte = wParam \ &H100 And &HFF&
End Function


Function lobyte(ByVal wParam As Integer)
    lobyte = wParam And &HFF&
End Function


Sub SocketsInitialize()
    Dim WSAD As WSADATA
    Dim iReturn As Integer
    Dim sLowByte As String, sHighByte As String, sMsg As String
    iReturn = WSAStartup(WS_VERSION_REQD, WSAD)


    If iReturn <> 0 Then
        MsgBox "Error: " & "Winsock.dll Error."
        Exit Sub
    End If
    If lobyte(WSAD.wversion) < WS_VERSION_MAJOR Or (lobyte(WSAD.wversion) = _
        WS_VERSION_MAJOR And hibyte(WSAD.wversion) < WS_VERSION_MINOR) Then
        sHighByte = Trim$(str$(hibyte(WSAD.wversion)))
        sLowByte = Trim$(str$(lobyte(WSAD.wversion)))
        sMsg = "Windows Sockets version " & sLowByte & "." & sHighByte
        MsgBox sMsg
        Exit Sub
    End If

End Sub

Public Function CurrentIP(ReturnExternalIP As Boolean)
    Dim hostname As String * 256
    Dim hostent_addr As Long
    Dim host As HOSTENT
    Dim hostip_addr As Long
    Dim temp_ip_address() As Byte
    Dim I As Integer
    Dim ip_address As String
    Dim IP As String

    If gethostname(hostname, 256) = SOCKET_ERROR Then
        Exit Function
    Else
        hostname = Trim$(hostname)
    End If
    hostent_addr = gethostbyname(hostname)
    '# Someone mentioned this might be the correct line above...testing.
    'hostent_addr = gethostbyname2(hostname)


    If hostent_addr = 0 Then
        MsgBox "Winsock.dll error."
        Exit Function
    End If
    RtlMoveMemory host, hostent_addr, LenB(host)
    RtlMoveMemory hostip_addr, host.hAddrList, 4
   
    Do
        ReDim temp_ip_address(1 To host.hLength)
        RtlMoveMemory temp_ip_address(1), hostip_addr, host.hLength


        For I = 1 To host.hLength
            ip_address = ip_address & temp_ip_address(I) & "."
        Next
        ip_address = Mid$(ip_address, 1, Len(ip_address) - 1)
    
    ' Return Both LAN and External IP Fix
    ' Master Yoda 30-05-2000
    ' ##########################################
    '   HERE'S THE PROBLEM!!!
        'TheIP = TheIP + ip_address
    ' ##########################################
    '   HERE'S THE FIX!!!
        Internal = TheIP        ' Send ONLY the External IP to the CurrentIP Function
        external = ip_address   ' Send the External IP to the  function parameter External
        TheIP = ip_address      ' Send LAN IP to the function para Internal
        
        ' You don't really need to return parameters,
        ' it just allows you to get both IPs :)
    ' ##########################################

        ip_address = ""
        host.hAddrList = host.hAddrList + LenB(host.hAddrList)
        RtlMoveMemory hostip_addr, host.hAddrList, 4
    Loop While (hostip_addr <> 0)
    
    
If ReturnExternalIP = True Then
    CurrentIP = external
Else
    CurrentIP = Internal
End If
End Function

Sub SocketsCleanup()
    Dim lReturn As Long
    lReturn = WSACleanup()


    If lReturn <> 0 Then
        MsgBox "Socket Error " & Trim$(str$(lReturn)) & " occurred In Cleanup "
        Exit Sub
    End If
End Sub

Sub OpenUrl(url As String)

    Shell ("Explorer " & url$), vbNormalNoFocus
End Sub


Public Function SocketsInitialize2() As Boolean

   Dim WSAD As WSADATA
   
   SocketsInitialize2 = WSAStartup(WS_VERSION_REQD, WSAD) = IP_SUCCESS
    
End Function


    

Public Function GetStatusCode(status As Long) As String

   Dim msg As String
   
   Select Case status
      Case IP_SUCCESS:               msg = "ip success"
      Case INADDR_NONE:              msg = "inet_addr: bad IP format"
      Case IP_BUF_TOO_SMALL:         msg = "ip buf too_small"
      Case IP_DEST_NET_UNREACHABLE:  msg = "ip dest net unreachable"
      Case IP_DEST_HOST_UNREACHABLE: msg = "ip dest host unreachable"
      Case IP_DEST_PROT_UNREACHABLE: msg = "ip dest prot unreachable"
      Case IP_DEST_PORT_UNREACHABLE: msg = "ip dest port unreachable"
      Case IP_NO_RESOURCES:          msg = "ip no resources"
      Case IP_BAD_OPTION:            msg = "ip bad option"
      Case IP_HW_ERROR:              msg = "ip hw_error"
      Case IP_PACKET_TOO_BIG:        msg = "ip packet too_big"
      Case IP_REQ_TIMED_OUT:         msg = "ip req timed out"
      Case IP_BAD_REQ:               msg = "ip bad req"
      Case IP_BAD_ROUTE:             msg = "ip bad route"
      Case IP_TTL_EXPIRED_TRANSIT:   msg = "ip ttl expired transit"
      Case IP_TTL_EXPIRED_REASSEM:   msg = "ip ttl expired reassem"
      Case IP_PARAM_PROBLEM:         msg = "ip param_problem"
      Case IP_SOURCE_QUENCH:         msg = "ip source quench"
      Case IP_OPTION_TOO_BIG:        msg = "ip option too_big"
      Case IP_BAD_DESTINATION:       msg = "ip bad destination"
      Case IP_ADDR_DELETED:          msg = "ip addr deleted"
      Case IP_SPEC_MTU_CHANGE:       msg = "ip spec mtu change"
      Case IP_MTU_CHANGE:            msg = "ip mtu_change"
      Case IP_UNLOAD:                msg = "ip unload"
      Case IP_ADDR_ADDED:            msg = "ip addr added"
      Case IP_GENERAL_FAILURE:       msg = "ip general failure"
      Case IP_PENDING:               msg = "ip pending"
      Case PING_TIMEOUT:             msg = "ping timeout"
      Case Else:                     msg = "unknown  msg returned"
   End Select
   
   GetStatusCode = CStr(status) & "   [ " & msg & " ]"
   
End Function

Public Function Ping(sAddress As String, sDataToSend As String, ECHO As ICMP_ECHO_REPLY) As Long

  'If Ping succeeds :
  '.RoundTripTime = time in ms for the ping to complete,
  '.Data is the data returned (NULL terminated)
  '.Address is the Ip address that actually replied
  '.DataSize is the size of the string in .Data
  '.Status will be 0
  '
  'If Ping fails .Status will be the error code
   
   Dim hPort As Long
   Dim dwAddress As Long
   
  'convert the address into a long representation
   dwAddress = inet_addr(sAddress)

  'if a valid address..
   If dwAddress <> INADDR_NONE Then
   
     'open a port
      hPort = IcmpCreateFile()
      
     'and if successful,
      If hPort Then
      
        'ping it.
         Call IcmpSendEcho(hPort, _
                           dwAddress, _
                           sDataToSend, _
                           Len(sDataToSend), _
                           0, _
                           ECHO, _
                           Len(ECHO), _
                           PING_TIMEOUT)

        'return the status as ping succes and close
         Ping = ECHO.status
         Call IcmpCloseHandle(hPort)
      
      End If
      
   Else:
        'the address format was probably invalid
         Ping = INADDR_NONE
         
   End If
  
End Function


